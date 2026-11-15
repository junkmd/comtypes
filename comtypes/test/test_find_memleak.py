"""Test for find_memleak module, including MallocSpy implementation."""

import gc
import unittest
from collections.abc import Callable
from typing import Any

from comtypes import hresult
from comtypes.malloc import CoGetMalloc, _CoRegisterMallocSpy, _CoRevokeMallocSpy
from comtypes.test.find_memleak import MallocSpy


def _revoke_existing_spy():
    """Try to revoke any existing malloc spy, ignoring errors."""
    try:
        _CoRevokeMallocSpy()
    except Exception:
        pass


class TestMallocSpy(unittest.TestCase):
    """Test the MallocSpy COMObject implementation."""

    def setUp(self):
        # Ensure no spy is registered before each test
        _revoke_existing_spy()
        gc.collect()
        gc.collect()
        gc.collect()

    def tearDown(self):
        # Clean up after each test
        _revoke_existing_spy()
        gc.collect()
        gc.collect()
        gc.collect()

    def _run_with_spy(self, test_func: Callable[[MallocSpy], Any]) -> None:
        """Helper to run a test with a registered spy, ensuring cleanup."""
        spy = MallocSpy()
        hr = _CoRegisterMallocSpy(spy)
        self.assertEqual(hr, hresult.S_OK)

        try:
            spy.reset()
            test_func(spy)
        finally:
            _CoRevokeMallocSpy()

    def test_mallocspy_tracks_allocations(self):
        """Test that MallocSpy tracks COM allocations."""

        def _test(spy: MallocSpy) -> None:
            # Do some COM allocations
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            ptr2 = malloc.Alloc(200)
            malloc.Free(ptr1)
            malloc.Free(ptr2)

            gc.collect()
            gc.collect()
            gc.collect()

            # Check that allocations were tracked
            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            # We should have at least 2 allocations and 2 frees
            self.assertGreaterEqual(alloc_count, 2)
            self.assertGreaterEqual(free_count, 2)

            # Net should be 0 since we freed everything
            net = spy.get_net_allocations()
            self.assertEqual(net, 0)

        self._run_with_spy(_test)

    def test_mallocspy_detects_leak(self):
        """Test that MallocSpy can detect a memory leak."""

        def _test(spy: MallocSpy) -> None:
            # Do some COM allocations but don't free them (leak)
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            ptr2 = malloc.Alloc(200)
            # Intentionally not freeing - this is a leak

            gc.collect()
            gc.collect()
            gc.collect()

            # Check that allocations were tracked
            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            # We should have at least 2 allocations and 0 frees
            self.assertGreaterEqual(alloc_count, 2)
            self.assertEqual(free_count, 0)

            # Net should be positive (leak detected)
            net = spy.get_net_allocations()
            self.assertGreaterEqual(net, 2)

            # Free the memory so we can revoke the spy
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            gc.collect()
            gc.collect()
            gc.collect()

        self._run_with_spy(_test)

    def test_mallocspy_reset(self):
        """Test that MallocSpy reset works correctly."""

        def _test(spy: MallocSpy) -> None:
            # Do some allocations
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            malloc.Free(ptr1)

            gc.collect()
            gc.collect()
            gc.collect()

            # Verify we have some counts
            self.assertGreater(spy.get_allocated_count(), 0)

            # Reset
            spy.reset()

            # Counts should be zero
            self.assertEqual(spy.get_allocated_count(), 0)
            self.assertEqual(spy.get_freed_count(), 0)
            self.assertEqual(spy.get_net_allocations(), 0)

        self._run_with_spy(_test)

    def test_mallocspy_realloc(self):
        """Test that MallocSpy tracks Realloc calls (PreRealloc/PostRealloc)."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            # Allocate initial memory
            ptr1 = malloc.Alloc(100)
            # Realloc to larger size
            ptr2 = malloc.Realloc(ptr1, 200)
            # Realloc to smaller size
            ptr3 = malloc.Realloc(ptr2, 50)
            malloc.Free(ptr3)

            gc.collect()
            gc.collect()
            gc.collect()

            # Check that realloc was tracked (counted as allocation)
            # Realloc involves PreRealloc and PostRealloc calls
            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            # We should have allocations (initial + reallocs) and frees
            self.assertGreaterEqual(alloc_count, 3)  # Alloc + 2 Reallocs
            self.assertGreaterEqual(free_count, 1)

        self._run_with_spy(_test)

    def test_mallocspy_getsize(self):
        """Test that MallocSpy tracks GetSize calls (PreGetSize/PostGetSize)."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            # Call GetSize - this should trigger PreGetSize/PostGetSize
            size = malloc.GetSize(ptr1)
            self.assertEqual(size, 100)
            malloc.Free(ptr1)

            gc.collect()
            gc.collect()
            gc.collect()

            # GetSize doesn't change allocation count but should be called
            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            self.assertGreaterEqual(alloc_count, 1)
            self.assertGreaterEqual(free_count, 1)

        self._run_with_spy(_test)

    def test_mallocspy_didalloc(self):
        """Test that MallocSpy tracks DidAlloc calls (PreDidAlloc/PostDidAlloc)."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            # Call DidAlloc - this should trigger PreDidAlloc/PostDidAlloc
            result = malloc.DidAlloc(ptr1)
            self.assertEqual(result, 1)
            # Also test with a non-allocated pointer
            result2 = malloc.DidAlloc(0x12345678)  # Random invalid pointer
            self.assertEqual(result2, 0)
            malloc.Free(ptr1)

            gc.collect()
            gc.collect()
            gc.collect()

            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            self.assertGreaterEqual(alloc_count, 1)
            self.assertGreaterEqual(free_count, 1)

        self._run_with_spy(_test)

    def test_mallocspy_heapminimize(self):
        """Test that MallocSpy tracks HeapMinimize calls (PreHeapMinimize/PostHeapMinimize)."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            malloc.Free(ptr1)
            # Call HeapMinimize - this should trigger PreHeapMinimize/PostHeapMinimize
            malloc.HeapMinimize()

            gc.collect()
            gc.collect()
            gc.collect()

            alloc_count = spy.get_allocated_count()
            free_count = spy.get_freed_count()

            self.assertGreaterEqual(alloc_count, 1)
            self.assertGreaterEqual(free_count, 1)

        self._run_with_spy(_test)


if __name__ == "__main__":
    unittest.main()
