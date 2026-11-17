import gc
import unittest
from collections.abc import Callable
from typing import Any

from comtypes import hresult
from comtypes.malloc import CoGetMalloc, _CoRegisterMallocSpy, _CoRevokeMallocSpy
from comtypes.test.find_memleak import MallocSpy, find_memleak, get_working_set_size


def _revoke_existing_spy() -> None:
    """Try to revoke any existing malloc spy, ignoring errors."""
    try:
        _CoRevokeMallocSpy()
    except OSError:
        pass


class CleanUpMixin:
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


class Test_MallocSpy(CleanUpMixin, unittest.TestCase):
    """Test the MallocSpy COMObject implementation."""

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

    def test_tracks_allocations(self):
        """Test that MallocSpy tracks COM allocations in bytes."""

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

            # Check that allocations were tracked in bytes
            allocated = spy.get_allocated_bytes()
            freed = spy.get_freed_bytes()

            # We should have at least 300 bytes allocated and 300 bytes freed
            self.assertEqual(allocated, 300)
            self.assertEqual(freed, 300)

            # Net should be 0 since we freed everything
            net = spy.get_net_bytes()
            self.assertEqual(net, 0)

        self._run_with_spy(_test)

    def test_detects_leak(self):
        """Test that MallocSpy can detect a memory leak in bytes."""

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
            allocated = spy.get_allocated_bytes()
            freed = spy.get_freed_bytes()

            # We should have 300 bytes allocated and 0 bytes freed
            self.assertEqual(allocated, 300)
            self.assertEqual(freed, 0)

            # Net should be positive (leak detected)
            net = spy.get_net_bytes()
            self.assertEqual(net, 300)

            # Free the memory so we can revoke the spy
            malloc.Free(ptr1)
            malloc.Free(ptr2)
            gc.collect()
            gc.collect()
            gc.collect()

        self._run_with_spy(_test)

    def test_reset(self):
        """Test that MallocSpy reset works correctly."""

        def _test(spy: MallocSpy) -> None:
            # Do some allocations
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            malloc.Free(ptr1)

            gc.collect()
            gc.collect()
            gc.collect()

            # Verify we have some bytes tracked
            self.assertGreater(spy.get_allocated_bytes(), 0)

            # Reset
            spy.reset()

            # All values should be zero
            self.assertEqual(spy.get_allocated_bytes(), 0)
            self.assertEqual(spy.get_freed_bytes(), 0)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_reset_forgets_outstanding_blocks(self):
        """Test that reset() forgets outstanding blocks; freeing them later doesn't affect counts."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            ptr2 = malloc.Alloc(200)

            gc.collect()
            gc.collect()
            gc.collect()

            # Reset before freeing
            spy.reset()

            # Now free the blocks
            malloc.Free(ptr1)
            malloc.Free(ptr2)

            gc.collect()
            gc.collect()
            gc.collect()

            # Freed bytes should not increase (blocks were forgotten)
            self.assertEqual(spy.get_freed_bytes(), 0)
            # Net should not go negative
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_realloc(self):
        """Test that MallocSpy tracks Realloc calls as free+alloc in bytes."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            # Allocate initial memory (100 bytes)
            ptr1 = malloc.Alloc(100)
            # Realloc to larger size (200 bytes) - frees 100, allocates 200
            ptr2 = malloc.Realloc(ptr1, 200)
            # Realloc to smaller size (50 bytes) - frees 200, allocates 50
            ptr3 = malloc.Realloc(ptr2, 50)
            malloc.Free(ptr3)

            gc.collect()
            gc.collect()
            gc.collect()

            # Allocations: 100 + 200 + 50 = 350
            # Frees: 100 + 200 + 50 = 350
            # Net: 0
            self.assertEqual(spy.get_allocated_bytes(), 350)
            self.assertEqual(spy.get_freed_bytes(), 350)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_getsize(self):
        """Test that MallocSpy tracks GetSize calls without changing byte counts."""

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

            # GetSize doesn't change allocation bytes
            self.assertEqual(spy.get_allocated_bytes(), 100)
            self.assertEqual(spy.get_freed_bytes(), 100)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_didalloc(self):
        """Test that MallocSpy tracks DidAlloc calls without changing byte counts."""

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

            # DidAlloc doesn't change allocation bytes
            self.assertEqual(spy.get_allocated_bytes(), 100)
            self.assertEqual(spy.get_freed_bytes(), 100)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)

    def test_heapminimize(self):
        """Test that MallocSpy tracks HeapMinimize calls without changing byte counts."""

        def _test(spy: MallocSpy) -> None:
            malloc = CoGetMalloc()
            ptr1 = malloc.Alloc(100)
            malloc.Free(ptr1)
            # Call HeapMinimize - this should trigger PreHeapMinimize/PostHeapMinimize
            malloc.HeapMinimize()

            gc.collect()
            gc.collect()
            gc.collect()

            # HeapMinimize doesn't change allocation bytes
            self.assertEqual(spy.get_allocated_bytes(), 100)
            self.assertEqual(spy.get_freed_bytes(), 100)
            self.assertEqual(spy.get_net_bytes(), 0)

        self._run_with_spy(_test)


class Test_find_memleak(CleanUpMixin, unittest.TestCase):
    def test_find_memleak_working_set_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        # `find_memleak()` takes its baseline with one `get_working_set_size()`
        # call and its final reading with another.
        # The first `get_working_set_size()` call in a process incurs one-time
        # costs (e.g. first-time code paths faulting in pages) that land
        # between those two readings, so a cold probe can make an otherwise
        # leak-free function look like it grew by a page.
        # Calling `get_working_set_size()` once here pays that cost beforehand.
        #
        # Note that this only warms up the probe itself. The warmup of
        # `no_leak()` is already done by `find_memleak()`.
        #
        # This is a workaround that makes the measurement stable in practice.
        # It does not address the fundamental limitation of Working Set
        # based measurement (see https://github.com/enthought/comtypes/issues/971):
        # the Working Set is the amount of resident pages, not the amount of
        # unreleased memory, so a stable result does not prove the absence of
        # a leak.
        _ = get_working_set_size()
        # Actual measurement. `find_memleak` already warms up `no_leak`.
        leak_amount = find_memleak(no_leak, target="working_set")
        # The leak amount should be 0 (no net memory increase)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_private_usage_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        # Call find_memleak - it should not detect a leak
        leak_amount = find_memleak(no_leak, target="private_usage")
        # The leak amount should be 0 (no net memory increase)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_mallocspy_no_leak(self):
        def no_leak():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            malloc.Free(ptr)

        # Call find_memleak - it should not detect a leak
        leak_amount = find_memleak(no_leak, target="mallocspy")
        # The leak amount should be 0 (no net allocations)
        self.assertEqual(leak_amount, 0)

    def test_find_memleak_invalid_taget(self):
        with self.assertRaises(TypeError):
            find_memleak(lambda: None, target="invalid target")


if __name__ == "__main__":
    unittest.main()
