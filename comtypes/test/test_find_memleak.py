import gc
import unittest
from collections.abc import Callable
from typing import Any

from comtypes import hresult
from comtypes.malloc import CoGetMalloc, _CoRegisterMallocSpy, _CoRevokeMallocSpy
from comtypes.test.find_memleak import (
    MallocSpy,
    _EmptyWorkingSet,
    create_memleak_detector,
    find_memleak,
    get_working_set_size,
)


def _revoke_existing_spy() -> None:
    """Try to revoke any existing malloc spy, ignoring errors."""
    try:
        _CoRevokeMallocSpy()
    except OSError:
        pass


def force_memory_reclaim() -> None:
    """Force memory reclaim by calling HeapMinimize and trimming working set.

    This helps ensure that freed COM memory is reflected in process memory metrics
    (Working Set / Private Usage) immediately.
    """
    # 1. Minimize COM heap
    try:
        malloc = CoGetMalloc()
    except OSError:
        pass
    malloc.HeapMinimize()
    # 2. Force trim working set (psapi.dll)
    try:
        # -1 is a pseudo-handle representing the current process.
        _EmptyWorkingSet(-1)
    except OSError:
        pass

    # 3. Garbage collection
    gc.collect()
    gc.collect()
    gc.collect()


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


class Test_create_memleak_detector(CleanUpMixin, unittest.TestCase):
    def test_working_set_detects_leak_and_cleanup(self):
        # Track allocated pointers so we can free them later
        allocated_pointers = []

        def leak_memory():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            allocated_pointers.append(ptr)
            # Intentionally not freeing - this is a leak

        def cleanup_memory():
            malloc = CoGetMalloc()
            for ptr in allocated_pointers:
                malloc.Free(ptr)
            allocated_pointers.clear()

        # Use detector as a context manager
        with create_memleak_detector("working_set") as detector:
            # Phase 1: Leak memory (10000 iterations to match find_memleak's measurement phase)
            for _ in range(10000):
                leak_memory()

            # Check that leak is detected (memory increased)
            # get_diff() can be called while the detector is active
            leak_amount = detector.get_diff()
            self.assertGreater(leak_amount, 0, "Leak should be detected before cleanup")

            # Phase 2: Cleanup
            cleanup_memory()

            force_memory_reclaim()

            # Check that leak is resolved (memory should return close to baseline)
            # With force_memory_reclaim, memory should be significantly reduced
            # Use a threshold proportional to the detected leak (25% of leak amount)
            # since process metrics may not return exactly to 0 due to OS memory management
            # (EmptyWorkingSet trims working set but Private Usage/commit charge may not fully return)
            post_cleanup_amount = detector.get_diff()
            # After cleanup with force_memory_reclaim, memory should be close to baseline
            # Allow up to 25% of the leak amount as residual (OS may not reclaim all pages immediately)
            threshold = max(
                leak_amount // 4, 1024 * 1024
            )  # At least 1MB, or 25% of leak
            self.assertLess(
                abs(post_cleanup_amount),
                threshold,
                f"Memory should return close to baseline after cleanup (threshold={threshold}, leak={leak_amount})",
            )

    def test_private_usage_detects_leak_and_cleanup(self):
        # Track allocated pointers so we can free them later
        allocated_pointers = []

        def leak_memory():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            allocated_pointers.append(ptr)
            # Intentionally not freeing - this is a leak

        def cleanup_memory():
            malloc = CoGetMalloc()
            for ptr in allocated_pointers:
                malloc.Free(ptr)
            allocated_pointers.clear()

        # Use detector as a context manager
        with create_memleak_detector("private_usage") as detector:
            # Phase 1: Leak memory (10000 iterations to match find_memleak's measurement phase)
            for _ in range(10000):
                leak_memory()

            # Check that leak is detected (memory increased)
            # get_diff() can be called while the detector is active
            leak_amount = detector.get_diff()
            self.assertGreater(leak_amount, 0, "Leak should be detected before cleanup")

            # Phase 2: Cleanup
            cleanup_memory()

            force_memory_reclaim()

            # Check that leak is resolved (memory should return close to baseline)
            # With force_memory_reclaim, memory should be significantly reduced
            # Use a threshold proportional to the detected leak (25% of leak amount)
            # since process metrics may not return exactly to 0 due to OS memory management
            # (EmptyWorkingSet trims working set but Private Usage/commit charge may not fully return)
            post_cleanup_amount = detector.get_diff()
            # After cleanup with force_memory_reclaim, memory should be close to baseline
            # Allow up to 25% of the leak amount as residual (OS may not reclaim all pages immediately)
            threshold = max(
                leak_amount // 4, 1024 * 1024
            )  # At least 1MB, or 25% of leak
            self.assertLess(
                abs(post_cleanup_amount),
                threshold,
                f"Memory should return close to baseline after cleanup (threshold={threshold}, leak={leak_amount})",
            )

    def test_mallocspy_detects_leak_and_cleanup(self):
        # Track allocated pointers so we can free them later
        allocated_pointers = []

        def leak_memory():
            malloc = CoGetMalloc()
            ptr = malloc.Alloc(1024)
            allocated_pointers.append(ptr)
            # Intentionally not freeing - this is a leak

        def cleanup_memory():
            malloc = CoGetMalloc()
            for ptr in allocated_pointers:
                malloc.Free(ptr)
            allocated_pointers.clear()

        # Use MallocSpyDetector as a context manager
        with create_memleak_detector("mallocspy") as detector:
            # Phase 1: Leak memory
            for _ in range(10):
                leak_memory()

            gc.collect()
            gc.collect()
            gc.collect()

            # Check that leak is detected (net allocations > 0)
            # get_diff() can be called while the detector is active
            leak_amount = detector.get_diff()
            self.assertGreater(leak_amount, 0, "Leak should be detected before cleanup")

            # Phase 2: Cleanup
            cleanup_memory()

            gc.collect()
            gc.collect()
            gc.collect()

            # Check that leak is resolved (net allocations == 0)
            post_cleanup_amount = detector.get_diff()
            self.assertEqual(
                post_cleanup_amount, 0, "No leak should remain after cleanup"
            )
        # Context manager exits here, revoking the spy
        self.assertEqual(detector.get_diff(), post_cleanup_amount)

    def test_invalid_taget(self):
        with self.assertRaises(TypeError):
            create_memleak_detector("invalid target")


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
