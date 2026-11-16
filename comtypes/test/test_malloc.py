import ctypes
import gc
import unittest
from dataclasses import dataclass

from comtypes import COMObject, hresult
from comtypes.malloc import (
    CoGetMalloc,
    IMallocSpy,
    _CoRegisterMallocSpy,
    _CoRevokeMallocSpy,
    _CoTaskMemAlloc,
    _CoTaskMemFree,
)


def _revoke_existing_spy() -> None:
    """Try to revoke any existing malloc spy, ignoring errors."""
    # _CoRevokeMallocSpy returns HRESULT, not raising exceptions
    # Keep trying to revoke until it succeeds or we get "not registered"
    # Retry on any error since outstanding allocations may take a moment to be fully released
    for _ in range(30):
        try:
            hr = _CoRevokeMallocSpy()
            # S_OK = 0, E_ACCESSDENIED = -2147024891, CO_E_NOT_REGISTERED = -2147220997
            # CO_E_ALREADY_REGISTERED = -2147220996
            if hr == 0 or hr == -2147220997:
                # Successfully revoked or wasn't registered
                break
            # If E_ACCESSDENIED or CO_E_ALREADY_REGISTERED or any other error, wait a bit and try again
            import time

            time.sleep(0.05)
            # Run gc again to try to free outstanding allocations
            gc.collect()
            gc.collect()
            gc.collect()
            continue
        except Exception:
            # Ignore all errors, but keep trying
            import time

            time.sleep(0.05)
            gc.collect()
            gc.collect()
            gc.collect()
            continue


class CleanUpMixin:
    def setUp(self):
        # Ensure no spy is registered before each test
        _revoke_existing_spy()
        gc.collect()
        gc.collect()
        gc.collect()

    def tearDown(self):
        # Clean up after each test
        # Only revoke if the test class didn't already handle it
        if not getattr(self, "_spy_revoked", False):
            _revoke_existing_spy()
        gc.collect()
        gc.collect()
        gc.collect()


@dataclass
class PtrAddr:
    """Wrapper for pointer addresses to distinguish them from other integer values in assertions."""

    value: int


class _CallRecorder:
    """Simple call recorder that accepts any arguments without signature checking."""

    def __init__(self, name: str):
        self.name = name
        self.calls: list[tuple] = []

    def __call__(self, *args, **kwargs):
        # Store arguments as a tuple to avoid COM marshaling issues
        # Pointers are passed as int values, so they can be stored directly
        self.calls.append(args)

    def assert_called_once(self):
        assert len(self.calls) == 1, f"Expected 1 call, got {len(self.calls)}"

    @property
    def call_args(self):
        class CallArgs:
            def __init__(self, args):
                self.args = args

        return CallArgs(self.calls[-1]) if self.calls else None


class MockMallocSpy(COMObject):
    """Test-only IMallocSpy implementation that records calls to a list.

    Each IMallocSpy_* method delegates to a _CallRecorder, allowing
    call order and argument verification without using unittest.mock.call/ANY.
    This avoids signature checking issues with MagicMock when called from COM.
    """

    _com_interfaces_ = [IMallocSpy]

    def __init__(self):
        super().__init__()
        # Create call recorders for each method
        self.PreAlloc = _CallRecorder("PreAlloc")
        self.PostAlloc = _CallRecorder("PostAlloc")
        self.PreFree = _CallRecorder("PreFree")
        self.PostFree = _CallRecorder("PostFree")
        self.PreRealloc = _CallRecorder("PreRealloc")
        self.PostRealloc = _CallRecorder("PostRealloc")
        self.PreGetSize = _CallRecorder("PreGetSize")
        self.PostGetSize = _CallRecorder("PostGetSize")
        self.PreDidAlloc = _CallRecorder("PreDidAlloc")
        self.PostDidAlloc = _CallRecorder("PostDidAlloc")
        self.PreHeapMinimize = _CallRecorder("PreHeapMinimize")
        self.PostHeapMinimize = _CallRecorder("PostHeapMinimize")
        # Track call order
        self._call_order: list[str] = []

    @property
    def call_names(self) -> list[str]:
        """Return list of method names called in order."""
        return self._call_order

    def reset_mocks(self) -> None:
        """Reset all call recorders."""
        for attr in [
            "PreAlloc",
            "PostAlloc",
            "PreFree",
            "PostFree",
            "PreRealloc",
            "PostRealloc",
            "PreGetSize",
            "PostGetSize",
            "PreDidAlloc",
            "PostDidAlloc",
            "PreHeapMinimize",
            "PostHeapMinimize",
        ]:
            getattr(self, attr).calls.clear()
        self._call_order.clear()

    def _record_call(self, name: str):
        self._call_order.append(name)

    # IMallocSpy methods - delegate to child recorders and return contract-compliant values

    def IMallocSpy_PreAlloc(self, this: int, cbRequest: int) -> int:
        self.PreAlloc(this, cbRequest)
        self._record_call("PreAlloc")
        return cbRequest

    def IMallocSpy_PostAlloc(self, this: int, pActual: int) -> int:
        self.PostAlloc(this, pActual)
        self._record_call("PostAlloc")
        return pActual

    def IMallocSpy_PreFree(self, this: int, pRequest: int, fSpyed: int) -> int:
        self.PreFree(this, pRequest, fSpyed)
        self._record_call("PreFree")
        return pRequest

    def IMallocSpy_PostFree(self, this: int, fSpyed: int) -> None:
        self.PostFree(this, fSpyed)
        self._record_call("PostFree")

    def IMallocSpy_PreRealloc(
        self,
        this: int,
        pRequest: int,
        cbRequest: int,
        ppNewRequest: "ctypes._Pointer[ctypes.c_void_p]",
        fSpyed: int,
    ) -> int:
        # ppNewRequest is an output parameter; we'll write pRequest to it per the contract
        # Pass the value that will be written as the third argument for verification
        new_ptr_value = pRequest
        if ppNewRequest:
            ppNewRequest[0] = pRequest
        self.PreRealloc(this, pRequest, cbRequest, new_ptr_value, fSpyed)
        self._record_call("PreRealloc")
        return cbRequest

    def IMallocSpy_PostRealloc(self, this: int, pActual: int, fSpyed: int) -> int:
        self.PostRealloc(this, pActual, fSpyed)
        self._record_call("PostRealloc")
        return pActual

    def IMallocSpy_PreGetSize(self, this: int, pRequest: int, fSpyed: int) -> int:
        self.PreGetSize(this, pRequest, fSpyed)
        self._record_call("PreGetSize")
        return pRequest

    def IMallocSpy_PostGetSize(self, this: int, cbActual: int, fSpyed: int) -> int:
        self.PostGetSize(this, cbActual, fSpyed)
        self._record_call("PostGetSize")
        return cbActual

    def IMallocSpy_PreDidAlloc(self, this: int, pRequest: int, fSpyed: int) -> int:
        self.PreDidAlloc(this, pRequest, fSpyed)
        self._record_call("PreDidAlloc")
        return pRequest

    def IMallocSpy_PostDidAlloc(
        self, this: int, pRequest: int, fSpyed: int, fActual: int
    ) -> int:
        self.PostDidAlloc(this, pRequest, fSpyed, fActual)
        self._record_call("PostDidAlloc")
        return fActual

    def IMallocSpy_PreHeapMinimize(self, this: int) -> None:
        self.PreHeapMinimize(this)
        self._record_call("PreHeapMinimize")

    def IMallocSpy_PostHeapMinimize(self, this: int) -> None:
        self.PostHeapMinimize(this)
        self._record_call("PostHeapMinimize")


class Test_MallocSpyRegistration(CleanUpMixin, unittest.TestCase):
    """Test IMallocSpy registration, revocation, and hook firing."""

    def setUp(self):
        super().setUp()
        self.spy = MockMallocSpy()
        # Try to revoke any existing spy first
        _revoke_existing_spy()
        hr = _CoRegisterMallocSpy(self.spy)
        # If spy is already registered (e.g., previous test couldn't revoke due to E_ACCESSDENIED),
        # try to revoke and re-register
        if (
            hr == -2147220996
        ):  # CO_E_ALREADY_REGISTERED / "Object is already registered"
            _revoke_existing_spy()
            hr = _CoRegisterMallocSpy(self.spy)
        self.assertEqual(hr, hresult.S_OK)
        self._spy_revoked = False
        # Ensure clean state by resetting mocks
        self.spy.reset_mocks()

    def tearDown(self):
        # Revoke spy, allowing E_ACCESSDENIED if there are outstanding allocations
        if not self._spy_revoked:
            # First, run gc to free any outstanding allocations that might prevent revocation
            gc.collect()
            gc.collect()
            gc.collect()
            # Keep trying to revoke until it succeeds or we get "not registered"
            # Retry on any error since outstanding allocations may take a moment to be fully released
            # Use many more retries with longer delays for stubborn cases
            for _ in range(200):
                try:
                    hr = _CoRevokeMallocSpy()
                    # S_OK = 0, E_ACCESSDENIED = -2147024891, CO_E_NOT_REGISTERED = -2147220997
                    if hr == 0 or hr == -2147220997:
                        # Successfully revoked or wasn't registered
                        self._spy_revoked = True
                        break
                    # If E_ACCESSDENIED or any other error, wait a bit and try again
                    import time

                    time.sleep(0.1)
                    # Run gc again to try to free outstanding allocations
                    gc.collect()
                    gc.collect()
                    gc.collect()
                    continue
                except PermissionError as e:
                    if e.winerror == -2147024891:  # E_ACCESSDENIED
                        import time

                        time.sleep(0.1)
                        # Run gc again to try to free outstanding allocations
                        gc.collect()
                        gc.collect()
                        gc.collect()
                        continue
                    else:
                        raise
                except Exception:
                    # Ignore all other errors, but keep trying
                    import time

                    time.sleep(0.1)
                    gc.collect()
                    gc.collect()
                    gc.collect()
                    continue
            # If we still couldn't revoke, mark as revoked anyway so setUp doesn't hang
            # The next test's setUp will call _revoke_existing_spy() which will try again
            self._spy_revoked = True
        super().tearDown()

    def test_no_spy_before_registration(self):
        """Before registration, Alloc/Free should not fire spy."""
        # Revoke the spy registered in setUp
        _CoRevokeMallocSpy()
        self._spy_revoked = True
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        # No calls should have been recorded
        self.assertEqual(self.spy.call_names, [])

    def test_registration_returns_s_ok(self):
        """_CoRegisterMallocSpy returns S_OK."""
        # Already registered in setUp, just verify we can revoke and re-register
        _CoRevokeMallocSpy()
        self._spy_revoked = True
        # Need a new spy instance since the old one was revoked
        new_spy = MockMallocSpy()
        hr = _CoRegisterMallocSpy(new_spy)
        self.assertEqual(hr, hresult.S_OK)
        # Clean up
        _CoRevokeMallocSpy()

    def test_alloc_free_fires_hooks(self):
        """_CoTaskMemAlloc/Free fire PreAlloc/PostAlloc and PreFree/PostFree."""
        self.spy.reset_mocks()

        ptr = _CoTaskMemAlloc(100)
        self.assertNotEqual(ptr, 0)
        _CoTaskMemFree(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        # Check call order
        self.assertEqual(
            self.spy.call_names, ["PreAlloc", "PostAlloc", "PreFree", "PostFree"]
        )

    def test_alloc_free_via_imalloc_fires_hooks(self):
        """IMalloc Alloc/Free fire PreAlloc/PostAlloc and PreFree/PostFree."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        self.assertNotEqual(ptr, 0)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.assertEqual(
            self.spy.call_names, ["PreAlloc", "PostAlloc", "PreFree", "PostFree"]
        )

    def test_realloc_fires_hooks(self):
        """IMalloc Realloc fires PreRealloc/PostRealloc."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr1 = malloc.Alloc(100)
        self.spy.reset_mocks()  # Reset after initial alloc

        ptr2 = malloc.Realloc(ptr1, 200)
        self.assertNotEqual(ptr2, 0)
        malloc.Free(ptr2)

        gc.collect()
        gc.collect()
        gc.collect()

        # Realloc should fire PreRealloc and PostRealloc
        self.assertEqual(
            self.spy.call_names, ["PreRealloc", "PostRealloc", "PreFree", "PostFree"]
        )

    def test_getsize_fires_hooks(self):
        """IMalloc GetSize fires PreGetSize/PostGetSize."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        self.spy.reset_mocks()

        size = malloc.GetSize(ptr)
        self.assertEqual(size, 100)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.assertEqual(
            self.spy.call_names, ["PreGetSize", "PostGetSize", "PreFree", "PostFree"]
        )

    def test_didalloc_fires_hooks(self):
        """IMalloc DidAlloc fires PreDidAlloc/PostDidAlloc."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        self.spy.reset_mocks()

        result = malloc.DidAlloc(ptr)
        self.assertEqual(result, 1)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.assertEqual(
            self.spy.call_names, ["PreDidAlloc", "PostDidAlloc", "PreFree", "PostFree"]
        )

    def test_heapminimize_fires_hooks(self):
        """IMalloc HeapMinimize fires PreHeapMinimize/PostHeapMinimize."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        malloc.HeapMinimize()

        gc.collect()
        gc.collect()
        gc.collect()

        self.assertEqual(self.spy.call_names, ["PreHeapMinimize", "PostHeapMinimize"])

    def test_prealloc_argument(self):
        """PreAlloc receives the requested size."""
        self.spy.reset_mocks()

        ptr = _CoTaskMemAlloc(12345)
        self.assertNotEqual(ptr, 0)
        _CoTaskMemFree(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.spy.PreAlloc.assert_called_once()
        # args[0] is 'this' (self), args[1] is cbRequest
        _, cbRequest = self.spy.PreAlloc.call_args.args
        self.assertEqual(cbRequest, 12345)

    def test_postalloc_argument(self):
        """PostAlloc receives the actual pointer (non-zero integer as proof)."""
        self.spy.reset_mocks()

        ptr = _CoTaskMemAlloc(100)
        self.assertNotEqual(ptr, 0)
        _CoTaskMemFree(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.spy.PostAlloc.assert_called_once()
        _, pActual = self.spy.PostAlloc.call_args.args
        # Verify a pointer was passed (non-zero)
        self.assertIsInstance(pActual, int)
        self.assertNotEqual(pActual, 0)
        # Can't assert exact value, but we can verify it matches our allocated pointer
        self.assertEqual(pActual, ptr)

    def test_prefree_argument(self):
        """PreFree receives the pointer (non-zero) and fSpyed flag."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        self.spy.reset_mocks()

        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.spy.PreFree.assert_called_once()
        _, pRequest, fSpyed = self.spy.PreFree.call_args.args
        # Verify a pointer was passed
        self.assertIsInstance(pRequest, int)
        self.assertNotEqual(pRequest, 0)
        self.assertEqual(pRequest, ptr)
        self.assertEqual(fSpyed, 1)  # Should be spied

    def test_prerealloc_arguments(self):
        """PreRealloc receives old pointer, new size, new pointer value, and fSpyed."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr1 = malloc.Alloc(100)
        self.spy.reset_mocks()

        ptr2 = malloc.Realloc(ptr1, 200)
        self.assertNotEqual(ptr2, 0)
        malloc.Free(ptr2)

        gc.collect()
        gc.collect()
        gc.collect()

        self.spy.PreRealloc.assert_called_once()
        _, pRequest, cbRequest, new_ptr_value, fSpyed = (
            self.spy.PreRealloc.call_args.args
        )
        # Verify old pointer was passed
        self.assertIsInstance(pRequest, int)
        self.assertNotEqual(pRequest, 0)
        self.assertEqual(pRequest, ptr1)
        # Verify new size
        self.assertEqual(cbRequest, 200)
        # Verify new pointer value (what was written to ppNewRequest)
        self.assertIsInstance(new_ptr_value, int)
        self.assertNotEqual(new_ptr_value, 0)
        self.assertEqual(new_ptr_value, ptr1)
        self.assertEqual(fSpyed, 1)

    def test_postgetsize_argument(self):
        """PostGetSize receives the actual size."""
        self.spy.reset_mocks()

        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        self.spy.reset_mocks()

        size = malloc.GetSize(ptr)
        self.assertEqual(size, 100)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        self.spy.PostGetSize.assert_called_once()
        _, cbActual, fSpyed = self.spy.PostGetSize.call_args.args
        self.assertEqual(cbActual, 100)
        self.assertEqual(fSpyed, 1)

    def test_non_com_activity_no_hooks(self):
        """Normal Python object creation/destruction does not fire spy."""
        self.spy.reset_mocks()

        # Create and destroy regular Python objects
        obj = object()
        del obj
        gc.collect()
        gc.collect()
        gc.collect()

        # No spy calls should have been made
        self.assertEqual(self.spy.call_names, [])

    def test_revocation_stops_hooks(self):
        """After revocation, Alloc/Free no longer fire spy."""
        # Revoke the spy
        hr = _CoRevokeMallocSpy()
        self.assertIn(hr, (hresult.S_OK, hresult.E_ACCESSDENIED))
        self._spy_revoked = True
        self.spy.reset_mocks()

        # Allocate and free after revocation
        malloc = CoGetMalloc()
        ptr = malloc.Alloc(100)
        malloc.Free(ptr)

        gc.collect()
        gc.collect()
        gc.collect()

        # No calls should be recorded
        self.assertEqual(self.spy.call_names, [])


if __name__ == "__main__":
    unittest.main()
