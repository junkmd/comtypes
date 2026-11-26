"""Tests for the lifetime of [out] record parameters.

A COM vtable method that returns a record through an [out] parameter
transfers ownership of the record's managed resources (BSTR, interface
pointers) to the caller.  comtypes releases those resources when the
Python record object is destroyed, by calling
`IRecordInfo::RecordClear` from a finalizer.

The tests use the out-of-process test double `CoOutRecordTest`, which
returns records that own a BSTR and an IUnknown* sentinel.  The server
keeps a per-instance count of live sentinels, which the tests read
through `GetLiveSentinelCount`.  A live count of 0 after the record has
been dropped means the record's resources were released.

Two tests (C2 and C3) are characterization tests of the system's
`IRecordInfo` implementation, not of comtypes.  They are kept as they
are because they record behavior that comtypes relies on.
"""

# TODO: Replace <ISSUE_NUMBER> with the actual issue number.  This module
# was added to fix the leak of managed resources owned by [out] record
# parameters.
_ISSUE_NUMBER = "<ISSUE_NUMBER>"

import ctypes
import faulthandler
import gc
import io
import multiprocessing
import tempfile
import unittest
import weakref
from _ctypes import COMError
from collections.abc import Callable
from pathlib import Path
from typing import Any

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.recordinfo import GetRecordInfoFromGuids

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    # Server type library not registered; tests depending on it will be skipped.
    IMPORT_FAILED = True


def _create_server() -> "ComtypesCppTestSrvLib.IOutRecordTest":
    """Create an out-of-process CoOutRecordTest instance.

    Uses CLSCTX_LOCAL_SERVER so that the record is marshaled across a
    process boundary.  This is what makes the BSTR and the IUnknown*
    sentinel real managed resources that must be released by the caller.
    """
    return CreateObject(
        "Comtypes.OutRecordTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IOutRecordTest,
    )


def _record_info() -> "ComtypesCppTestSrvLib.IRecordInfo":
    return GetRecordInfoFromGuids(
        *ComtypesCppTestSrvLib.StructOutRecordTest._recordinfo_
    )


class _PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_ulong),
        ("PageFaultCount", ctypes.c_ulong),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
        ("PrivateUsage", ctypes.c_size_t),
    ]


_psapi = ctypes.WinDLL("psapi")
_GetProcessMemoryInfo = _psapi.GetProcessMemoryInfo
_GetProcessMemoryInfo.argtypes = [
    ctypes.c_void_p,
    ctypes.POINTER(_PROCESS_MEMORY_COUNTERS_EX),
    ctypes.c_ulong,
]
_GetProcessMemoryInfo.restype = ctypes.c_int

_kernel32 = ctypes.WinDLL("kernel32")
_GetCurrentProcess = _kernel32.GetCurrentProcess
_GetCurrentProcess.restype = ctypes.c_void_p


def _private_usage() -> int:
    counters = _PROCESS_MEMORY_COUNTERS_EX()
    counters.cb = ctypes.sizeof(counters)
    _GetProcessMemoryInfo(_GetCurrentProcess(), ctypes.byref(counters), counters.cb)
    return counters.PrivateUsage


def _enable_fault_log(tmpfile: Path) -> io.TextIOWrapper:
    """Send faulthandler output to `tmpfile` until the process exits."""
    f = tmpfile.open(mode="w", encoding="utf-8")
    # We keep `faulthandler` enabled and the file open until the process
    # exits so that crash information can be captured.
    faulthandler.enable(file=f)
    return f


def _run_in_mp(fn: Callable[[Path], Any]) -> tuple[int | None, str]:
    """Run a function in a separate process and return exit code and output.

    Using spawn isolates the process so crashes like segfaults do not affect
    the main test process.
    """
    with tempfile.TemporaryDirectory() as td:
        tmpdir = Path(td)
        tmpfile = tmpdir / "fault.txt"
        ctx = multiprocessing.get_context("spawn")
        process = ctx.Process(target=fn, args=(tmpfile,))
        process.start()
        process.join()
        output = tmpfile.read_text()
    return process.exitcode, output


# The record is kept in a module-level list so that it is still alive when
# the child process exits.  This is what makes the test exercise the
# `atexit=False` behavior of the finalizer.
_held_records: list[Any] = []


def _hold_out_record(tmpfile: Path) -> None:
    _enable_fault_log(tmpfile)
    ifc = _create_server()
    _held_records.append(ifc.GetRecord(4))


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestOutRecordLifetime(unittest.TestCase):
    def test_t1_out_record_is_released(self):
        """An [out] record releases its resources when it is destroyed.

        The record returned by GetRecord owns a BSTR and an IUnknown*
        sentinel.  Dropping the Python record and collecting garbage must
        release them, so the server reports no live sentinel.

        Why this shows there is no leak: the sentinel is only destroyed
        when its reference count reaches zero, and the server's live count
        is decremented exactly when the sentinel is destroyed.  A live
        count of 0 therefore means the reference held by the record was
        released.  If comtypes did not release the record (for example, if
        it returned the raw record without adopting it), the reference
        would never be released and the live count would stay at 1.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t2_retval_record_is_released(self):
        """A [retval] record is released just like a plain [out] record.

        GetRecordRetval declares its record parameter as [out, retval].
        The [retval] flag must not change the ownership handling, so the
        record must be released when it is destroyed.
        """
        ifc = _create_server()
        rec = ifc.GetRecordRetval()
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t3_multiple_out_records_are_released(self):
        """Every [out] record in a multi-out method is released.

        GetRecordsAround returns two records with a LONG in between.  Both
        records must be released, and the declaration order of the results
        (record, LONG, record) must be preserved.
        """
        ifc = _create_server()
        first, middle, last = ifc.GetRecordsAround()
        self.assertEqual(first.tag, 10)
        self.assertEqual(middle, 42)
        self.assertEqual(last.tag, 20)
        self.assertEqual(ifc.GetLiveSentinelCount(), 2)

        del first, last
        gc.collect()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t4_inout_record_is_not_released(self):
        """An [in, out] record is not auto-released.

        The ownership of an [in, out] record's contents is ambiguous: the
        caller passes the record in and the server may replace its
        contents.  comtypes therefore does not adopt [in, out] records.
        Dropping the result leaves the sentinel alive.
        """
        ifc = _create_server()
        inout = ComtypesCppTestSrvLib.StructOutRecordTest()
        inout.tag = 5
        modified = ifc.ModifyRecord(inout)
        self.assertEqual(modified.tag, 6)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del modified
        gc.collect()

        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

    def test_t5_fail_get_record_raises_and_creates_nothing(self):
        """A failing method raises and creates no sentinel.

        FailGetRecord returns E_FAIL without touching the out record, so
        the call raises COMError and no sentinel is created.  This also
        shows that the adoption wrapper does not run when the call fails.
        """
        ifc = _create_server()

        with self.assertRaises(COMError):
            ifc.FailGetRecord()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t6_from_buffer_copy_does_not_release(self):
        """A copy made with from_buffer_copy does not own the resources.

        from_buffer_copy copies the raw bytes but not the ownership token,
        so dropping the copy must not release the sentinel.  Dropping the
        original record must release it.  This shows that the ownership is
        attached to the record instance, not to the bytes.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        copy = ComtypesCppTestSrvLib.StructOutRecordTest.from_buffer_copy(rec)
        del copy
        gc.collect()
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()
        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t7_explicit_record_clear_then_del(self):
        """An explicit RecordClear followed by destruction must not double free.

        The record's finalizer also calls RecordClear.  If the system's
        RecordClear zeroes the record, the second call is a no-op.  This
        test checks that no crash happens and that the sentinel is
        released exactly once.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        ri = _record_info()
        ri.RecordClear(ctypes.addressof(rec))  # type: ignore
        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

        del rec
        gc.collect()
        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t8_reading_fields_does_not_release(self):
        """Reading the record's fields does not release its resources.

        The record must stay valid while it is alive.  Reading the BSTR
        repeatedly and reading the interface pointer (keeping the returned
        pointer alive) must not release them, so the live count stays at 1
        until the record is dropped.

        The interface pointer is kept in a local variable on purpose: the
        pointer returned by the field does not own a reference, so letting
        it be garbage collected would release the record's reference.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)

        self.assertEqual(rec.tag, 1)
        self.assertEqual(rec.text, "aaaa")
        self.assertEqual(rec.tag, 1)
        self.assertEqual(rec.text, "aaaa")
        unk = rec.unk
        self.assertTrue(unk)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del unk
        del rec
        gc.collect()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_t9_huge_bstr_is_released(self):
        """A huge BSTR in an [out] record is released.

        The record is marshaled to this process, so the BSTR is copied into
        this process.  Returning 100 records with a 4 MiB BSTR each and
        dropping them must not grow this process's private usage by more
        than a small fraction of 100 * 4 MiB.

        Why this shows there is no leak: the BSTR is a heap allocation in
        this process, so leaking it would grow private usage by roughly the
        full 100 * 4 MiB.  The threshold is based on the S4 measurement,
        where leaking 100 * 4 MiB BSTRs grew private usage by about 100% of
        the expected size and freeing them grew it by 0%.  A BSTR cannot be
        observed directly from Python, so this is a hypothesis based on the
        measured behavior.
        """
        ifc = _create_server()
        n = 100
        size = 2 * 1024 * 1024  # wchars -> 4 MiB per BSTR
        expected = n * size * 2

        before = _private_usage()
        for _ in range(n):
            rec = ifc.GetRecord(size)
            del rec
        gc.collect()
        after = _private_usage()

        self.assertLessEqual(after - before, expected * 0.1)

    def test_t10_child_process_exits_normally(self):
        """Holding an [out] record must not prevent a clean interpreter exit.

        The finalizer is registered with atexit=False, so it does not run at
        interpreter shutdown.  A child process that holds an out record and
        exits must exit with code 0.  If the finalizer ran at shutdown, it
        would call RecordClear after COM has been uninitialized, which can
        crash the process.
        """
        exitcode, output = _run_in_mp(_hold_out_record)
        self.assertEqual(exitcode, 0, msg=f"Process crashed with output: {output}")

    def test_t11_record_owns_its_release(self):
        """The record itself carries the release obligation.

        The record returned by the call is not copied: the finalizer that
        releases its resources is attached to the record object itself.
        This test checks that the record is weak-referenceable and has a
        finalizer attached, and that dropping it releases the sentinel.

        Why this matters: a design that copied the record into a separate
        buffer would attach the finalizer to the buffer instead, so the
        record would have no weakrefs.  Attaching the finalizer to the
        record avoids the shallow copy and the layout assumptions that a
        copy would require.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertTrue(weakref.getweakrefs(rec))
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()
        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_return_shapes(self):
        """The shape of the values returned by [out] records.

        A single [out] record and a [retval] record are returned as the
        record instance itself (not a pointer).  Multiple [out] parameters
        are returned as a tuple in declaration order.  An [in, out] record
        is returned as the record instance with the server's modification
        applied.
        """
        ifc = _create_server()

        rec = ifc.GetRecord(4)
        self.assertIsInstance(rec, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(rec.tag, 1)
        self.assertEqual(rec.text, "aaaa")

        retval = ifc.GetRecordRetval()
        self.assertIsInstance(retval, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(retval.tag, 2)

        first, middle, last = ifc.GetRecordsAround()
        self.assertIsInstance(first, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(first.tag, 10)
        self.assertEqual(middle, 42)
        self.assertIsInstance(last, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(last.tag, 20)

        inout = ComtypesCppTestSrvLib.StructOutRecordTest()
        inout.tag = 5
        modified = ifc.ModifyRecord(inout)
        self.assertIsInstance(modified, ComtypesCppTestSrvLib.StructOutRecordTest)
        self.assertEqual(modified.tag, 6)

    def test_c2_explicit_record_clear_releases(self):
        """CHARACTERIZATION: an explicit RecordClear releases the record.

        Calling IRecordInfo::RecordClear on the record releases the BSTR
        and the sentinel, so the live count drops to 0.  This is a
        characterization of the system's IRecordInfo implementation, which
        comtypes relies on; it is not a test of comtypes itself.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        ri = _record_info()
        ri.RecordClear(ctypes.addressof(rec))  # type: ignore

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_c3_record_clear_zeroes_pointer_fields(self):
        """CHARACTERIZATION: RecordClear zeroes the record's pointer fields.

        After RecordClear, the words that hold the record's pointer fields
        (the BSTR and the interface pointer) are expected to be zero.  The
        test reads the raw bytes with ctypes.string_at instead of reading
        the fields, so that it never dereferences a pointer that has
        already been released.  The non-pointer field (tag) is not zeroed;
        only the pointer fields are.

        This is a characterization of the system's IRecordInfo
        implementation.  comtypes relies on it: the finalizer calls
        RecordClear again when the record is destroyed, and that second
        call is only safe if the first call zeroed the pointer fields.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)

        ri = _record_info()
        ri.RecordClear(ctypes.addressof(rec))  # type: ignore

        raw = ctypes.string_at(ctypes.addressof(rec), ctypes.sizeof(rec))
        wordsize = ctypes.sizeof(ctypes.c_void_p)
        for field in ("text", "unk"):
            offset = getattr(ComtypesCppTestSrvLib.StructOutRecordTest, field).offset
            word = int.from_bytes(raw[offset : offset + wordsize], "little")
            self.assertEqual(
                word,
                0,
                msg="CHARACTERIZATION: RecordClear is expected to zero the "
                "record's pointer fields.",
            )


if __name__ == "__main__":
    unittest.main()
