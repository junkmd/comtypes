"""Characterization tests for the lifetime of [out] record parameters.

A COM vtable method that returns a record through an [out] parameter
transfers ownership of the record's managed resources (BSTR, interface
pointers) to the caller.  comtypes currently does not release those
resources when the Python record object is destroyed, so they leak.

These tests record the current (leaking) behavior.  They are
characterization tests: they describe what the code does today, not what
it should do.  When <ISSUE_NUMBER> is fixed, the tests that record the
leak must be changed into normal-path tests.

The tests use the out-of-process test double `CoOutRecordTest`, which
returns records that own a BSTR and an IUnknown* sentinel.  The server
keeps a per-instance count of live sentinels, which the tests read
through `GetLiveSentinelCount`.  A live count of 1 after the record has
been dropped means the sentinel (and therefore the record's resources)
was leaked.
"""

import ctypes
import gc
import unittest
from _ctypes import COMError

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.recordinfo import GetRecordInfoFromGuids

# TODO: Replace <ISSUE_NUMBER> with the actual issue number.
_ISSUE_NUMBER = "<ISSUE_NUMBER>"

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


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestOutRecordLifetime(unittest.TestCase):
    def test_c1_out_record_is_leaked(self):
        """CHARACTERIZATION: an [out] record is not released on destruction.

        The record returned by GetRecord owns a BSTR and an IUnknown*
        sentinel.  Dropping the Python record and collecting garbage does
        not release them, so the server still reports one live sentinel.
        This is the leak that <ISSUE_NUMBER> is about; when it is fixed,
        this test must be changed to expect a live count of 0.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()

        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            1,
            msg="CHARACTERIZATION(#<ISSUE_NUMBER>): the [out] record is "
            "expected to be released once the issue is fixed.",
        )

    def test_c2_explicit_record_clear_releases(self):
        """CHARACTERIZATION: an explicit RecordClear releases the record.

        Calling IRecordInfo::RecordClear on the record releases the BSTR
        and the sentinel, so the live count drops to 0.  This shows that
        the resources are releasable; the leak is only that comtypes does
        not do this automatically.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        ri = _record_info()
        ri.RecordClear(ctypes.addressof(rec))  # type: ignore

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_c3_record_clear_zeroes_raw_bytes(self):
        """CHARACTERIZATION: RecordClear zeroes the record's raw bytes.

        After RecordClear, every pointer-width word of the record's raw
        bytes is expected to be zero.  The test reads the raw bytes with
        ctypes.string_at instead of reading the fields, so that it never
        dereferences a pointer that has already been released.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)

        ri = _record_info()
        ri.RecordClear(ctypes.addressof(rec))  # type: ignore

        raw = ctypes.string_at(ctypes.addressof(rec), ctypes.sizeof(rec))
        wordsize = ctypes.sizeof(ctypes.c_void_p)
        for i in range(0, len(raw), wordsize):
            word = int.from_bytes(raw[i : i + wordsize], "little")
            self.assertEqual(
                word,
                0,
                msg="CHARACTERIZATION(#<ISSUE_NUMBER>): RecordClear is "
                "expected to zero the record's raw bytes.",
            )

    def test_c4_return_shapes(self):
        """CHARACTERIZATION: the shape of the values returned by [out] records.

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

    def test_c5_fail_get_record_raises_and_creates_nothing(self):
        """CHARACTERIZATION: a failing method raises and creates no sentinel.

        FailGetRecord returns E_FAIL without touching the out record, so
        the call raises COMError and no sentinel is created.
        """
        ifc = _create_server()

        with self.assertRaises(COMError):
            ifc.FailGetRecord()

        self.assertEqual(ifc.GetLiveSentinelCount(), 0)

    def test_c6_huge_bstr_is_leaked(self):
        """CHARACTERIZATION: a huge BSTR in an [out] record is leaked.

        The record is marshaled to this process, so the BSTR is copied
        into this process.  Dropping the record without releasing it leaks
        that copy.  Returning 100 records with a 4 MiB BSTR each and
        dropping them is expected to grow this process's private usage by
        at least half of 100 * 4 MiB.  The threshold is based on the S4
        measurement, where leaking 100 * 4 MiB BSTRs grew private usage by
        about 100% of the expected size and freeing them grew it by 0%.
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

        self.assertGreaterEqual(
            after - before,
            expected * 0.5,
            msg="CHARACTERIZATION(#<ISSUE_NUMBER>): the huge BSTRs are "
            "expected to be released once the issue is fixed.",
        )


if __name__ == "__main__":
    unittest.main()
