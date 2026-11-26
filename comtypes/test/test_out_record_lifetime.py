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

Two tests (C2 and C3) are characterization tests of the system's
`IRecordInfo` implementation, not of comtypes.  They are kept as they
are because they record behavior that comtypes relies on.
"""

# TODO: Replace <ISSUE_NUMBER> with the actual issue number.  This module
# was added to fix the leak of managed resources owned by [out] record
# parameters.
_ISSUE_NUMBER = "<ISSUE_NUMBER>"

import ctypes
import gc
import unittest
import weakref
from _ctypes import COMError
from pathlib import Path
from typing import Any

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule
from comtypes.recordinfo import GetRecordInfoFromGuids
from comtypes.test.fault_log_helper import enable_fault_log, run_in_mp
from comtypes.test.find_memleak import find_memleak

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


# The record is kept in a module-level list so that it is still alive when
# the child process exits.  This is what makes the test exercise the
# `atexit=False` behavior of the finalizer.
_held_records: list[Any] = []


def _hold_out_record(tmpfile: Path) -> None:
    enable_fault_log(tmpfile)
    ifc = _create_server()
    _held_records.append(ifc.GetRecord(4))


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestOutRecordLifetime(unittest.TestCase):
    def test_t1_out_record_is_released(self):
        """An [out] record should release its resources when it is destroyed.

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  The record returned by GetRecord owns a
        BSTR and an IUnknown* sentinel, but comtypes does not adopt the
        record, so dropping the Python record and collecting garbage does
        not release them.  The server therefore still reports one live
        sentinel.

        If comtypes adopted the [out] record and released it on
        destruction, the live count would drop to 0 and this test would
        fail.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()

        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            1,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the [out] record is "
            "expected to be released once the issue is fixed, so the live "
            "sentinel count should be 0; but comtypes does not adopt the "
            "record, so the count stays at 1",
        )

    def test_t2_retval_record_is_released(self):
        """A [retval] record should be released just like a plain [out] record.

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  GetRecordRetval declares its record
        parameter as [out, retval], but comtypes does not adopt it, so
        dropping the record does not release the sentinel.

        If comtypes adopted the [retval] record and released it on
        destruction, the live count would drop to 0 and this test would
        fail.
        """
        ifc = _create_server()
        rec = ifc.GetRecordRetval()
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()

        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            1,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the [retval] record is "
            "expected to be released once the issue is fixed, so the live "
            "sentinel count should be 0; but comtypes does not adopt the "
            "record, so the count stays at 1",
        )

    def test_t3_multiple_out_records_are_released(self):
        """Every [out] record in a multi-out method should be released.

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  GetRecordsAround returns two records with
        a LONG in between.  The declaration order of the results (record,
        LONG, record) is preserved, but comtypes does not adopt either
        record, so dropping them does not release the sentinels.

        If comtypes adopted both [out] records and released them on
        destruction, the live count would drop to 0 and this test would
        fail.
        """
        ifc = _create_server()
        first, middle, last = ifc.GetRecordsAround()
        self.assertEqual(first.tag, 10)
        self.assertEqual(middle, 42)
        self.assertEqual(last.tag, 20)
        self.assertEqual(ifc.GetLiveSentinelCount(), 2)

        del first, last
        gc.collect()

        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            2,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): both [out] records are "
            "expected to be released once the issue is fixed, so the live "
            "sentinel count should be 0; but comtypes does not adopt the "
            "records, so the count stays at 2",
        )

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

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  from_buffer_copy copies the raw bytes but
        not the ownership token, so dropping the copy does not release the
        sentinel.  comtypes does not adopt the original record either, so
        dropping the original does not release it either.

        If comtypes adopted the original [out] record and released it on
        destruction, the live count would drop to 0 and this test would
        fail.
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
        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            1,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the original [out] "
            "record is expected to be released once the issue is fixed, so "
            "the live sentinel count should be 0; but comtypes does not "
            "adopt the record, so the count stays at 1",
        )

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
        """A huge BSTR in an [out] record should be released.

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  The record is marshaled to this process,
        so the BSTR is copied into this process.  comtypes does not adopt
        the record, so dropping it leaks that copy.  Returning 100 records
        with a 4 MiB BSTR each and dropping them grows this process's
        private usage by roughly the full 100 * 4 MiB.

        The threshold is based on the S4 measurement, where leaking
        100 * 4 MiB BSTRs grew private usage by about 100% of the expected
        size and freeing them grew it by 0%.  A BSTR cannot be observed
        directly from Python, so this is a hypothesis based on the measured
        behavior.  If comtypes released the BSTRs, the growth would drop to
        about 0% and this test would fail.
        """
        ifc = _create_server()
        n = 100
        size = 2 * 1024 * 1024  # wchars -> 4 MiB per BSTR
        expected = n * size * 2

        def get_and_drop_record() -> None:
            rec = ifc.GetRecord(size)
            del rec

        leaked = find_memleak(get_and_drop_record, loops=(1, n), target="private_usage")

        self.assertGreaterEqual(
            leaked,
            expected * 0.5,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the huge BSTRs are "
            "expected to be released once the issue is fixed, so the "
            "private usage should not grow; but comtypes does not adopt "
            "the record, so the BSTRs leak",
        )

    def test_t10_child_process_exits_normally(self):
        """Holding an [out] record must not prevent a clean interpreter exit.

        The finalizer is registered with atexit=False, so it does not run at
        interpreter shutdown.  A child process that holds an out record and
        exits must exit with code 0.  If the finalizer ran at shutdown, it
        would call RecordClear after COM has been uninitialized, which can
        crash the process.
        """
        exitcode, output = run_in_mp(_hold_out_record)
        self.assertEqual(exitcode, 0, msg=f"Process crashed with output: {output}")

    def test_t11_record_owns_its_release(self):
        """The record itself should carry the release obligation.

        CHARACTERIZATION: this records the current behavior, which is not
        necessarily desirable.  The record returned by the call has no
        finalizer attached, so it is not weak-referenceable and dropping
        it does not release the sentinel.

        If comtypes attached a finalizer to the record object itself, the
        record would have weakrefs and dropping it would release the
        sentinel; this test would then fail.
        """
        ifc = _create_server()
        rec = ifc.GetRecord(4)
        self.assertEqual(
            weakref.getweakrefs(rec),
            [],
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the record is expected "
            "to carry a finalizer (and therefore be weak-referenceable) "
            "once the issue is fixed; but comtypes does not adopt the "
            "record, so it has no weakrefs",
        )
        self.assertEqual(ifc.GetLiveSentinelCount(), 1)

        del rec
        gc.collect()
        self.assertEqual(
            ifc.GetLiveSentinelCount(),
            1,
            msg=f"CHARACTERIZATION(#{_ISSUE_NUMBER}): the record is expected "
            "to release its sentinel when it is destroyed once the issue is "
            "fixed, so the live count should be 0; but comtypes does not "
            "adopt the record, so the count stays at 1",
        )

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
        """An explicit RecordClear releases the record.

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
        """RecordClear zeroes the record's pointer fields.

        After RecordClear, the words that hold the record's pointer fields
        (the BSTR and the interface pointer) are expected to be zero.  The
        test reads the raw bytes with ctypes.string_at instead of reading
        the fields, so that it never dereferences a pointer that has
        already been released.  The non-pointer field (tag) is not zeroed;
        only the pointer fields are.

        comtypes relies on it: the finalizer calls RecordClear again when
        the record is destroyed, and that second call is only safe if the
        first call zeroed the pointer fields.
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
            self.assertEqual(word, 0)


if __name__ == "__main__":
    unittest.main()
