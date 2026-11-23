"""DATE (VT_DATE) parameter marshaling tests.

These tests exercise a native COM test double that exposes DATE parameters
and return values through both a dual interface (vtable) and a dispinterface
(IDispatch::Invoke).

The type library declares DATE as a typedef of double.  As a result the dual
interface exposes DATE as a plain ``double`` on the vtable, while the
dispinterface marshals it as a ``VT_DATE`` VARIANT.  The two paths therefore
behave differently when a ``datetime.datetime`` is passed in or returned.

The tests in this module record the current behavior.  They are
characterization tests: the observed behavior is not necessarily the desired
behavior.
"""

import ctypes
import datetime
import unittest
from ctypes import c_double, pointer

from comtypes import CLSCTX_LOCAL_SERVER, COMError
from comtypes.client import CreateObject, GetModule

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


# 2024-01-02 12:00:00 expressed as an OLE Automation date (days since
# 1899-12-30, where the fractional part is the time of day).  The native test
# double uses this exact value in InitDate and VerifyDate.
FIXED_DATE = 45293.5
FIXED_DATETIME = datetime.datetime(2024, 1, 2, 12, 0, 0)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class Test_DualDateParamTest(unittest.TestCase):
    """DATE parameters through the dual (vtable) interface.

    The type library declares DATE as a typedef of double, so the vtable
    exposes DATE as ``c_double``.  These tests record how that affects passing
    and returning dates.
    """

    def _create(self) -> "ComtypesCppTestSrvLib.IDualDateParamTest":
        return CreateObject(
            "Comtypes.DispDateParamTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualDateParamTest,
        )

    def test_echo_date_float_returns_float(self):
        """A float OLE date round-trips unchanged through the vtable.

        Because DATE is a plain double on the vtable, the value is passed and
        returned as a float with no datetime conversion.  If this fails, the
        vtable DATE representation changed.
        """
        result = self._create().EchoDate(FIXED_DATE)
        self.assertEqual(result, FIXED_DATE)
        self.assertIsInstance(result, float)

    def test_echo_date_datetime_raises_argument_error(self):
        """A datetime cannot be passed to the vtable DATE parameter.

        CHARACTERIZATION: the vtable DATE is ``c_double``, whose ``from_param``
        rejects ``datetime.datetime``.  This is the marshaling problem: a date
        object cannot be sent through the dual interface.  If DATE learns to
        marshal datetime, this test is expected to fail and should be turned
        into a normal-path test.
        """
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION: DATE as a double typedef rejects datetime; "
            "if DATE learns to marshal datetime this test should fail.",
        ):
            self._create().EchoDate(FIXED_DATETIME)

    def test_verify_date_float(self):
        """The server recognizes the fixed OLE date passed as a float."""
        self.assertTrue(self._create().VerifyDate(FIXED_DATE))

    def test_verify_date_datetime_raises_argument_error(self):
        """A datetime cannot be passed to the vtable DATE parameter.

        CHARACTERIZATION: as with EchoDate, the vtable DATE is ``c_double``
        and rejects ``datetime.datetime``.  If DATE learns to marshal datetime,
        this test is expected to fail and should be turned into a normal-path
        test.
        """
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION: DATE as a double typedef rejects datetime; "
            "if DATE learns to marshal datetime this test should fail.",
        ):
            self._create().VerifyDate(FIXED_DATETIME)

    def test_init_date_pointer(self):
        """An [in, out] DATE is written back through a pointer.

        The vtable [in, out] DATE is a pointer to double, so the server can
        overwrite the caller's buffer.  This records that the pointer path
        works even though the value path cannot take a datetime.
        """
        value = c_double(0.0)
        self._create().InitDate(pointer(value))
        self.assertEqual(value.value, FIXED_DATE)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class Test_DispDateParamTest(unittest.TestCase):
    """DATE parameters through the dispinterface (IDispatch::Invoke).

    The dispinterface marshals DATE as a ``VT_DATE`` VARIANT, so a
    ``datetime.datetime`` is accepted and returned as a datetime.  These tests
    record the difference from the dual interface.
    """

    def _create(self) -> "ComtypesCppTestSrvLib.IDispDateParamTest":
        return CreateObject(
            "Comtypes.DispDateParamTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispDateParamTest,
        )

    def test_echo_date_datetime_returns_datetime(self):
        """A datetime round-trips as a datetime through the dispinterface.

        The VARIANT is VT_DATE, so the dispinterface accepts and returns a
        ``datetime.datetime`` even though the vtable DATE is a double.  This is
        the behavior the dual interface lacks.
        """
        result = self._create().EchoDate(FIXED_DATETIME)
        self.assertEqual(result, FIXED_DATETIME)
        self.assertIsInstance(result, datetime.datetime)

    def test_echo_date_float_returns_datetime(self):
        """A float OLE date is returned as a datetime through the dispinterface.

        CHARACTERIZATION: the dispinterface marshals DATE as VT_DATE, so even
        a float input comes back as a ``datetime.datetime``.  This asymmetry
        (float in, datetime out) is a consequence of DATE being a double
        typedef on the vtable but VT_DATE on the dispinterface.
        """
        result = self._create().EchoDate(FIXED_DATE)
        self.assertEqual(result, FIXED_DATETIME)
        self.assertIsInstance(result, datetime.datetime)

    def test_verify_date_float(self):
        """The server recognizes the fixed OLE date passed as a float."""
        self.assertTrue(self._create().VerifyDate(FIXED_DATE))

    def test_verify_date_datetime(self):
        """The server recognizes the fixed OLE date passed as a datetime."""
        self.assertTrue(self._create().VerifyDate(FIXED_DATETIME))

    def test_init_date_pointer_raises_comerror(self):
        """An [in, out] DATE cannot be passed as a pointer to the dispinterface.

        CHARACTERIZATION: the dispinterface expects a VT_DATE VARIANT for the
        [in, out] DATE, but a ``pointer(c_double)`` is marshaled as a
        VT_BYREF|VT_R8 VARIANT, which the server rejects with
        DISP_E_TYPEMISMATCH.  If DATE learns to marshal datetime, this test is
        expected to fail and should be turned into a normal-path test.
        """
        value = c_double(0.0)
        with self.assertRaises(
            COMError,
            msg="CHARACTERIZATION: the dispinterface rejects a pointer(c_double) "
            "for an [in, out] DATE; if DATE learns to marshal datetime this "
            "test should fail.",
        ):
            self._create().InitDate(pointer(value))


if __name__ == "__main__":
    unittest.main()
