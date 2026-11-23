"""DATE (VT_DATE) parameter marshaling tests.

These tests exercise a native COM test double that exposes DATE parameters
and return values through both a dual interface (vtable) and a dispinterface
(IDispatch::Invoke).

The type library declares DATE as a typedef of double.  ``comtypes`` maps that
typedef to ``comtypes.automation.DATE``, a ``double`` subclass that converts to
and from ``datetime.datetime``.  The dual interface therefore accepts and
returns date objects, and the dispinterface marshals DATE as a ``VT_DATE``
VARIANT, so both paths agree on ``datetime.datetime``.
"""

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

    The vtable exposes DATE as ``comtypes.automation.DATE``, a ``double``
    subclass that converts to and from ``datetime.datetime``.  These tests
    verify that a date object can be passed in and is returned as a date
    object.
    """

    def _create(self) -> "ComtypesCppTestSrvLib.IDualDateParamTest":
        return CreateObject(
            "Comtypes.DispDateParamTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDualDateParamTest,
        )

    def test_echo_date_datetime_returns_datetime(self):
        """A datetime round-trips as a datetime through the vtable.

        DATE converts a ``datetime.datetime`` to an OLE date on the way in and
        back to a ``datetime.datetime`` on the way out, so the dual interface
        can marshal date objects.  If this fails, DATE no longer accepts or
        returns ``datetime.datetime``.
        """
        result = self._create().EchoDate(FIXED_DATETIME)
        self.assertEqual(result, FIXED_DATETIME)
        self.assertIsInstance(result, datetime.datetime)

    def test_echo_date_float_returns_datetime(self):
        """A float OLE date is returned as a datetime through the vtable.

        DATE accepts a numeric OLE date and returns a ``datetime.datetime``, so
        the vtable and the dispinterface agree on the return type.
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

    def test_init_date_pointer(self):
        """An [in, out] DATE is written back through a pointer.

        The vtable [in, out] DATE is a pointer to DATE, so the server can
        overwrite the caller's buffer.  The caller passes a ``DATE`` instance
        and reads the raw OLE date back from ``.value``.
        """
        value = ComtypesCppTestSrvLib.DATE(0.0)
        self._create().InitDate(pointer(value))
        self.assertEqual(value.value, FIXED_DATE)


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class Test_DispDateParamTest(unittest.TestCase):
    """DATE parameters through the dispinterface (IDispatch::Invoke).

    The dispinterface marshals DATE as a ``VT_DATE`` VARIANT, so a
    ``datetime.datetime`` is accepted and returned as a datetime.
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
        ``datetime.datetime``.
        """
        result = self._create().EchoDate(FIXED_DATETIME)
        self.assertEqual(result, FIXED_DATETIME)
        self.assertIsInstance(result, datetime.datetime)

    def test_echo_date_float_returns_datetime(self):
        """A float OLE date is returned as a datetime through the dispinterface.

        The dispinterface marshals DATE as VT_DATE, so a float input comes back
        as a ``datetime.datetime``.
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
        DISP_E_TYPEMISMATCH.  This records the remaining gap between the
        dispinterface and the dual interface for [in, out] DATE parameters.
        """
        value = c_double(0.0)
        with self.assertRaises(
            COMError,
            msg="CHARACTERIZATION: the dispinterface rejects a pointer(c_double) "
            "for an [in, out] DATE.",
        ):
            self._create().InitDate(pointer(value))


if __name__ == "__main__":
    unittest.main()
