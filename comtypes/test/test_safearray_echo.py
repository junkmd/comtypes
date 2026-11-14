"""Characterization tests for SAFEARRAY(VARIANT) marshaling behavior.

These tests observe and fix the current behavior of how comtypes handles
SAFEARRAY(VARIANT) across various marshaling paths:
- SAFEARRAY(VARIANT) echo
- Get VARTYPE of each element in a SAFEARRAY(VARIANT)

This is a characterization test suite: it records the CURRENT behavior,
not the desired behavior. If the current behavior seems wrong, the test
still asserts the observed values. Docstrings explain what is being
fixed and why it matters.
"""

import datetime
import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.automation import VARIANT, VT_DATE, VT_R8
from comtypes.client import CreateObject, GetModule

# Load the test server type library
ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


def _create_safearray_echo_ifc():
    """Create the ISafearrayEchoTest interface (vtable calls)."""
    return CreateObject(
        "Comtypes.CoComtypesSafearrayEchoTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.ISafearrayEchoTest,
    )


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestSafeArrayVariantMarshaling(unittest.TestCase):
    """SAFEARRAY(VARIANT) marshaling tests using ISafearrayEchoTest."""

    def setUp(self):
        self.safearray_echo = _create_safearray_echo_ifc()

    def test_safearray_variant_roundtrip_mixed(self):
        """SAFEARRAY(VARIANT) に VT_DATE と VT_R8 を混在させる。
        現状: 要素は unwrapped されて返る (datetime/float の混在)。
        """
        arr = [VARIANT(), VARIANT()]
        arr[0].vt = VT_DATE
        arr[0]._.VT_R8 = 44031.5
        arr[1].vt = VT_R8
        arr[1]._.VT_R8 = 44031.5 + 1.0

        result = self.safearray_echo.EchoSafeArrayVariant(arr)
        self.assertIsInstance(result, (list, tuple))
        self.assertEqual(len(result), 2)
        # 要素は unwrapped される
        self.assertIsInstance(result[0], datetime.datetime)
        self.assertIsInstance(result[1], float)
        self.assertEqual(result[0], datetime.datetime(2020, 7, 19, 12, 0, 0))
        self.assertEqual(result[1], 44031.5 + 1.0)

    def test_safearray_variant_get_vts(self):
        """SAFEARRAY(VARIANT) の各要素の vt をサーバーに取らせる。"""
        arr = [VARIANT(), VARIANT()]
        arr[0].vt = VT_DATE
        arr[0]._.VT_R8 = 44031.5
        arr[1].vt = VT_R8
        arr[1]._.VT_R8 = 44031.5 + 1.0

        vts = self.safearray_echo.GetVtsFromSafeArrayVariant(arr)
        self.assertIsInstance(vts, (list, tuple))
        self.assertEqual(len(vts), 2)
        self.assertEqual(vts[0], VT_DATE)
        self.assertEqual(vts[1], VT_R8)


if __name__ == "__main__":
    unittest.main()
