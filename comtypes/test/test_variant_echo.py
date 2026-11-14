"""Characterization tests for VARIANT round-trip behavior using IVariantEchoTest.

These tests observe and fix the current behavior of how comtypes handles
VARIANT round-trip across various marshaling paths:
- VARIANT echo (EchoVariant)
- GetVariantVt (observes what VT the server sees)

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


def _create_variant_echo_ifc():
    """Create the IVariantEchoTest interface (vtable calls)."""
    return CreateObject(
        "Comtypes.CoComtypesVariantEchoTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IVariantEchoTest,
    )


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestVariantEcho(unittest.TestCase):
    def setUp(self):
        self.variant_echo = _create_variant_echo_ifc()

    def test_dispatch_date_roundtrip_float(self):
        """dispatch 経由で float を渡すとサーバーは VT_R8 を受け取る。

        戻り値は datetime になる (VARIANT 経由で VT_DATE として返されるため)。
        サーバーは正しい OLE date (44031.5 = 2020-07-19 12:00:00) を返す。
        """
        # サーバーが受け取った vt を観測
        vt = self.variant_echo.GetVariantVt(VARIANT(44031.5))
        # float を渡すと VT_R8 として受け取られる
        self.assertEqual(vt, VT_R8)

    def test_dispatch_date_roundtrip_datetime(self):
        """dispatch 経由で datetime を渡すとサーバーは VT_DATE を受け取る。"""
        vt = self.variant_echo.GetVariantVt(
            VARIANT(datetime.datetime(2020, 7, 19, 12, 0, 0))
        )
        # datetime を渡すと VT_DATE として受け取られる
        self.assertEqual(vt, VT_DATE)

    def test_variant_roundtrip_datetime_returns_datetime(self):
        """VT_DATE の VARIANT を往復させると datetime が返る (unwrapped)。

        手動で vt=VT_DATE, value=44031.5 を設定した VARIANT を渡すと、
        戻り値は datetime(2020, 7, 19, 12, 0, 0) になる (44031.5 の逆変換)。
        """
        v = VARIANT()
        v.vt = VT_DATE
        v._.VT_R8 = 44031.5
        result = self.variant_echo.EchoVariant(v)
        # 現状: VARIANT ではなく unwrapped value が返る
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(result, datetime.datetime(2020, 7, 19, 12, 0, 0))

    def test_variant_roundtrip_float_returns_float(self):
        """VT_R8 の VARIANT を往復させると float が返る (unwrapped)。"""
        v = VARIANT()
        v.vt = VT_R8
        v._.VT_R8 = 44031.5
        result = self.variant_echo.EchoVariant(v)
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)

    def test_pass_datetime_as_variant_returns_datetime(self):
        v = VARIANT(datetime.datetime(2020, 7, 19, 12, 0, 0))
        result = self.variant_echo.EchoVariant(v)
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(result, datetime.datetime(2020, 7, 19, 12, 0, 0))

    def test_pass_float_as_variant_returns_float(self):
        v = VARIANT(44031.5)
        result = self.variant_echo.EchoVariant(v)
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)


if __name__ == "__main__":
    unittest.main()
