"""Characterization tests for VT_DATE / DATE marshaling behavior.

These tests observe and fix the current behavior of how comtypes handles
COM DATE (VT_DATE) and double (VT_R8) values across various marshaling paths:
- VARIANT value getter/setter
- Typed DATE arguments (vtable and dispatch)
- VARIANT arguments
- SAFEARRAY(DATE)
- Properties
- Record fields
- VT_BYREF|VT_DATE

SAFEARRAY(VARIANT) tests are in test_safearray_echo.py.

This is a characterization test suite: it records the CURRENT behavior,
not the desired behavior. If the current behavior seems wrong, the test
still asserts the observed values. Docstrings explain what is being
fixed and why it matters.
"""

import ctypes
import datetime
import unittest
from ctypes import byref, c_double, memmove, sizeof

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.automation import VARIANT, VT_ARRAY, VT_BYREF, VT_DATE, VT_R8
from comtypes.client import CreateObject, GetModule
from comtypes.safearray import _midlSAFEARRAY

# Load the test server type library
ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


def _create_dual_ifc():
    """Create the dual interface (vtable calls)."""
    return CreateObject(
        "Comtypes.CoComtypesDateCompatibilityTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IDualDateTest,
    )


def _create_disp_ifc():
    """Create the dispinterface (Invoke/VARIANT calls)."""
    return CreateObject(
        "Comtypes.CoComtypesDateCompatibilityTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IDispDateTest,
    )


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestVariantDateBehavior(unittest.TestCase):
    """観点A: VARIANT 単体 (テストダブル不要)

    VARIANT の value プロパティを通じた datetime/float の扱いを観測する。
    言語境界を越えない純粋な Python 側の挙動。
    """

    def test_variant_from_datetime_sets_vt_date(self):
        """VARIANT(datetime) は vt=VT_DATE になり、value は datetime を返す。"""
        dt = datetime.datetime(2020, 7, 19, 12, 0, 0)
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        self.assertIsInstance(v.value, datetime.datetime)

    def test_variant_from_float_sets_vt_r8(self):
        """VARIANT(float) は vt=VT_R8 になり、value は float を返す。"""
        v = VARIANT(44031.5)
        self.assertEqual(v.vt, VT_R8)
        self.assertEqual(v.value, 44031.5)
        self.assertIsInstance(v.value, float)

    def test_variant_manual_vt_date_value_returns_datetime(self):
        """手動で vt=VT_DATE を設定した VARIANT の value は datetime を返す。

        現状の挙動: 内部値 44031.5 は datetime(2020, 7, 19, 12, 0, 0) に変換される。
        これは元の datetime と一致する (forward/backward は 44031.5 で一貫)。
        """
        v = VARIANT()
        v.vt = VT_DATE
        v._.VT_R8 = 44031.5
        self.assertEqual(v.vt, VT_DATE)
        self.assertIsInstance(v.value, datetime.datetime)
        self.assertEqual(v.value, datetime.datetime(2020, 7, 19, 12, 0, 0))

    def test_variant_manual_vt_r8_value_returns_float(self):
        """手動で vt=VT_R8 を設定した VARIANT の value は float を返す。"""
        v = VARIANT()
        v.vt = VT_R8
        v._.VT_R8 = 44031.5
        self.assertEqual(v.vt, VT_R8)
        self.assertIsInstance(v.value, float)
        self.assertEqual(v.value, 44031.5)

    def test_variant_date_only_datetime_date_raises(self):
        """datetime.date (時刻なし) を渡すと TypeError になる (現状の挙動)。"""
        d = datetime.date(2020, 7, 19)
        with self.assertRaises(
            TypeError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            VARIANT(d)

    def test_variant_tz_aware_datetime_raises(self):
        """tz-aware datetime を渡すと TypeError になる (naive/aware subtraction エラー)。"""
        dt = datetime.datetime(2020, 7, 19, 12, 0, 0, tzinfo=datetime.timezone.utc)
        with self.assertRaises(
            TypeError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            VARIANT(dt)

    def test_variant_datetime_time_raises(self):
        """datetime.time を渡すと TypeError になる。"""
        t = datetime.time(12, 0, 0)
        with self.assertRaises(
            TypeError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            VARIANT(t)

    def test_variant_datetime_timedelta_raises(self):
        """datetime.timedelta を渡すと TypeError になる。"""
        td = datetime.timedelta(days=1)
        with self.assertRaises(
            TypeError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            VARIANT(td)

    def test_variant_boundary_zero_date(self):
        """DATE の 0 (1899-12-30) の境界値。"""
        dt = datetime.datetime(1899, 12, 30, 0, 0, 0)
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        # 内部の double 値は 0.0
        self.assertEqual(v._.VT_R8, 0.0)

    def test_variant_boundary_negative_date(self):
        """DATE の負の値 (1899-12-30 より前)。"""
        dt = datetime.datetime(1899, 12, 29, 0, 0, 0)  # -1.0
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        self.assertEqual(v._.VT_R8, -1.0)

    def test_variant_boundary_fractional_day(self):
        """日付の小数部 (正午 = 0.5) - 観測値を固定。"""
        dt = datetime.datetime(2020, 7, 19, 12, 0, 0)
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        # 観測された内部値
        self.assertEqual(v._.VT_R8, 44031.5)

    def test_variant_boundary_microsecond_preserved(self):
        """マイクロ秒は丸められず保持される (現状の挙動)。"""
        dt = datetime.datetime(2020, 7, 19, 12, 0, 0, 1)
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        # マイクロ秒が保持される
        self.assertEqual(v.value, datetime.datetime(2020, 7, 19, 12, 0, 0, 1))

    def test_variant_boundary_datetime_min_accepted(self):
        """datetime.min (年 1) は例外にならず受け入れられる (現状の挙動)。"""
        # 現状では例外にならない
        v = VARIANT(datetime.datetime.min)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, datetime.datetime.min)

    def test_variant_boundary_datetime_max_overflow(self):
        """datetime.max (年 9999) は _get_value で OverflowError になる。"""
        v = VARIANT(datetime.datetime.max)
        self.assertEqual(v.vt, VT_DATE)
        # value アクセス時に OverflowError
        with self.assertRaises(OverflowError):
            _ = v.value

    def test_variant_value_reassignment_changes_vt(self):
        """VARIANT.value を再設定すると vt が変わる。"""
        v = VARIANT()
        v.value = datetime.datetime(2020, 7, 19, 12, 0, 0)
        self.assertEqual(v.vt, VT_DATE)

        v.value = 3.14
        self.assertEqual(v.vt, VT_R8)
        self.assertEqual(v.value, 3.14)

        v.value = "string"
        self.assertEqual(v.vt, 8)  # VT_BSTR
        self.assertEqual(v.value, "string")


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestTypedDateVtable(unittest.TestCase):
    """観点B: 型付き DATE 引数 (vtable 経由、言語境界あり)

    dual インターフェースの vtable 経由で呼び出されるメソッド。
    Python の datetime/float がどうマーシャリングされるかを観測。
    現状: ctypes は DATE (c_double) を期待するため、float のみ受け付ける。
    """

    def setUp(self):
        self.ifc = _create_dual_ifc()

    def test_date_roundtrip_float_in_float_out(self):
        """[in] DATE に float を渡し、[out,retval] DATE で float が返る。"""
        result = self.ifc.EchoDate(44031.5)
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)

    def test_date_roundtrip_int_in(self):
        """[in] DATE に int を渡せる (float に変換される)。"""
        result = self.ifc.EchoDate(44000)
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44000.0)

    def test_date_roundtrip_datetime_raises(self):
        """[in] DATE に datetime を渡すと ctypes.ArgumentError になる。"""
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            self.ifc.EchoDate(datetime.datetime(2020, 7, 19, 12, 0, 0))

    def test_date_out_returns_float(self):
        """[out,retval] DATE でサーバーが作った値は float で返る。

        サーバーは 44031.5 (正しい OLE date) を返し、vtable 経由では float として受け取られる。
        """
        result = self.ifc.GetDate44031dot5()
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)

    def test_double_roundtrip_comparison(self):
        """対比: [in] double / [out,retval] double の挙動 (同じく float)。"""
        result = self.ifc.EchoDouble(44031.5)
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)

    def test_double_out_returns_float(self):
        """[out,retval] double は float を返す (44031.5)。"""
        result = self.ifc.GetDouble44031dot5()
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestSafeArrayDateMarshaling(unittest.TestCase):
    """観点E: SAFEARRAY

    SAFEARRAY(DATE) の要素の型・値の扱い。
    現状: SAFEARRAY(DATE) は float (c_double) 配列として扱われる。
    SAFEARRAY(VARIANT) のテストは test_safearray_echo.py に移動済み。
    """

    def setUp(self):
        self.dual = _create_dual_ifc()
        self.disp = _create_disp_ifc()

    def test_safearray_date_roundtrip_float_list(self):
        """SAFEARRAY(DATE) に float リストを渡すと float リストが返る。"""
        dates = [44031.5, 44031.5 + 1.0]
        result = self.dual.EchoSafeArrayDate(dates)
        self.assertIsInstance(result, (list, tuple))
        self.assertEqual(len(result), 2)
        for elem in result:
            self.assertIsInstance(
                elem,
                float,
                msg="CHARACTERIZATION(#236): "
                "If https://github.com/enthought/comtypes/issues/236 is fixed, "
                "this failure is expected; change it to a normal-path test.",
            )
        self.assertEqual(result[0], 44031.5)
        self.assertEqual(result[1], 44031.5 + 1.0)

    def test_safearray_date_roundtrip_datetime_list_raises(self):
        """SAFEARRAY(DATE) に datetime リストを渡すと ctypes.ArgumentError。"""
        dates = [
            datetime.datetime(2020, 7, 19, 12, 0, 0),
            datetime.datetime(2020, 7, 19, 12, 0, 0) + datetime.timedelta(days=1),
        ]
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            self.dual.EchoSafeArrayDate(dates)

    def test_variant_with_vt_array_date(self):
        """VARIANT に VT_ARRAY|VT_DATE を入れたときの .value。"""
        # SAFEARRAY(DATE) を作成して VARIANT に設定
        sa = _midlSAFEARRAY(c_double).create([44031.5, 44031.5 + 1.0])
        v = VARIANT()
        memmove(byref(v._), byref(sa), sizeof(sa))
        v.vt = VT_ARRAY | VT_DATE

        self.assertTrue(v.vt & 0x2000)  # VT_ARRAY bit
        self.assertEqual(v.vt & 0xFFF, VT_DATE)  # VT_DATE bit
        self.assertIsInstance(v.value, (list, tuple))
        # 要素は float
        for elem in v.value:
            self.assertIsInstance(
                elem,
                float,
                msg="CHARACTERIZATION(#236): "
                "If https://github.com/enthought/comtypes/issues/236 is fixed, "
                "this failure is expected; change it to a normal-path test.",
            )


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestPropertyAndRecordAndByRef(unittest.TestCase):
    """観点F: プロパティ、構造体フィールド、byref

    DATE 型のプロパティ get/put、構造体フィールド、VT_BYREF|VT_DATE。
    """

    def setUp(self):
        self.dual = _create_dual_ifc()
        self.disp = _create_disp_ifc()

    def test_date_property_get(self):
        """DATE プロパティの get - float で返る。"""
        result = self.dual.DateProp
        self.assertIsInstance(result, float)
        # 初期値は 0.0
        self.assertEqual(result, 0.0)

    def test_date_property_set_float(self):
        """DATE プロパティに float を set できる。"""
        self.dual.DateProp = 44031.5
        result = self.dual.DateProp
        self.assertIsInstance(result, float)
        self.assertEqual(result, 44031.5)

    def test_date_property_set_datetime_raises(self):
        """DATE プロパティに datetime を set しようとすると ctypes.ArgumentError。"""
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            self.dual.DateProp = datetime.datetime(2020, 7, 19, 12, 0, 0)

    def test_record_date_field_roundtrip(self):
        """DATE フィールドを持つ構造体の往復 - フィールドは float として扱われる。"""
        rec = ComtypesCppTestSrvLib.StructDateRecord()
        rec.dateField = 44031.5
        rec.id = 42

        result = self.dual.EchoRecordDate(rec)
        self.assertEqual(result.id, 42)
        # dateField は float として扱われる
        self.assertIsInstance(result.dateField, float)
        self.assertEqual(result.dateField, 44031.5)

    def test_byref_date_variant(self):
        """VT_BYREF|VT_DATE を持つ VARIANT の書き戻し。
        現状: null reference pointer エラーになる (未実装/バグ)。
        """
        v = VARIANT()
        v.vt = VT_BYREF | VT_DATE
        # VARIANT の union には pdate フィールドはなく、VT_R8 (c_double) を使う
        v._.VT_R8 = 44031.5

        # 現状では COMError (null reference pointer) になる
        with self.assertRaises(ctypes.COMError):
            self.dual.AddOneToByRefDateVariant(byref(v))


if __name__ == "__main__":
    unittest.main()
