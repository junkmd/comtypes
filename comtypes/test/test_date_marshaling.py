"""Characterization tests for VT_DATE / DATE marshaling behavior.

These tests observe and fix the current behavior of how comtypes handles
COM DATE (VT_DATE) and double (VT_R8) values across various marshaling paths:
- VARIANT value getter/setter
- Typed DATE arguments (vtable and dispatch)
- VARIANT arguments
- SAFEARRAY(DATE) and SAFEARRAY(VARIANT)
- Properties
- Record fields
- VT_BYREF|VT_DATE

This is a characterization test suite: it records the CURRENT behavior,
not the desired behavior. If the current behavior seems wrong, the test
still asserts the observed values. Docstrings explain what is being
fixed and why it matters.
"""

import ctypes
import datetime
import unittest
from ctypes import byref, c_double, pointer

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.automation import VARIANT, VT_BYREF, VT_DATE, VT_R8
from comtypes.client import CreateObject, GetModule

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
        "Comtypes.DateTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IDualDateTest,
    )


def _create_disp_ifc():
    """Create the dispinterface (Invoke/VARIANT calls)."""
    return CreateObject(
        "Comtypes.DateTest",
        clsctx=CLSCTX_LOCAL_SERVER,
        interface=ComtypesCppTestSrvLib.IDispDateTest,
    )


# Observed constants from test runs:
#
# Key finding: There's a ~31 day discrepancy in comtypes' DATE conversion.
# - Server returns 44000.5 (intended as 2020-07-19 12:00:00)
# - Python receives datetime(2020, 6, 18, 12, 0, 0) - 31 days earlier
# - Python's VARIANT(datetime(2020,7,19,12,0,0)) computes internal 44031.5
# - That 44031.5 also converts back to datetime(2020, 7, 19, 12, 0, 0) - the original!
# - But 44000.5 converts to datetime(2020, 6, 18, 12, 0, 0)
#
# This asymmetry is the CURRENT behavior we are characterizing.

SERVER_RETURN_VALUE = 44000.5  # What server sends
SERVER_RETURN_DATETIME_OBSERVED = datetime.datetime(
    2020, 6, 18, 12, 0, 0
)  # What Python actually receives from server

# When Python creates VARIANT from datetime(2020,7,19,12,0,0):
VARIANT_DATETIME_INPUT = datetime.datetime(2020, 7, 19, 12, 0, 0)
VARIANT_INTERNAL_VALUE = (
    44031.5  # What comtypes computes internally for the above datetime
)
VARIANT_ROUNDTRIP_DATETIME = datetime.datetime(
    2020, 7, 19, 12, 0, 0
)  # What comes back when manually setting 44031.5


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestVariantDateBehavior(unittest.TestCase):
    """観点A: VARIANT 単体 (テストダブル不要)

    VARIANT の value プロパティを通じた datetime/float の扱いを観測する。
    言語境界を越えない純粋な Python 側の挙動。
    """

    def test_variant_from_datetime_sets_vt_date(self):
        """VARIANT(datetime) は vt=VT_DATE になり、value は datetime を返す。"""
        dt = VARIANT_DATETIME_INPUT
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        self.assertIsInstance(v.value, datetime.datetime)

    def test_variant_from_float_sets_vt_r8(self):
        """VARIANT(float) は vt=VT_R8 になり、value は float を返す。"""
        v = VARIANT(SERVER_RETURN_VALUE)
        self.assertEqual(v.vt, VT_R8)
        self.assertEqual(v.value, SERVER_RETURN_VALUE)
        self.assertIsInstance(v.value, float)

    def test_variant_manual_vt_date_value_returns_datetime(self):
        """手動で vt=VT_DATE を設定した VARIANT の value は datetime を返す。

        現状の挙動: 内部値 44031.5 は datetime(2020, 7, 19, 12, 0, 0) に変換される。
        これは元の datetime と一致する (forward/backward は 44031.5 で一貫)。
        """
        v = VARIANT()
        v.vt = VT_DATE
        v._.VT_R8 = VARIANT_INTERNAL_VALUE
        self.assertEqual(v.vt, VT_DATE)
        self.assertIsInstance(v.value, datetime.datetime)
        self.assertEqual(v.value, VARIANT_ROUNDTRIP_DATETIME)

    def test_variant_manual_vt_r8_value_returns_float(self):
        """手動で vt=VT_R8 を設定した VARIANT の value は float を返す。"""
        v = VARIANT()
        v.vt = VT_R8
        v._.VT_R8 = VARIANT_INTERNAL_VALUE
        self.assertEqual(v.vt, VT_R8)
        self.assertIsInstance(v.value, float)
        self.assertEqual(v.value, VARIANT_INTERNAL_VALUE)

    def test_variant_date_only_datetime_date_raises(self):
        """datetime.date (時刻なし) を渡すと TypeError になる (現状の挙動)。"""
        d = datetime.date(2020, 7, 19)
        with self.assertRaises(TypeError):
            VARIANT(d)

    def test_variant_tz_aware_datetime_raises(self):
        """tz-aware datetime を渡すと TypeError になる (naive/aware subtraction エラー)。"""
        dt = datetime.datetime(2020, 7, 19, 12, 0, 0, tzinfo=datetime.timezone.utc)
        with self.assertRaises(TypeError):
            VARIANT(dt)

    def test_variant_datetime_time_raises(self):
        """datetime.time を渡すと TypeError になる。"""
        t = datetime.time(12, 0, 0)
        with self.assertRaises(TypeError):
            VARIANT(t)

    def test_variant_datetime_timedelta_raises(self):
        """datetime.timedelta を渡すと TypeError になる。"""
        td = datetime.timedelta(days=1)
        with self.assertRaises(TypeError):
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
        dt = VARIANT_DATETIME_INPUT
        v = VARIANT(dt)
        self.assertEqual(v.vt, VT_DATE)
        self.assertEqual(v.value, dt)
        # 観測された内部値
        self.assertEqual(v._.VT_R8, VARIANT_INTERNAL_VALUE)

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
        v.value = VARIANT_DATETIME_INPUT
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
        result = self.ifc.DateRoundTrip(SERVER_RETURN_VALUE)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_date_roundtrip_int_in(self):
        """[in] DATE に int を渡せる (float に変換される)。"""
        result = self.ifc.DateRoundTrip(44000)
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
            self.ifc.DateRoundTrip(VARIANT_DATETIME_INPUT)

    def test_date_out_returns_float(self):
        """[out,retval] DATE でサーバーが作った値は float で返る。

        サーバーは 44000.5 を返すが、vtable 経由では float として受け取られる。
        """
        result = self.ifc.DateOut()
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_double_roundtrip_comparison(self):
        """対比: [in] double / [out,retval] double の挙動 (同じく float)。"""
        result = self.ifc.DoubleRoundTrip(SERVER_RETURN_VALUE)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_double_out_returns_float(self):
        """[out,retval] double は float を返す。"""
        result = self.ifc.DoubleOut()
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestTypedDateDispatch(unittest.TestCase):
    """観点C: 動的 dispatch (Invoke 経由)

    同じメソッドを dispinterface 経由で呼び、VARIANT 受け渡しでの
    挙動を観測。サーバー側が受け取った vt と生の double 値も観測。
    """

    def setUp(self):
        self.ifc = _create_disp_ifc()

    def test_dispatch_date_roundtrip_float(self):
        """dispatch 経由で float を渡すとサーバーは VT_R8 を受け取る。

        戻り値は datetime になる (VARIANT 経由で VT_DATE として返されるため)。
        観測値: datetime(2020, 6, 18, 12, 0, 0) - 31 日のズレあり。
        """
        result = self.ifc.DateRoundTrip(SERVER_RETURN_VALUE)
        # 現状: float を渡しても datetime で返ってくる (ズレあり)
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(
            result,
            SERVER_RETURN_DATETIME_OBSERVED,
            msg="CHARACTERIZATION(#519): "
            "If https://github.com/enthought/comtypes/issues/519 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        )

        # サーバーが受け取った vt を観測
        vt = self.ifc.GetVariantVt(VARIANT(SERVER_RETURN_VALUE))
        # 現状: float を渡すと VT_R8 として受け取られる
        self.assertEqual(vt, VT_R8)

    def test_dispatch_date_roundtrip_datetime(self):
        """dispatch 経由で datetime を渡すとサーバーは VT_DATE を受け取る。"""
        result = self.ifc.DateRoundTrip(VARIANT_DATETIME_INPUT)
        self.assertIsInstance(result, datetime.datetime)
        # 戻り値は入力と同じ datetime (ただし内部的にはズレがある)
        self.assertEqual(result, VARIANT_DATETIME_INPUT)

        vt = self.ifc.GetVariantVt(VARIANT(VARIANT_DATETIME_INPUT))
        # 現状: datetime を渡すと VT_DATE として受け取られる
        self.assertEqual(vt, VT_DATE)

    def test_dispatch_get_variant_date_value_datetime(self):
        """サーバーが受け取った VT_DATE の生の double 値を観測。"""
        raw = self.ifc.GetVariantDateValue(VARIANT(VARIANT_DATETIME_INPUT))
        self.assertIsInstance(raw, float)
        # 観測値: 44031.5 (Python 側で計算される内部値)
        self.assertEqual(raw, VARIANT_INTERNAL_VALUE)

    def test_dispatch_get_variant_date_value_float(self):
        """サーバーが受け取った VT_R8 の生の double 値を観測。"""
        raw = self.ifc.GetVariantDateValue(VARIANT(SERVER_RETURN_VALUE))
        self.assertIsInstance(raw, float)
        self.assertEqual(raw, SERVER_RETURN_VALUE)

    def test_dispatch_date_out(self):
        """dispatch 経由の [out,retval] DATE - datetime で返る (ズレあり)。"""
        result = self.ifc.DateOut()
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(
            result,
            SERVER_RETURN_DATETIME_OBSERVED,
            msg="CHARACTERIZATION(#519): "
            "If https://github.com/enthought/comtypes/issues/519 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        )

    def test_dispatch_double_roundtrip(self):
        """対比: dispatch 経由の double - float で返る。"""
        result = self.ifc.DoubleRoundTrip(SERVER_RETURN_VALUE)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestVariantArgumentMarshaling(unittest.TestCase):
    """観点D: VARIANT 型の引数・戻り値

    VARIANT を引数として渡す/受け取る場合の vt と値の扱い。
    現状: VARIANT 往復メソッドは、受け取った VARIANT の .value (unwrapped) を返す。
    """

    def setUp(self):
        self.dual = _create_dual_ifc()
        self.disp = _create_disp_ifc()

    def test_variant_roundtrip_datetime_returns_datetime(self):
        """VT_DATE の VARIANT を往復させると datetime が返る (unwrapped)。

        手動で vt=VT_DATE, value=44031.5 を設定した VARIANT を渡すと、
        戻り値は datetime(2020, 7, 19, 12, 0, 0) になる (44031.5 の逆変換)。
        """
        v = VARIANT()
        v.vt = VT_DATE
        v._.VT_R8 = VARIANT_INTERNAL_VALUE
        result = self.dual.VariantRoundTrip(v)
        # 現状: VARIANT ではなく unwrapped value が返る
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(result, VARIANT_ROUNDTRIP_DATETIME)

    def test_variant_roundtrip_float_returns_float(self):
        """VT_R8 の VARIANT を往復させると float が返る (unwrapped)。"""
        v = VARIANT()
        v.vt = VT_R8
        v._.VT_R8 = SERVER_RETURN_VALUE
        result = self.dual.VariantRoundTrip(v)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_server_creates_variant_date_returns_datetime(self):
        """サーバーが作る VT_DATE VARIANT は datetime で返る (unwrapped, ズレあり)。

        サーバーは 44000.5 で VT_DATE を作るが、Python では
        datetime(2020, 6, 18, 12, 0, 0) として受け取られる。
        """
        result = self.dual.CreateVariantDate(SERVER_RETURN_VALUE)
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(
            result,
            SERVER_RETURN_DATETIME_OBSERVED,
            msg="CHARACTERIZATION(#519): "
            "If https://github.com/enthought/comtypes/issues/519 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        )

    def test_server_creates_variant_r8_returns_float(self):
        """サーバーが作る VT_R8 VARIANT は float で返る (unwrapped)。"""
        result = self.dual.CreateVariantR8(SERVER_RETURN_VALUE)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_dispatch_pass_datetime_as_variant_returns_datetime(self):
        """dispatch 経由で datetime を VARIANT 引数に渡すと datetime が返る。"""
        v = VARIANT(VARIANT_DATETIME_INPUT)
        result = self.disp.VariantRoundTrip(v)
        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(result, VARIANT_DATETIME_INPUT)

    def test_dispatch_pass_float_as_variant_returns_float(self):
        """dispatch 経由で float を VARIANT 引数に渡すと float が返る。"""
        v = VARIANT(SERVER_RETURN_VALUE)
        result = self.disp.VariantRoundTrip(v)
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)


@unittest.skipIf(IMPORT_FAILED, "Requires the C++ out-of-process COM server.")
class TestSafeArrayDateMarshaling(unittest.TestCase):
    """観点E: SAFEARRAY

    SAFEARRAY(DATE) と SAFEARRAY(VARIANT) の要素の型・値の扱い。
    現状: SAFEARRAY(DATE) は float (c_double) 配列として扱われる。
    """

    def setUp(self):
        self.dual = _create_dual_ifc()
        self.disp = _create_disp_ifc()

    def test_safearray_date_roundtrip_float_list(self):
        """SAFEARRAY(DATE) に float リストを渡すと float リストが返る。"""
        dates = [SERVER_RETURN_VALUE, SERVER_RETURN_VALUE + 1.0]
        result = self.dual.SafeArrayDateRoundTrip(dates)
        self.assertIsInstance(result, (list, tuple))
        self.assertEqual(len(result), 2)
        for elem in result:
            self.assertIsInstance(elem, float)
        self.assertEqual(result[0], SERVER_RETURN_VALUE)
        self.assertEqual(result[1], SERVER_RETURN_VALUE + 1.0)

    def test_safearray_date_roundtrip_datetime_list_raises(self):
        """SAFEARRAY(DATE) に datetime リストを渡すと ctypes.ArgumentError。"""
        dates = [
            VARIANT_DATETIME_INPUT,
            VARIANT_DATETIME_INPUT + datetime.timedelta(days=1),
        ]
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            self.dual.SafeArrayDateRoundTrip(dates)

    def test_safearray_variant_roundtrip_mixed(self):
        """SAFEARRAY(VARIANT) に VT_DATE と VT_R8 を混在させる。
        現状: 要素は unwrapped されて返る (datetime/float の混在)。
        """
        arr = [VARIANT(), VARIANT()]
        arr[0].vt = VT_DATE
        arr[0]._.VT_R8 = VARIANT_INTERNAL_VALUE
        arr[1].vt = VT_R8
        arr[1]._.VT_R8 = SERVER_RETURN_VALUE + 1.0

        result = self.dual.SafeArrayVariantRoundTrip(arr)
        self.assertIsInstance(result, (list, tuple))
        self.assertEqual(len(result), 2)
        # 要素は unwrapped される
        self.assertIsInstance(result[0], datetime.datetime)
        self.assertIsInstance(result[1], float)
        self.assertEqual(result[0], VARIANT_ROUNDTRIP_DATETIME)
        self.assertEqual(result[1], SERVER_RETURN_VALUE + 1.0)

    def test_safearray_variant_get_vts(self):
        """SAFEARRAY(VARIANT) の各要素の vt をサーバーに取らせる。"""
        arr = [VARIANT(), VARIANT()]
        arr[0].vt = VT_DATE
        arr[0]._.VT_R8 = VARIANT_INTERNAL_VALUE
        arr[1].vt = VT_R8
        arr[1]._.VT_R8 = SERVER_RETURN_VALUE + 1.0

        vts = self.dual.SafeArrayVariantGetVts(arr)
        self.assertIsInstance(vts, (list, tuple))
        self.assertEqual(len(vts), 2)
        self.assertEqual(vts[0], VT_DATE)
        self.assertEqual(vts[1], VT_R8)

    def test_variant_with_vt_array_date(self):
        """VARIANT に VT_ARRAY|VT_DATE を入れたときの .value。"""
        dates = [SERVER_RETURN_VALUE, SERVER_RETURN_VALUE + 1.0]
        v = VARIANT(dates)
        self.assertTrue(v.vt & 0x2000)  # VT_ARRAY bit
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
        self.dual.DateProp = SERVER_RETURN_VALUE
        result = self.dual.DateProp
        self.assertIsInstance(result, float)
        self.assertEqual(result, SERVER_RETURN_VALUE)

    def test_date_property_set_datetime_raises(self):
        """DATE プロパティに datetime を set しようとすると ctypes.ArgumentError。"""
        with self.assertRaises(
            ctypes.ArgumentError,
            msg="CHARACTERIZATION(#152): "
            "If https://github.com/enthought/comtypes/issues/152 is fixed, "
            "this failure is expected; change it to a normal-path test.",
        ):
            self.dual.DateProp = VARIANT_DATETIME_INPUT

    def test_record_date_field_roundtrip(self):
        """DATE フィールドを持つ構造体の往復 - フィールドは float として扱われる。"""
        rec = ComtypesCppTestSrvLib.StructDateRecord()
        rec.dateField = SERVER_RETURN_VALUE
        rec.id = 42

        result = self.dual.RecordDateRoundTrip(rec)
        self.assertEqual(result.id, 42)
        # dateField は float として扱われる
        self.assertIsInstance(result.dateField, float)
        self.assertEqual(result.dateField, SERVER_RETURN_VALUE)

    def test_byref_date_variant(self):
        """VT_BYREF|VT_DATE を持つ VARIANT の書き戻し。
        現状: null reference pointer エラーになる (未実装/バグ)。
        """
        v = VARIANT()
        v.vt = VT_BYREF | VT_DATE
        date_val = c_double(SERVER_RETURN_VALUE)
        v._.pdate = pointer(date_val)

        # 現状では COMError (null reference pointer) になる
        with self.assertRaises(ctypes.COMError):
            self.dual.ByRefDateVariant(byref(v))


if __name__ == "__main__":
    unittest.main()
