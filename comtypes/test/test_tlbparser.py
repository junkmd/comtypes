import ctypes
import unittest
from _ctypes import COMError
from unittest.mock import Mock

from comtypes import GUID, automation, typeinfo
from comtypes.tools import tlbparser


class TestTypeLibraryDescription(unittest.TestCase):
    def _create_tlib_mock(self):
        libattr = Mock(
            spec=typeinfo.TLIBATTR,
            guid=GUID("{12345678-1234-1234-1234-123456789ABC}"),
            lcid=0,
            wMajorVerNum=1,
            wMinorVerNum=2,
        )
        return Mock(spec=typeinfo.ITypeLib, GetLibAttr=Mock(return_value=libattr))

    def test_includes_name_and_version(self):
        tlib = self._create_tlib_mock()
        tlib.GetDocumentation.return_value = ("MyLib", "", "", 0)
        self.assertEqual(
            tlbparser._tlib_description(tlib),
            "name='MyLib', libid='{12345678-1234-1234-1234-123456789ABC}', lcid=0, major=1, minor=2",
        )

    def test_falls_back_to_guid_when_name_is_none(self):
        tlib = self._create_tlib_mock()
        tlib.GetDocumentation.return_value = (None, "", "", 0)
        self.assertEqual(
            tlbparser._tlib_description(tlib),
            "libid='{12345678-1234-1234-1234-123456789ABC}', lcid=0, major=1, minor=2",
        )

    def test_falls_back_to_guid_when_documentation_raises_comerror(self):
        tlib = self._create_tlib_mock()
        tlib.GetDocumentation.side_effect = COMError(-2147024894, "boom", (None,) * 5)
        self.assertEqual(
            tlbparser._tlib_description(tlib),
            "libid='{12345678-1234-1234-1234-123456789ABC}', lcid=0, major=1, minor=2",
        )


class TestUsesVT_I1(unittest.TestCase):
    def test_direct_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_I1)
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_byref_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_BYREF | automation.VT_I1)
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_array_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_ARRAY | automation.VT_I1)
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_ptr_to_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_PTR)
        tdesc._.lptdesc = ctypes.POINTER(typeinfo.TYPEDESC)(
            typeinfo.TYPEDESC(vt=automation.VT_I1)
        )
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_safearray_of_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_SAFEARRAY)
        tdesc._.lptdesc = ctypes.POINTER(typeinfo.TYPEDESC)(
            typeinfo.TYPEDESC(vt=automation.VT_I1)
        )
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_carray_of_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_CARRAY)
        arraydesc = typeinfo.tagARRAYDESC()
        arraydesc.tdescElem = typeinfo.TYPEDESC(vt=automation.VT_I1)
        tdesc._.lpadesc = ctypes.POINTER(typeinfo.tagARRAYDESC)(arraydesc)
        self.assertTrue(tlbparser._uses_vt_i1(tdesc))

    def test_userdefined_alias_to_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_USERDEFINED)
        # 7 is only a placeholder reference ID for this mock.
        # It is not a flag, and any valid HREFTYPE value would work
        # as long as the mock resolves the alias to a type whose
        # TYPEDESC is VT_I1.
        tdesc._.hreftype = 7
        target_tinfo = Mock(spec=typeinfo.ITypeInfo)
        target_ta = Mock(spec=typeinfo.TYPEATTR)
        target_ta.tdescAlias = typeinfo.TYPEDESC(vt=automation.VT_I1)
        target_tinfo.GetTypeAttr.return_value = target_ta
        tinfo = Mock(spec=typeinfo.ITypeInfo)
        tinfo.GetRefTypeInfo.return_value = target_tinfo
        self.assertTrue(tlbparser._uses_vt_i1(tdesc, tinfo))

    def test_byref_userdefined_alias_to_vt_i1_is_detected(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_BYREF | automation.VT_USERDEFINED)
        # 7 is only a placeholder reference ID for this mock.
        # It is not a flag, and any valid HREFTYPE value would work
        # as long as the mock resolves the alias to a type whose
        # TYPEDESC is VT_I1.
        tdesc._.hreftype = 7
        target_tinfo = Mock(spec=typeinfo.ITypeInfo)
        target_ta = Mock(spec=typeinfo.TYPEATTR)
        target_ta.tdescAlias = typeinfo.TYPEDESC(vt=automation.VT_I1)
        target_tinfo.GetTypeAttr.return_value = target_ta
        tinfo = Mock(spec=typeinfo.ITypeInfo)
        tinfo.GetRefTypeInfo.return_value = target_tinfo
        self.assertTrue(tlbparser._uses_vt_i1(tdesc, tinfo))

    def test_variant_is_not_treated_as_vt_i1(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_VARIANT)
        self.assertFalse(tlbparser._uses_vt_i1(tdesc))

    def test_non_i1_pointer_is_not_treated_as_vt_i1(self):
        tdesc = typeinfo.TYPEDESC(vt=automation.VT_PTR)
        tdesc._.lptdesc = ctypes.POINTER(typeinfo.TYPEDESC)(
            typeinfo.TYPEDESC(vt=automation.VT_I2)
        )
        self.assertFalse(tlbparser._uses_vt_i1(tdesc))


if __name__ == "__main__":
    unittest.main()
