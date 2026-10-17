import ctypes
import unittest

from comtypes import automation, typeinfo
from comtypes.tools import tlbparser


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
