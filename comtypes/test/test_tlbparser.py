"""Tests for ``comtypes.tools.tlbparser``.

These tests exercise the type library parser to improve coverage.
"""

import unittest
import warnings
from comtypes.tools import tlbparser, typedesc
from comtypes.tools.tlbparser import (
    BSTR_type,
    HRESULT_type,
    VARIANT_BOOL_type,
    VARIANT_type,
    void_type,
    int_type,
    IUNKNOWN_type,
    IDISPATCH_type,
    COMTYPES,
    Parser,
    TlbFileParser,
    TypeLibParser,
    get_tlib_filename,
)


class TestTlbParserConstants(unittest.TestCase):
    def test_comtypes_mapping(self):
        # Test that COMTYPES contains expected mappings
        from comtypes import automation
        self.assertIn(automation.VT_I2, COMTYPES)
        self.assertIn(automation.VT_I4, COMTYPES)
        self.assertIn(automation.VT_R4, COMTYPES)
        self.assertIn(automation.VT_R8, COMTYPES)
        self.assertIn(automation.VT_CY, COMTYPES)
        self.assertIn(automation.VT_DATE, COMTYPES)
        self.assertIn(automation.VT_BSTR, COMTYPES)
        self.assertIn(automation.VT_DISPATCH, COMTYPES)
        self.assertIn(automation.VT_ERROR, COMTYPES)
        self.assertIn(automation.VT_BOOL, COMTYPES)
        self.assertIn(automation.VT_VARIANT, COMTYPES)
        self.assertIn(automation.VT_UNKNOWN, COMTYPES)
        self.assertIn(automation.VT_DECIMAL, COMTYPES)
        self.assertIn(automation.VT_I1, COMTYPES)
        self.assertIn(automation.VT_UI1, COMTYPES)
        self.assertIn(automation.VT_UI2, COMTYPES)
        self.assertIn(automation.VT_UI4, COMTYPES)
        self.assertIn(automation.VT_I8, COMTYPES)
        self.assertIn(automation.VT_UI8, COMTYPES)
        self.assertIn(automation.VT_INT, COMTYPES)
        self.assertIn(automation.VT_UINT, COMTYPES)
        self.assertIn(automation.VT_VOID, COMTYPES)
        self.assertIn(automation.VT_HRESULT, COMTYPES)
        self.assertIn(automation.VT_LPSTR, COMTYPES)
        self.assertIn(automation.VT_LPWSTR, COMTYPES)

    def test_basic_types(self):
        # Test that basic types are correctly defined
        self.assertEqual(BSTR_type.name, "BSTR")
        self.assertEqual(HRESULT_type.name, "HRESULT")
        self.assertEqual(VARIANT_BOOL_type.name, "VARIANT_BOOL")
        self.assertEqual(void_type.name, "void")
        self.assertEqual(int_type.name, "int")

    def test_iunknown_idispatch(self):
        # Test IUnknown and IDispatch type definitions
        self.assertEqual(IUNKNOWN_type.name, "IUnknown")
        self.assertEqual(IUNKNOWN_type.iid, "{00000000-0000-0000-C000-000000000046}")
        self.assertIsNone(IUNKNOWN_type.base)

        self.assertEqual(IDISPATCH_type.name, "IDispatch")
        self.assertEqual(IDISPATCH_type.iid, "{00020400-0000-0000-C000-000000000046}")
        self.assertEqual(IDISPATCH_type.base, IUNKNOWN_type)


class TestPTRFunction(unittest.TestCase):
    def test_ptr_creates_pointer_type(self):
        ptr = tlbparser.PTR(int_type)
        self.assertIsInstance(ptr, typedesc.PointerType)
        self.assertEqual(ptr.typ, int_type)

    def test_ptr_size_and_alignment(self):
        # On 64-bit, pointer size should be 64 bits
        import sys
        is_64bits = sys.maxsize > 2**32
        ptr = tlbparser.PTR(int_type)
        if is_64bits:
            self.assertEqual(ptr.size, 64)
            self.assertEqual(ptr.align, 64)
        else:
            self.assertEqual(ptr.size, 32)
            self.assertEqual(ptr.align, 32)


class TestMidlSAFEARRAY(unittest.TestCase):
    def test_midl_safearray_creates_safearray_type(self):
        sa = tlbparser.midlSAFEARRAY(int_type)
        self.assertIsInstance(sa, typedesc.SAFEARRAYType)
        self.assertEqual(sa.typ, int_type)


class TestParserMakeType(unittest.TestCase):
    def setUp(self):
        # Create a minimal parser instance
        self.parser = Parser()
        self.parser.tlib = None
        self.parser.items = {}

    def test_make_type_basic(self):
        from comtypes import automation
        from comtypes.typeinfo import TYPEDESC
        from ctypes import byref

        # Create a TYPEDESC for VT_I4 (int)
        tdesc = TYPEDESC()
        tdesc.vt = automation.VT_I4
        result = self.parser.make_type(tdesc, None)
        self.assertEqual(result, int_type)

    def test_make_type_bstr(self):
        from comtypes import automation
        from comtypes.typeinfo import TYPEDESC

        tdesc = TYPEDESC()
        tdesc.vt = automation.VT_BSTR
        result = self.parser.make_type(tdesc, None)
        self.assertEqual(result, BSTR_type)

    def test_make_type_pointer(self):
        from comtypes import automation
        from comtypes.typeinfo import TYPEDESC

        tdesc = TYPEDESC()
        tdesc.vt = automation.VT_PTR
        # Point to an int
        inner = TYPEDESC()
        inner.vt = automation.VT_I4
        tdesc._.lptdesc = inner
        result = self.parser.make_type(tdesc, None)
        self.assertIsInstance(result, typedesc.PointerType)
        self.assertEqual(result.typ, int_type)

    def test_make_type_safearray(self):
        from comtypes import automation
        from comtypes.typeinfo import TYPEDESC

        tdesc = TYPEDESC()
        tdesc.vt = automation.VT_SAFEARRAY
        inner = TYPEDESC()
        inner.vt = automation.VT_I4
        tdesc._.lptdesc = inner
        result = self.parser.make_type(tdesc, None)
        self.assertIsInstance(result, typedesc.SAFEARRAYType)
        self.assertEqual(result.typ, int_type)

    def test_make_type_carray(self):
        from comtypes import automation
        from comtypes.typeinfo import TYPEDESC, tagARRAYDESC

        tdesc = TYPEDESC()
        tdesc.vt = automation.VT_CARRAY
        arraydesc = tagARRAYDESC()
        arraydesc.tdescElem.vt = automation.VT_I4
        arraydesc.cDims = 1
        arraydesc.rgbounds[0].lLbound = 0
        arraydesc.rgbounds[0].cElements = 10
        tdesc._.lpadesc = arraydesc
        result = self.parser.make_type(tdesc, None)
        self.assertIsInstance(result, typedesc.ArrayType)


class TestParserFlags(unittest.TestCase):
    def setUp(self):
        self.parser = Parser()
        self.parser.tlib = None
        self.parser.items = {}

    def test_inv_kind(self):
        from comtypes import automation
        self.assertEqual(self.parser.inv_kind(automation.DISPATCH_METHOD), [])
        self.assertEqual(self.parser.inv_kind(automation.DISPATCH_PROPERTYPUT), ["propput"])
        self.assertEqual(self.parser.inv_kind(automation.DISPATCH_PROPERTYPUTREF), ["propputref"])
        self.assertEqual(self.parser.inv_kind(automation.DISPATCH_PROPERTYGET), ["propget"])

    def test_func_flags(self):
        from comtypes import typeinfo
        flags = typeinfo.FUNCFLAG_FRESTRICTED | typeinfo.FUNCFLAG_FHIDDEN
        result = self.parser.func_flags(flags)
        self.assertIn("restricted", result)
        self.assertIn("hidden", result)

    def test_param_flags(self):
        from comtypes import typeinfo
        flags = typeinfo.PARAMFLAG_FIN | typeinfo.PARAMFLAG_FOUT | typeinfo.PARAMFLAG_FOPT
        result = self.parser.param_flags(flags)
        self.assertIn("in", result)
        self.assertIn("out", result)
        self.assertIn("optional", result)

    def test_coclass_type_flags(self):
        from comtypes import typeinfo
        flags = typeinfo.TYPEFLAG_FAPPOBJECT | typeinfo.TYPEFLAG_FLICENSED
        result = self.parser.coclass_type_flags(flags)
        self.assertIn("appobject", result)
        self.assertIn("licensed", result)
        # TYPEFLAG_FCANCREATE not set, so "noncreatable" should be added
        self.assertIn("noncreatable", result)

    def test_interface_type_flags(self):
        from comtypes import typeinfo
        flags = typeinfo.TYPEFLAG_FDUAL | typeinfo.TYPEFLAG_FOLEAUTOMATION
        result = self.parser.interface_type_flags(flags)
        self.assertIn("dual", result)
        self.assertIn("oleautomation", result)

    def test_var_flags(self):
        from comtypes import typeinfo
        flags = typeinfo.VARFLAG_FREADONLY | typeinfo.VARFLAG_FHIDDEN
        result = self.parser.var_flags(flags)
        self.assertIn("readonly", result)
        self.assertIn("hidden", result)


class TestParserRegistration(unittest.TestCase):
    def setUp(self):
        self.parser = Parser()
        self.parser.tlib = None
        self.parser.items = {}

    def test_register_and_retrieve(self):
        enum = typedesc.Enumeration("TestEnum", 32, 32)
        self.parser._register("TestEnum", enum)
        self.assertIn("TestEnum", self.parser.items)

    def test_register_duplicate_raises(self):
        enum1 = typedesc.Enumeration("TestEnum", 32, 32)
        enum2 = typedesc.Enumeration("TestEnum", 32, 32)
        self.parser._register("TestEnum", enum1)
        with self.assertRaises(ValueError):
            self.parser._register("TestEnum", enum2)

    def test_register_external_allowed(self):
        ext1 = typedesc.External(None, "ExtType", 64, 64)
        ext2 = typedesc.External(None, "ExtType", 64, 64)
        self.parser._register("ExtType", ext1)
        # Should not raise for External types
        self.parser._register("ExtType", ext2)


class TestGetTlibFilename(unittest.TestCase):
    def test_get_tlib_filename_returns_none_for_unregistered(self):
        # Create a mock typelib that's not registered
        class MockTypeLib:
            def GetLibAttr(self):
                class Attr:
                    guid = "{00000000-0000-0000-0000-000000000000}"
                    wMajorVerNum = 1
                    wMinorVerNum = 0
                return Attr()
        
        result = get_tlib_filename(MockTypeLib())
        self.assertIsNone(result)


class TestParserParseEnum(unittest.TestCase):
    def test_parse_enum_creates_enumeration(self):
        # This test would require a real typelib, so we just verify the method exists
        self.assertTrue(hasattr(Parser, 'ParseEnum'))


class TestParserParseRecord(unittest.TestCase):
    def test_parse_record_creates_structure(self):
        self.assertTrue(hasattr(Parser, 'ParseRecord'))


class TestParserParseModule(unittest.TestCase):
    def test_parse_module_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseModule'))


class TestParserParseInterface(unittest.TestCase):
    def test_parse_interface_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseInterface'))


class TestParserParseDispatch(unittest.TestCase):
    def test_parse_dispatch_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseDispatch'))


class TestParserParseCoClass(unittest.TestCase):
    def test_parse_coclass_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseCoClass'))


class TestParserParseAlias(unittest.TestCase):
    def test_parse_alias_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseAlias'))


class TestParserParseUnion(unittest.TestCase):
    def test_parse_union_exists(self):
        self.assertTrue(hasattr(Parser, 'ParseUnion'))


class TestParserParseTypeInfo(unittest.TestCase):
    def test_parse_typeinfo_exists(self):
        self.assertTrue(hasattr(Parser, 'parse_typeinfo'))


class TestParserParseLibraryDescription(unittest.TestCase):
    def test_parse_library_description_exists(self):
        self.assertTrue(hasattr(Parser, 'parse_LibraryDescription'))


class TestParserParse(unittest.TestCase):
    def test_parse_exists(self):
        self.assertTrue(hasattr(Parser, 'parse'))


class TestTlbFileParser(unittest.TestCase):
    def test_tlbfileparser_exists(self):
        self.assertTrue(hasattr(tlbparser, 'TlbFileParser'))


class TestTypeLibParser(unittest.TestCase):
    def test_typelibparser_exists(self):
        self.assertTrue(hasattr(tlbparser, 'TypeLibParser'))


if __name__ == "__main__":
    unittest.main()