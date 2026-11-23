"""Tests for ``comtypes.tools.codegenerator.codegenerator``.

These tests exercise the CodeGenerator class to improve coverage.
"""

import unittest
from comtypes.tools import typedesc
from comtypes.tools.codegenerator import codegenerator
from comtypes.tools.tlbparser import (
    BSTR_type,
    HRESULT_type,
    VARIANT_BOOL_type,
    VARIANT_type,
    void_type,
    int_type,
    IUNKNOWN_type,
    IDISPATCH_type,
)


class TestCodeGenerator(unittest.TestCase):
    def setUp(self):
        self.gen = codegenerator.CodeGenerator()

    def test_generate_fundamental_type(self):
        # FundamentalType should be a no-op
        self.gen.generate(int_type)
        self.assertIn(int_type, self.gen.done)

    def test_generate_typedef(self):
        my_int = typedesc.Typedef("MyInt", int_type)
        self.gen.generate(my_int)
        self.assertIn(my_int, self.gen.done)
        self.assertIn("MyInt", self.gen.names)

    def test_generate_enum(self):
        enum = typedesc.Enumeration("TestEnum", 32, 32)
        val1 = typedesc.EnumValue("VAL1", 1, enum)
        val2 = typedesc.EnumValue("VAL2", 2, enum)
        enum.add_value(val1)
        enum.add_value(val2)
        self.gen.generate(enum)
        self.assertIn(enum, self.gen.done)
        self.assertIn("TestEnum", self.gen.names)
        self.assertIn("VAL1", self.gen.names)
        self.assertIn("VAL2", self.gen.names)

    def test_generate_structure(self):
        struct = typedesc.Structure("TestStruct", align=32, size=64)
        field = typedesc.Field("field1", int_type, None, 0)
        struct.add_member(field)
        self.gen.generate(struct)
        self.assertIn(struct, self.gen.done)
        self.assertIn("TestStruct", self.gen.names)

    def test_generate_union(self):
        union = typedesc.Union("TestUnion", align=32, size=64)
        field = typedesc.Field("field1", int_type, None, 0)
        union.add_member(field)
        self.gen.generate(union)
        self.assertIn(union, self.gen.done)
        self.assertIn("TestUnion", self.gen.names)

    def test_generate_pointer_type(self):
        ptr = typedesc.PointerType(int_type, 64, 64)
        self.gen.generate(ptr)
        self.assertIn(ptr, self.gen.done)

    def test_generate_array_type(self):
        arr = typedesc.ArrayType(int_type, 0, 9)
        self.gen.generate(arr)
        self.assertIn(arr, self.gen.done)

    def test_generate_safearray_type(self):
        sa = typedesc.SAFEARRAYType(int_type)
        self.gen.generate(sa)
        self.assertIn(sa, self.gen.done)

    def test_generate_constant(self):
        const = typedesc.Constant("MY_CONST", int_type, 42, "doc")
        self.gen.generate(const)
        self.assertIn(const, self.gen.done)
        self.assertIn("MY_CONST", self.gen.names)

    def test_generate_com_interface(self):
        itf = typedesc.ComInterface(
            "ITest", IUNKNOWN_type, "{00000000-0000-0000-0000-000000000001}", [], None
        )
        mth = typedesc.ComMethod(1, 1610678272, "TestMethod", HRESULT_type, [], None)
        itf.extend_members([mth])
        self.gen.generate(itf)
        self.assertIn(itf, self.gen.done)
        self.assertIn("ITest", self.gen.names)

    def test_generate_disp_interface(self):
        itf = typedesc.DispInterface(
            "IDispTest", IDISPATCH_type, "{00000000-0000-0000-0000-000000000002}", [], None
        )
        prop = typedesc.DispProperty(1, "TestProp", BSTR_type, ["readonly"], None)
        itf.add_member(prop)
        self.gen.generate(itf)
        self.assertIn(itf, self.gen.done)
        self.assertIn("IDispTest", self.gen.names)

    def test_generate_coclass(self):
        from comtypes.typeinfo import TLIBATTR
        tlibattr = TLIBATTR()
        tlibattr.guid = "{00000000-0000-0000-0000-000000000003}"
        tlibattr.wMajorVerNum = 1
        tlibattr.wMinorVerNum = 0
        tlibattr.lcid = 0
        tlibattr.wLibFlags = 0
        tlibattr.syskind = 1  # SYS_WIN32

        coclass = typedesc.CoClass(
            "TestClass", "{00000000-0000-0000-0000-000000000004}", [], tlibattr, None
        )
        itf = typedesc.ComInterface(
            "ITest", IUNKNOWN_type, "{00000000-0000-0000-0000-000000000005}", [], None
        )
        coclass.add_interface(itf, 0)
        self.gen.generate(coclass)
        self.assertIn(coclass, self.gen.done)
        self.assertIn("TestClass", self.gen.names)

    def test_generate_typelib(self):
        tlib = typedesc.TypeLib("TestLib", "{00000000-0000-0000-0000-000000000006}", 1, 0, None)
        self.gen.generate(tlib)
        self.assertIn(tlib, self.gen.done)
        self.assertIn("Library", self.gen.names)

    def test_generate_external(self):
        from comtypes.typeinfo import ITypeLib
        from ctypes import POINTER
        # Create a mock typelib
        class MockTypeLib:
            def GetLibAttr(self):
                class Attr:
                    guid = "{00000000-0000-0000-0000-000000000007}"
                    wMajorVerNum = 1
                    wMinorVerNum = 0
                    lcid = 0
                return Attr()
            def GetDocumentation(self, index):
                return ("MockLib", None, None, None)
        
        ext = typedesc.External(MockTypeLib(), "ExternalType", 64, 64)
        self.gen.generate(ext)
        self.assertIn(ext, self.gen.done)

    def test_generate_wrapper_code(self):
        # Test generating wrapper code for a simple typelib
        tlib = typedesc.TypeLib("TestLib", "{00000000-0000-0000-0000-000000000008}", 1, 0, None)
        code = self.gen.generate_wrapper_code([tlib], None)
        self.assertIn("TestLib", code)
        self.assertIn("__all__", code)

    def test_generate_friendly_code(self):
        # Test generating friendly code
        tlib = typedesc.TypeLib("TestLib", "{00000000-0000-0000-0000-000000000009}", 1, 0, None)
        self.gen.generate(tlib)
        code = self.gen.generate_friendly_code("comtypes.gen._test")
        self.assertIn("import comtypes.gen._test as __wrapper_module__", code)
        self.assertIn("__all__", code)

    def test_known_symbols(self):
        # Test that known symbols are imported instead of generated
        gen = codegenerator.CodeGenerator(known_symbols={"IUnknown": "comtypes"})
        iunknown = typedesc.ComInterface(
            "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", ["hidden"], None
        )
        gen.generate(iunknown)
        # Should have imported IUnknown from comtypes
        self.assertIn("IUnknown", gen.imports.get_symbols())

    def test_known_interfaces(self):
        # Test that known interfaces are imported instead of generated
        gen = codegenerator.CodeGenerator(known_interfaces={"IUnknown": "{00000000-0000-0000-C000-000000000046}"})
        iunknown = typedesc.ComInterface(
            "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", ["hidden"], None
        )
        gen.generate(iunknown)
        # Should have imported IUnknown from comtypes
        self.assertIn("IUnknown", gen.imports.get_symbols())

    def test_adjust_blank(self):
        # Test the adjust_blank context manager
        with self.gen.adjust_blank("class") as f:
            print("class Test:", file=f)
        self.assertEqual(self.gen.last_item, "class")

        with self.gen.adjust_blank("attribute") as f:
            print("attr = 1", file=f)
        self.assertEqual(self.gen.last_item, "attribute")

    def test_make_dunder_all_part(self):
        self.gen.names = {"A", "B", "C"}
        result = self.gen._make_dunder_all_part()
        self.assertIn("__all__", result)
        self.assertIn("'A'", result)
        self.assertIn("'B'", result)
        self.assertIn("'C'", result)

    def test_make_friendly_module_import_part(self):
        self.gen.names = {"A", "B"}
        self.gen.imports.add("comtypes", "GUID")
        result = self.gen._make_friendly_module_import_part("comtypes.gen._test")
        self.assertIn("from comtypes.gen._test import", result)
        self.assertIn("A", result)
        self.assertIn("B", result)

    def test_need_variant_imports(self):
        from decimal import Decimal
        import datetime
        self.gen.need_VARIANT_imports(Decimal("1.5"))
        self.assertIn("Decimal", self.gen.imports.get_symbols())
        self.gen.need_VARIANT_imports(datetime.datetime(2024, 1, 1))
        self.assertIn("datetime", self.gen.imports.get_symbols())


if __name__ == "__main__":
    unittest.main()