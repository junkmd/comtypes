"""Tests for ``comtypes.tools.codegenerator.namespaces``.

These tests exercise the namespace management classes to improve coverage.
"""

import unittest
import warnings
from comtypes.tools.codegenerator import namespaces


class TestImportedNamespaces(unittest.TestCase):
    def test_add_module(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("datetime")
        self.assertIn("datetime", imports)
        self.assertEqual(imports.getvalue(), "import datetime")

    def test_add_wildcard_import(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("ctypes", "*")
        self.assertIn(("ctypes", "*"), imports)
        self.assertEqual(imports.getvalue(), "from ctypes import *")

    def test_add_specific_symbol(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("decimal", "Decimal")
        self.assertIn(("decimal", "Decimal"), imports)
        self.assertEqual(imports.getvalue(), "from decimal import Decimal")

    def test_add_multiple_symbols_same_module(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("comtypes", "GUID")
        imports.add("comtypes", "COMMETHOD")
        imports.add("comtypes", "IUnknown")
        self.assertIn("GUID", imports.get_symbols())
        self.assertIn("COMMETHOD", imports.get_symbols())
        self.assertIn("IUnknown", imports.get_symbols())
        result = imports.getvalue()
        self.assertIn("from comtypes import", result)
        self.assertIn("GUID", result)
        self.assertIn("COMMETHOD", result)
        self.assertIn("IUnknown", result)

    def test_add_via_symbols_mapping(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("GUID", symbols={"GUID": "comtypes"})
        self.assertIn("GUID", imports.get_symbols())
        self.assertEqual(imports.getvalue(), "from comtypes import GUID")

    def test_contains_module(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("datetime")
        self.assertTrue("datetime" in imports)
        self.assertFalse("os" in imports)

    def test_contains_tuple(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("ctypes", "*")
        self.assertTrue(("ctypes", "*") in imports)
        self.assertFalse(("ctypes", "c_int") in imports)

    def test_get_symbols(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("datetime")
        imports.add("decimal", "Decimal")
        imports.add("ctypes", "*")
        symbols = imports.get_symbols()
        self.assertEqual(symbols, {"Decimal"})

    def test_getvalue_multiple_modules(self):
        imports = namespaces.ImportedNamespaces()
        imports.add("datetime")
        imports.add("decimal", "Decimal")
        imports.add("comtypes", "GUID")
        imports.add("comtypes", "COMMETHOD")
        result = imports.getvalue()
        lines = result.split("\n")
        self.assertEqual(len(lines), 3)
        self.assertTrue(lines[0].startswith("import datetime"))
        self.assertTrue(lines[1].startswith("from decimal import"))
        self.assertTrue(lines[2].startswith("from comtypes import"))

    def test_long_line_wrapping(self):
        imports = namespaces.ImportedNamespaces()
        # Add many symbols to trigger line wrapping
        for i in range(20):
            imports.add("comtypes", f"SYMBOL{i}")
        result = imports.getvalue()
        self.assertIn("from comtypes import (", result)
        self.assertIn(")", result)


class TestDeclaredNamespaces(unittest.TestCase):
    def test_add_simple(self):
        decl = namespaces.DeclaredNamespaces()
        decl.add("STRING", "c_char_p")
        self.assertIn("STRING", decl.get_symbols())
        self.assertEqual(decl.getvalue(), "STRING = c_char_p")

    def test_add_with_comment(self):
        decl = namespaces.DeclaredNamespaces()
        decl.add("_lcid", "0", "change this if required")
        self.assertIn("_lcid", decl.get_symbols())
        self.assertEqual(decl.getvalue(), "_lcid = 0  # change this if required")

    def test_add_multiple(self):
        decl = namespaces.DeclaredNamespaces()
        decl.add("STRING", "c_char_p")
        decl.add("WSTRING", "c_wchar_p")
        decl.add("_lcid", "0", "comment")
        result = decl.getvalue()
        lines = result.split("\n")
        self.assertEqual(len(lines), 3)
        self.assertIn("STRING = c_char_p", lines[0])
        self.assertIn("WSTRING = c_wchar_p", lines[1])
        self.assertIn("_lcid = 0  # comment", lines[2])

    def test_get_symbols(self):
        decl = namespaces.DeclaredNamespaces()
        decl.add("A", "1")
        decl.add("B", "2")
        self.assertEqual(decl.get_symbols(), {"A", "B"})


class TestEnumerationNamespaces(unittest.TestCase):
    def test_add_single_enum(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Foo", "ham", 1)
        enums.add("Foo", "spam", 2)
        self.assertTrue("Foo" in enums)
        self.assertFalse("Bar" in enums)
        self.assertTrue(bool(enums))

    def test_add_multiple_enums(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Foo", "ham", 1)
        enums.add("Bar", "bacon", 3)
        self.assertTrue("Foo" in enums)
        self.assertTrue("Bar" in enums)

    def test_duplicate_member_warning(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Bar", "egg", 4)
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            enums.add("Bar", "egg", 5)
            self.assertEqual(len(w), 1)
            self.assertIn("duplicated", str(w[0].message))

    def test_get_symbols(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Foo", "ham", 1)
        enums.add("Bar", "bacon", 3)
        self.assertEqual(enums.get_symbols(), {"Foo", "Bar"})

    def test_to_constants(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Foo", "ham", 1)
        enums.add("Foo", "spam", 2)
        enums.add("Bar", "bacon", 3)
        enums.add("Baz", "mix", -1)
        enums.add("Baz", "juice", 0)
        result = enums.to_constants()
        self.assertIn("# values for enumeration 'Foo'", result)
        self.assertIn("ham = 1", result)
        self.assertIn("spam = 2", result)
        self.assertIn("Foo = c_int  # enum", result)
        self.assertIn("# values for enumeration 'Bar'", result)
        self.assertIn("bacon = 3", result)
        self.assertIn("Bar = c_int  # enum", result)
        self.assertIn("# values for enumeration 'Baz'", result)
        self.assertIn("mix = -1", result)
        self.assertIn("juice = 0", result)
        self.assertIn("Baz = c_int  # enum", result)

    def test_to_enums(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Foo", "ham", 1)
        enums.add("Foo", "spam", 2)
        enums.add("Bar", "bacon", 3)
        enums.add("Baz", "mix", -1)
        enums.add("Baz", "juice", 0)
        enumcode, enumbases = enums.to_enums()
        self.assertIn("class Foo(IntFlag):", enumcode)
        self.assertIn("    ham = 1", enumcode)
        self.assertIn("    spam = 2", enumcode)
        self.assertIn("class Bar(IntFlag):", enumcode)
        self.assertIn("    bacon = 3", enumcode)
        self.assertIn("class Baz(IntEnum):", enumcode)
        self.assertIn("    mix = -1", enumcode)
        self.assertIn("    juice = 0", enumcode)
        self.assertIn("IntFlag", enumbases)
        self.assertIn("IntEnum", enumbases)

    def test_duplicate_in_to_enums(self):
        enums = namespaces.EnumerationNamespaces()
        enums.add("Bar", "egg", 4)
        enums.add("Bar", "egg", 5)
        enumcode, _ = enums.to_enums()
        # Should have commented out the first duplicate
        self.assertIn("# egg = 4  # duplicated", enumcode)
        self.assertIn("egg = 5  # duplicated", enumcode)

    def test_contains_and_bool(self):
        enums = namespaces.EnumerationNamespaces()
        self.assertFalse(bool(enums))
        self.assertFalse("Foo" in enums)
        enums.add("Foo", "ham", 1)
        self.assertTrue(bool(enums))
        self.assertTrue("Foo" in enums)


if __name__ == "__main__":
    unittest.main()