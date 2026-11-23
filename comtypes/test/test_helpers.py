"""Tests for the code-generation helpers in ``comtypes.tools.codegenerator.helpers``.

These helpers turn parsed type descriptions into the literal Python source
that is written into the generated modules. A wrong literal here means the
generated module does not import or does not describe the COM member
correctly, so the exact output is worth pinning down.
"""

import unittest

from comtypes.tools import typedesc
from comtypes.tools.codegenerator import helpers


class Test_lcid(unittest.TestCase):
    def test_repr_is_the_lcid_marker(self):
        # ``lcid`` is a sentinel used as the default value of an LCID
        # parameter. Its repr is what ends up in the generated source.
        self.assertEqual(repr(helpers.lcid), "_lcid")


class Test_to_arg_definition(unittest.TestCase):
    def test_single_line_without_default(self):
        # A short argument definition is emitted on one line.
        result = helpers._to_arg_definition("c_int", "a", ["in"], None)
        self.assertEqual(result, "        (['in'], c_int, 'a')")

    def test_single_line_with_default(self):
        # A short argument definition with a default keeps the default.
        result = helpers._to_arg_definition("c_int", "a", ["in"], 0)
        self.assertEqual(result, "        (['in'], c_int, 'a', 0)")

    def test_wraps_long_definition_without_default(self):
        # A definition longer than 80 characters is wrapped over several
        # lines so the generated source stays readable.
        result = helpers._to_arg_definition(
            "POINTER(POINTER(SomeVeryLongTypeName))",
            "a_very_long_argument_name",
            ["in"],
            None,
        )
        expected = (
            "        (\n"
            "            ['in'],\n"
            "            POINTER(POINTER(SomeVeryLongTypeName)),\n"
            "            'a_very_long_argument_name',\n"
            "        )"
        )
        self.assertEqual(result, expected)

    def test_wraps_long_definition_with_default(self):
        # The wrapped form with a default has no trailing comma after the
        # argument name; the default follows on its own line.
        result = helpers._to_arg_definition(
            "POINTER(POINTER(SomeVeryLongTypeName))",
            "a_very_long_argument_name",
            ["in"],
            0,
        )
        expected = (
            "        (\n"
            "            ['in'],\n"
            "            POINTER(POINTER(SomeVeryLongTypeName)),\n"
            "            'a_very_long_argument_name',\n"
            "            0\n"
            "        )"
        )
        self.assertEqual(result, expected)


class Test_ComMethodGenerator(unittest.TestCase):
    def test_wraps_long_noargs_method(self):
        # A COMMETHOD line longer than 80 characters is wrapped.
        method = typedesc.ComMethod(
            1,
            0,
            "a_very_long_method_name_here",
            typedesc.Typedef(
                "SomeVeryLongTypeNameForReturn",
                typedesc.FundamentalType("int", 32, 32),
            ),
            [],
            None,
        )
        result = helpers.ComMethodGenerator(method, isdual=False).generate()
        expected = (
            "    COMMETHOD(\n"
            "        [],\n"
            "        SomeVeryLongTypeNameForReturn,\n"
            "        'a_very_long_method_name_here',\n"
            "    ),"
        )
        self.assertEqual(result, expected)

    def test_openarray_argument_is_marked_in(self):
        # An open array is declared in the typelib as a COM interface
        # parameter. It is emitted as OPENARRAY and forced to be an [in]
        # parameter, because comtypes cannot pass an interface as an array.
        itf = typedesc.ComInterface(
            "IFoo", None, "{00000000-0000-0000-0000-000000000000}", [], None
        )
        method = typedesc.ComMethod(
            1,
            0,
            "Do",
            typedesc.Typedef("HRESULT", typedesc.FundamentalType("int", 32, 32)),
            [],
            None,
        )
        method.add_argument(itf, "data", ["out"], None)
        result = helpers.ComMethodGenerator(method, isdual=False).generate()
        self.assertIn("OPENARRAY", result)
        self.assertIn("['out', 'in']", result)

    def test_lcid_argument_uses_lcid_default(self):
        # An argument flagged ``lcid`` gets the ``lcid`` sentinel as its
        # default, so the generated wrapper fills in the current locale.
        method = typedesc.ComMethod(
            1,
            0,
            "Do",
            typedesc.Typedef("HRESULT", typedesc.FundamentalType("int", 32, 32)),
            [],
            None,
        )
        method.add_argument(
            typedesc.FundamentalType("int", 32, 32), "lcid", ["in", "lcid"], None
        )
        result = helpers.ComMethodGenerator(method, isdual=False).generate()
        self.assertIn("_lcid", result)


class Test_DispMethodGenerator(unittest.TestCase):
    def test_wraps_long_noargs_method(self):
        # A DISPMETHOD line longer than 80 characters is wrapped.
        method = typedesc.DispMethod(
            1,
            1,
            "a_very_long_method_name_here",
            typedesc.Typedef(
                "SomeVeryLongTypeNameForReturn",
                typedesc.FundamentalType("int", 32, 32),
            ),
            [],
            None,
        )
        result = helpers.DispMethodGenerator(method).generate()
        expected = (
            "    DISPMETHOD(\n"
            "        [dispid(1)],\n"
            "        SomeVeryLongTypeNameForReturn,\n"
            "        'a_very_long_method_name_here',\n"
            "    ),"
        )
        self.assertEqual(result, expected)


class Test_DispPropertyGenerator(unittest.TestCase):
    def test_wraps_long_property(self):
        # A DISPPROPERTY line longer than 80 characters is wrapped. Unlike
        # the method forms there is no trailing comma after the name.
        prop = typedesc.DispProperty(
            1,
            "a_very_long_property_name_here",
            typedesc.Typedef(
                "SomeVeryLongTypeNameForReturn",
                typedesc.FundamentalType("int", 32, 32),
            ),
            [],
            None,
        )
        result = helpers.DispPropertyGenerator(prop).generate()
        expected = (
            "    DISPPROPERTY(\n"
            "        [dispid(1)],\n"
            "        SomeVeryLongTypeNameForReturn,\n"
            "        'a_very_long_property_name_here'\n"
            "    ),"
        )
        self.assertEqual(result, expected)


class Test_TypeNamer(unittest.TestCase):
    def test_safearray_type(self):
        # A SAFEARRAY is emitted as ``_midlSAFEARRAY(element)``.
        namer = helpers.TypeNamer()
        result = namer(typedesc.SAFEARRAYType(typedesc.FundamentalType("int", 32, 32)))
        self.assertEqual(result, "_midlSAFEARRAY(c_int)")

    def test_array_type(self):
        # An array is emitted as ``element * count`` where count is
        # ``max + 1``.
        namer = helpers.TypeNamer()
        result = namer(
            typedesc.ArrayType(typedesc.FundamentalType("int", 32, 32), 0, 3)
        )
        self.assertEqual(result, "c_int * 4")

    def test_unnamed_enumeration_is_c_int(self):
        # An anonymous enum has no name to refer to, so it is emitted as
        # ``c_int``.
        namer = helpers.TypeNamer()
        result = namer(typedesc.Enumeration("", 32, 32))
        self.assertEqual(result, "c_int")

    def test_enum_value_that_is_a_keyword_gets_a_suffix(self):
        # An enum member whose name is a Python keyword must be renamed so
        # the generated module still imports.
        namer = helpers.TypeNamer()
        enum = typedesc.Enumeration("E", 32, 32)
        value = typedesc.EnumValue("class", 1, enum)
        self.assertEqual(namer(value), "class_")

    def test_pointer_to_char_is_string(self):
        # A ``char *`` is emitted as the ``STRING`` typedef, not as a raw
        # pointer, so comtypes marshals it as a string.
        namer = helpers.TypeNamer()
        result = namer(
            typedesc.PointerType(typedesc.FundamentalType("char", 8, 8), 64, 64)
        )
        self.assertEqual(result, "STRING")

    def test_pointer_to_wchar_t_is_wstring(self):
        # A ``wchar_t *`` is emitted as the ``WSTRING`` typedef.
        namer = helpers.TypeNamer()
        result = namer(
            typedesc.PointerType(typedesc.FundamentalType("wchar_t", 16, 16), 64, 64)
        )
        self.assertEqual(result, "WSTRING")


if __name__ == "__main__":
    unittest.main()
