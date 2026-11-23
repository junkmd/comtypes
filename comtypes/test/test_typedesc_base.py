"""Tests for the plain data classes in ``comtypes.tools.typedesc_base``.

These classes are the in-memory representation of a parsed C header. The
code generator reads their attributes to emit Python wrappers, so a wrong
attribute name or a missing ``int()`` conversion would silently produce
broken generated code. The tests below pin down the constructor contract of
each class and the small amount of behaviour in ``_HasArgs``.
"""

import unittest

from comtypes.tools import typedesc_base


class TestArgument(unittest.TestCase):
    def test_stores_atype_and_name(self):
        # ``Argument`` is the building block of every callable's argument
        # list. If it stopped storing the type and name it was given, the
        # generated wrappers would lose the parameter information.
        arg = typedesc_base.Argument("int", "value")
        self.assertEqual(arg.atype, "int")
        self.assertEqual(arg.name, "value")


class TestHasArgs(unittest.TestCase):
    def test_starts_with_no_arguments(self):
        # A freshly created callable has an empty argument list. Code that
        # iterates the arguments must not see a stale or shared list.
        has_args = typedesc_base._HasArgs()
        self.assertEqual(has_args.arguments, [])

    def test_add_argument_appends_in_order(self):
        # ``add_argument`` is how the parser records each parameter. The
        # arguments must be appended in order so the generated signature
        # keeps the original parameter order.
        has_args = typedesc_base._HasArgs()
        first = typedesc_base.Argument("int", "a")
        second = typedesc_base.Argument("double", "b")
        has_args.add_argument(first)
        has_args.add_argument(second)
        self.assertEqual(has_args.arguments, [first, second])

    def test_iter_arg_types_yields_atypes(self):
        # ``iterArgTypes`` is used to build the ctypes ``argtypes`` tuple. It
        # must yield the type of each argument in order.
        has_args = typedesc_base._HasArgs()
        has_args.add_argument(typedesc_base.Argument("int", "a"))
        has_args.add_argument(typedesc_base.Argument("double", "b"))
        self.assertEqual(list(has_args.iterArgTypes()), ["int", "double"])

    def test_iter_arg_names_yields_names(self):
        # ``iterArgNames`` is used to build the Python parameter names. It
        # must yield the name of each argument in order.
        has_args = typedesc_base._HasArgs()
        has_args.add_argument(typedesc_base.Argument("int", "a"))
        has_args.add_argument(typedesc_base.Argument("double", "b"))
        self.assertEqual(list(has_args.iterArgNames()), ["a", "b"])

    def test_fixup_argtypes_replaces_types_in_place(self):
        # ``fixup_argtypes`` rewrites argument types through a mapping, e.g.
        # when a typedef is resolved to its underlying type. The argument
        # objects must be updated in place so later readers see the new type.
        has_args = typedesc_base._HasArgs()
        arg = typedesc_base.Argument("Alias", "a")
        has_args.add_argument(arg)
        has_args.fixup_argtypes({"Alias": "int"})
        self.assertEqual(arg.atype, "int")


class TestAlias(unittest.TestCase):
    def test_stores_name_and_alias_with_no_type(self):
        # ``Alias`` represents a C ``#define A B``. The optional ``typ`` is
        # None when the alias has no known type.
        alias = typedesc_base.Alias("A", "B")
        self.assertEqual(alias.name, "A")
        self.assertEqual(alias.alias, "B")
        self.assertIsNone(alias.typ)

    def test_stores_explicit_type(self):
        # When the parser knows the type of the alias it must be kept.
        alias = typedesc_base.Alias("A", "B", "int")
        self.assertEqual(alias.typ, "int")


class TestMacro(unittest.TestCase):
    def test_stores_name_args_and_body(self):
        # ``Macro`` represents a C preprocessor macro with arguments. All
        # three parts are needed to reproduce the definition.
        macro = typedesc_base.Macro("CD_INDRIVE", "(status)", "((int)status > 0)")
        self.assertEqual(macro.name, "CD_INDRIVE")
        self.assertEqual(macro.args, "(status)")
        self.assertEqual(macro.body, "((int)status > 0)")


class TestFile(unittest.TestCase):
    def test_stores_name(self):
        # ``File`` records the source file a declaration came from.
        file = typedesc_base.File("foo.h")
        self.assertEqual(file.name, "foo.h")


class TestFunction(unittest.TestCase):
    def test_stores_attributes_and_starts_with_no_arguments(self):
        # ``Function`` is a free function. It must keep its return type,
        # attributes and extern flag, and start with an empty argument list.
        func = typedesc_base.Function("f", "int", ["dllimport"], "extern")
        self.assertEqual(func.name, "f")
        self.assertEqual(func.returns, "int")
        self.assertEqual(func.attributes, ["dllimport"])
        self.assertEqual(func.extern, "extern")
        self.assertEqual(func.arguments, [])


class TestConstructor(unittest.TestCase):
    def test_stores_name_and_starts_with_no_arguments(self):
        # ``Constructor`` is a C++ constructor. It only carries a name and
        # starts with an empty argument list.
        ctor = typedesc_base.Constructor("C")
        self.assertEqual(ctor.name, "C")
        self.assertEqual(ctor.arguments, [])


class TestOperatorFunction(unittest.TestCase):
    def test_stores_name_and_returns(self):
        # ``OperatorFunction`` is a C++ operator overload. It must keep the
        # operator name and the return type.
        op = typedesc_base.OperatorFunction("operator+", "int")
        self.assertEqual(op.name, "operator+")
        self.assertEqual(op.returns, "int")
        self.assertEqual(op.arguments, [])


class TestFunctionType(unittest.TestCase):
    def test_stores_returns_and_attributes(self):
        # ``FunctionType`` describes a function pointer type. It must keep
        # the return type and the calling-convention attributes.
        func_type = typedesc_base.FunctionType("int", ["__stdcall__"])
        self.assertEqual(func_type.returns, "int")
        self.assertEqual(func_type.attributes, ["__stdcall__"])
        self.assertEqual(func_type.arguments, [])


class TestMethod(unittest.TestCase):
    def test_stores_name_and_returns(self):
        # ``Method`` is a C++ member function. It must keep the method name
        # and the return type.
        method = typedesc_base.Method("Do", "HRESULT")
        self.assertEqual(method.name, "Do")
        self.assertEqual(method.returns, "HRESULT")
        self.assertEqual(method.arguments, [])


class TestUnion(unittest.TestCase):
    def test_stores_name_align_size_and_artificial(self):
        # ``Union`` is a C union. ``align`` and ``size`` are converted to int
        # because the parser may pass them as strings.
        union = typedesc_base.Union("U", "32", "64", artificial=True)
        self.assertEqual(union.name, "U")
        self.assertEqual(union.align, 32)
        self.assertEqual(union.size, 64)
        self.assertTrue(union.artificial)

    def test_size_is_none_when_unknown(self):
        # An incomplete union has no known size; it must be None, not 0.
        union = typedesc_base.Union("U", 32, None)
        self.assertIsNone(union.size)

    def test_artificial_defaults_to_none(self):
        # ``artificial`` marks compiler-generated types. It defaults to None
        # when the caller does not say otherwise.
        union = typedesc_base.Union("U", 32, 64)
        self.assertIsNone(union.artificial)

    def test_starts_with_no_members_and_has_head_and_body(self):
        # A union starts empty and exposes a head/body pair used by the code
        # generator to emit the class definition. Both must point back to the
        # union itself.
        union = typedesc_base.Union("U", 32, 64)
        self.assertEqual(union.members, [])
        self.assertIs(union.get_head().struct, union)
        self.assertIs(union.get_body().struct, union)


class TestCvQualifiedType(unittest.TestCase):
    def test_stores_typ_const_and_volatile(self):
        # ``CvQualifiedType`` wraps a type with its const/volatile qualifiers.
        cv = typedesc_base.CvQualifiedType("int", True, False)
        self.assertEqual(cv.typ, "int")
        self.assertTrue(cv.const)
        self.assertFalse(cv.volatile)


class TestVariable(unittest.TestCase):
    def test_stores_name_and_type_with_no_initializer(self):
        # A global variable without an initializer has ``init`` set to None.
        var = typedesc_base.Variable("g", "int")
        self.assertEqual(var.name, "g")
        self.assertEqual(var.typ, "int")
        self.assertIsNone(var.init)

    def test_stores_initializer(self):
        # When the declaration has an initializer it must be kept.
        var = typedesc_base.Variable("g", "int", "42")
        self.assertEqual(var.init, "42")


if __name__ == "__main__":
    unittest.main()
