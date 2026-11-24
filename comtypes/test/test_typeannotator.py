"""Tests for ``comtypes.tools.codegenerator.typeannotator``.

These tests exercise the type annotator to improve coverage.
"""

import ast
import unittest

from comtypes.tools import typedesc
from comtypes.tools.codegenerator import typeannotator
from comtypes.tools.tlbparser import (
    BSTR_type,
    HRESULT_type,
    VARIANT_BOOL_type,
    VARIANT_type,
    void_type,
    int_type,
)

iunknown_type = typedesc.ComInterface(
    "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", ["hidden"], None
)
idispatch_type = typedesc.ComInterface(
    "IDispatch",
    iunknown_type,
    "{00020400-0000-0000-C000-000000000046}",
    ["restricted"],
    None,
)


class Test_AvoidUsingKeywords(unittest.TestCase):
    def _create_typedesc_disp_interface(self) -> typedesc.DispInterface:
        guid = "{00000000-0000-0000-0000-000000000000}"
        itf = typedesc.DispInterface("IDispDerived", idispatch_type, guid, [], None)
        ham = typedesc.DispProperty(0, "ham", BSTR_type, ["readonly"], None)
        bacon = typedesc.DispMethod(
            472, 1, "bacon", VARIANT_BOOL_type, ["hidden"], "bacon's \\ help"
        )
        bacon.add_argument(VARIANT_type, "and", ["in"], None)
        bacon.add_argument(VARIANT_type, "foo", ["in"], None)
        get_spam = typedesc.DispMethod(
            6, 2, "spam", VARIANT_type, ["propget"], "spam help"
        )
        get_spam.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        put_spam = typedesc.DispMethod(
            6, 4, "spam", void_type, ["propput"], "spam setter"
        )
        put_spam.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        put_spam.add_argument(VARIANT_type, "arg2", ["in"], None)
        except_ = typedesc.DispProperty(1, "except", BSTR_type, ["readonly"], "help")
        raise_ = typedesc.DispMethod(
            474, 1, "raise", VARIANT_BOOL_type, ["hidden"], None
        )
        raise_.add_argument(VARIANT_type, "foo", ["in"], None)
        raise_.add_argument(VARIANT_type, "bar", ["in", "optional"], None)
        get_def = typedesc.DispMethod(8, 2, "def", VARIANT_type, ["propget"], None)
        get_def.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        put_def = typedesc.DispMethod(8, 4, "def", void_type, ["propput"], None)
        put_def.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        put_def.add_argument(VARIANT_type, "arg2", ["in"], None)
        egg = typedesc.DispMethod(643, 1, "egg", VARIANT_BOOL_type, [], None)
        for m in [
            ham,
            bacon,
            get_spam,
            put_spam,
            except_,
            raise_,
            get_def,
            put_def,
            egg,
        ]:
            itf.add_member(m)
        return itf

    def test_disp_interface(self):
        itf = self._create_typedesc_disp_interface()
        expected = (
            "        @property  # dispprop\n"
            "        def ham(self) -> hints.Incomplete: ...\n"
            "        pass  # @property  # dispprop\n"
            '        pass  # avoid using a keyword for def except(self) -> hints.Incomplete: """help"""; ...\n'  # noqa
            '        def bacon(self, *args: hints.Any, **kwargs: hints.Any) -> hints.Incomplete: """bacon\'s \\\\ help"""; ...\n'  # noqa
            '        def _get_spam(self, arg1: hints.Incomplete = ..., /) -> hints.Incomplete: """spam help"""; ...\n'  # noqa
            "        def _set_spam(self, arg1: hints.Incomplete = ..., /, *args: hints.Unpack[tuple[hints.Incomplete]]) -> hints.Incomplete: ...\n"  # noqa
            '        spam = hints.named_property(\'spam\', _get_spam, _set_spam); """spam help"""\n'
            "        pass  # avoid using a keyword for def raise(self, foo: hints.Incomplete, bar: hints.Incomplete = ..., /) -> hints.Incomplete: ...\n"  # noqa
            "        def _get_def(self, arg1: hints.Incomplete = ..., /) -> hints.Incomplete: ...\n"  # noqa
            "        def _set_def(self, arg1: hints.Incomplete = ..., /, *args: hints.Unpack[tuple[hints.Incomplete]]) -> hints.Incomplete: ...\n"  # noqa
            "        pass  # avoid using a keyword for def = hints.named_property('def', _get_def, _set_def)\n"  # noqa
            "        def egg(self) -> hints.Incomplete: ..."  # noqa
        )
        self.assertEqual(
            expected, typeannotator.DispInterfaceMembersAnnotator(itf).generate()
        )

    def test_valid_syntax_dispmethods(self):
        itf = self._create_typedesc_disp_interface()
        definition = "\n".join(
            (
                "class ISomeInterface(IDispatch):",
                "    if TYPE_CHECKING:",
                f"{typeannotator.DispInterfaceMembersAnnotator(itf).generate()}",
            )
        )
        ast.parse(definition, mode="exec")

    def _create_typedesc_com_interface(self) -> typedesc.ComInterface:
        guid = "{00000000-0000-0000-0000-000000000000}"
        itf = typedesc.ComInterface(
            "IUnkDerived", iunknown_type, guid, ["hidden"], None
        )
        spam = typedesc.ComMethod(
            2, 1610678269, "spam", HRESULT_type, ["propget"], None
        )
        get_ham = typedesc.ComMethod(
            2, 1610678270, "ham", HRESULT_type, ["propget"], "ham help"
        )
        put_ham = typedesc.ComMethod(
            4, 1610678270, "ham", HRESULT_type, ["propput"], "ham setter"
        )
        put_ham.add_argument(VARIANT_type, "arg1", ["in"], None)
        bacon = typedesc.ComMethod(
            1, 1610678271, "bacon", HRESULT_type, [], "bacon's \\ help"
        )
        bacon.add_argument(VARIANT_type, "foo", ["in"], None)
        bacon.add_argument(VARIANT_type, "or", ["in"], None)
        global_ = typedesc.ComMethod(
            2, 1610678272, "global", HRESULT_type, ["propget"], None
        )
        get_class = typedesc.ComMethod(
            2, 1610678273, "class", HRESULT_type, ["propget"], None
        )
        get_class.add_argument(VARIANT_type, "arg1", ["in"], None)
        put_class = typedesc.ComMethod(
            4, 1610678273, "class", HRESULT_type, ["propput"], None
        )
        put_class.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        put_class.add_argument(VARIANT_type, "arg2", ["in"], None)
        pass_ = typedesc.ComMethod(1, 1610678274, "pass", HRESULT_type, [], None)
        pass_.add_argument(VARIANT_type, "foo", ["in"], None)
        pass_.add_argument(VARIANT_type, "bar", ["in", "optional"], None)
        members = [spam, get_ham, put_ham, bacon, global_, get_class, put_class, pass_]
        itf.extend_members(members)
        return itf

    def test_com_interface(self):
        itf = self._create_typedesc_com_interface()
        expected = (
            "        def _get_spam(self) -> hints.Hresult: ...\n"
            "        spam = hints.normal_property(_get_spam)\n"
            '        def _get_ham(self) -> hints.Hresult: """ham help"""; ...\n'
            "        def _set_ham(self, arg1: hints.Incomplete) -> hints.Hresult: ...\n"
            '        ham = hints.normal_property(_get_ham, _set_ham); """ham help"""\n'
            '        def bacon(self, *args: hints.Any, **kwargs: hints.Any) -> hints.Hresult: """bacon\'s \\\\ help"""; ...\n'  # noqa
            "        def _get_global(self) -> hints.Hresult: ...\n"
            "        pass  # avoid using a keyword for global = hints.normal_property(_get_global)\n"  # noqa
            "        def _get_class(self, arg1: hints.Incomplete) -> hints.Hresult: ...\n"
            "        def _set_class(self, arg1: hints.Incomplete = ..., /, *args: hints.Unpack[tuple[hints.Incomplete]]) -> hints.Hresult: ...\n"  # noqa
            "        pass  # avoid using a keyword for class = hints.named_property('class', _get_class, _set_class)\n"  # noqa
            "        pass  # avoid using a keyword for def pass(self, foo: hints.Incomplete, bar: hints.Incomplete = ...) -> hints.Hresult: ..."  # noqa
        )
        self.assertEqual(
            expected, typeannotator.ComInterfaceMembersAnnotator(itf).generate()
        )

    def test_valid_syntax_commethods(self):
        itf = self._create_typedesc_com_interface()
        definition = "\n".join(
            (
                "class ISomeInterface(IUnknown):",
                "    if TYPE_CHECKING:",
                f"{typeannotator.ComInterfaceMembersAnnotator(itf).generate()}",
            )
        )
        ast.parse(definition, mode="exec")


class TestComMethodAnnotator(unittest.TestCase):
    def test_com_method_annotator_basic(self):
        mth = typedesc.ComMethod(
            1, 1610678272, "TestMethod", HRESULT_type, [], "test doc"
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in", "optional"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)
        self.assertIn("arg1: hints.Incomplete", result)
        self.assertIn("arg2: hints.Incomplete = ...", result)
        self.assertIn("-> hints.Hresult", result)
        self.assertIn('"""test doc"""', result)

    def test_com_method_annotator_with_out_args(self):
        mth = typedesc.ComMethod(
            1, 1610678272, "TestMethod", HRESULT_type, [], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        mth.add_argument(VARIANT_type, "arg2", ["out"], None)
        mth.add_argument(VARIANT_type, "arg3", ["in", "out"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("->", result)
        # Should have tuple return type for multiple out args
        self.assertIn("tuple[", result)

    def test_com_method_annotator_keyword_name(self):
        mth = typedesc.ComMethod(
            1, 1610678272, "class", HRESULT_type, [], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("class")
        self.assertIn("pass  # avoid using a keyword", result)

    def test_com_method_annotator_keyword_arg(self):
        mth = typedesc.ComMethod(
            1, 1610678272, "TestMethod", HRESULT_type, [], None
        )
        mth.add_argument(VARIANT_type, "class", ["in"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("*args: hints.Any", result)
        self.assertIn("**kwargs: hints.Any", result)

    def test_com_method_annotator_put(self):
        mth = typedesc.ComMethod(
            4, 1610678272, "TestMethod", HRESULT_type, ["propput"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)

    def test_com_method_annotator_putref(self):
        mth = typedesc.ComMethod(
            8, 1610678272, "TestMethod", HRESULT_type, ["propputref"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)

    def test_com_method_annotator_required_after_optional(self):
        # Test the case where required params follow optional ones
        mth = typedesc.ComMethod(
            4, 1610678272, "TestMethod", HRESULT_type, ["propput"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in"], None)  # required after optional
        anno = typeannotator.ComMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("/", result)
        self.assertIn("Unpack[tuple[", result)


class TestDispMethodAnnotator(unittest.TestCase):
    def test_disp_method_annotator_basic(self):
        mth = typedesc.DispMethod(
            1, 1, "TestMethod", VARIANT_type, [], "test doc"
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in", "optional"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)
        self.assertIn("arg1: hints.Incomplete", result)
        self.assertIn("arg2: hints.Incomplete = ...", result)
        self.assertIn("-> hints.Incomplete", result)
        self.assertIn('"""test doc"""', result)

    def test_disp_method_annotator_keyword_name(self):
        mth = typedesc.DispMethod(
            1, 1, "class", VARIANT_type, [], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("class")
        self.assertIn("pass  # avoid using a keyword", result)

    def test_disp_method_annotator_keyword_arg(self):
        mth = typedesc.DispMethod(
            1, 1, "TestMethod", VARIANT_type, [], None
        )
        mth.add_argument(VARIANT_type, "class", ["in"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("*args: hints.Any", result)
        self.assertIn("**kwargs: hints.Any", result)

    def test_disp_method_annotator_put(self):
        mth = typedesc.DispMethod(
            4, 1, "TestMethod", void_type, ["propput"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)

    def test_disp_method_annotator_putref(self):
        mth = typedesc.DispMethod(
            8, 1, "TestMethod", void_type, ["propputref"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("def TestMethod(", result)

    def test_disp_method_annotator_required_after_optional(self):
        mth = typedesc.DispMethod(
            4, 1, "TestMethod", void_type, ["propput"], None
        )
        mth.add_argument(VARIANT_type, "arg1", ["in", "optional"], None)
        mth.add_argument(VARIANT_type, "arg2", ["in"], None)
        anno = typeannotator.DispMethodAnnotator(mth)
        result = anno.getvalue("TestMethod")
        self.assertIn("/", result)
        self.assertIn("Unpack[tuple[", result)


class TestToOutType(unittest.TestCase):
    def test_to_outtype_pointer(self):
        ptr = typedesc.PointerType(int_type, 64, 64)
        result = typeannotator._to_outtype(ptr)
        self.assertEqual(result, "hints.Incomplete")

    def test_to_outtype_dispinterface(self):
        itf = typedesc.DispInterface("IDispTest", idispatch_type, "{}", [], None)
        result = typeannotator._to_outtype(itf)
        self.assertEqual(result, "'IDispTest'")

    def test_to_outtype_cominterface(self):
        itf = typedesc.ComInterface("IComTest", iunknown_type, "{}", [], None)
        result = typeannotator._to_outtype(itf)
        self.assertEqual(result, "'IComTest'")

    def test_to_outtype_coclass(self):
        from comtypes.typeinfo import TLIBATTR
        tlibattr = TLIBATTR()
        tlibattr.guid = "{}"
        tlibattr.wMajorVerNum = 1
        tlibattr.wMinorVerNum = 0
        tlibattr.lcid = 0
        tlibattr.wLibFlags = 0
        tlibattr.syskind = 1

        coclass = typedesc.CoClass("TestClass", "{}", [], tlibattr, None)
        itf = typedesc.ComInterface("IComTest", iunknown_type, "{}", [], None)
        coclass.add_interface(itf, 0)
        result = typeannotator._to_outtype(coclass)
        self.assertIn("hints.Annotated[", result)
        self.assertIn("hints.FirstComItfOf", result)

    def test_to_outtype_default(self):
        result = typeannotator._to_outtype(int_type)
        self.assertEqual(result, "hints.Incomplete")


class TestGenerateTrailingParams(unittest.TestCase):
    def test_generate_trailing_params(self):
        specs = [(int_type, "arg1", None), (int_type, "arg2", None)]
        result = typeannotator._generate_trailing_params(specs)
        self.assertIn("*args: hints.Unpack[tuple[", result)
        self.assertIn("hints.Incomplete", result)


class TestToDocstr(unittest.TestCase):
    def test_to_docstr_basic(self):
        result = typeannotator._to_docstr("simple doc")
        self.assertEqual(result, '"""simple doc"""')

    def test_to_docstr_with_backslash(self):
        result = typeannotator._to_docstr("path\\to\\file")
        self.assertEqual(result, '"""path\\\\to\\\\file"""')

    def test_to_docstr_with_quotes(self):
        result = typeannotator._to_docstr('say "hello"')
        self.assertEqual(result, '"""say \\"hello\\""""')


class TestMethodsAnnotatorIterMethods(unittest.TestCase):
    def test_iter_methods_categorization(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        m1 = typedesc.ComMethod(1, 1, "Method1", HRESULT_type, [], None)
        m2 = typedesc.ComMethod(2, 2, "Prop1", HRESULT_type, ["propget"], None)
        m3 = typedesc.ComMethod(4, 2, "Prop1", HRESULT_type, ["propput"], None)
        m3.add_argument(VARIANT_type, "arg1", ["in"], None)
        m4 = typedesc.ComMethod(8, 3, "Prop2", HRESULT_type, ["propputref"], None)
        m4.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([m1, m2, m3, m4])

        annotator = typeannotator.ComMethodsAnnotator()
        methods = list(annotator._iter_methods(itf.members))
        self.assertEqual(len(methods), 3)  # Method1, Prop1, Prop2

        # Check Method1 (regular method)
        name, fmth, fget, fput, fputref = methods[0]
        self.assertEqual(name, "Method1")
        self.assertIsNotNone(fmth)
        self.assertIsNone(fget)
        self.assertIsNone(fput)
        self.assertIsNone(fputref)

        # Check Prop1 (get + put)
        name, fmth, fget, fput, fputref = methods[1]
        self.assertEqual(name, "Prop1")
        self.assertIsNone(fmth)
        self.assertIsNotNone(fget)
        self.assertIsNotNone(fput)
        self.assertIsNone(fputref)

        # Check Prop2 (putref only)
        name, fmth, fget, fput, fputref = methods[2]
        self.assertEqual(name, "Prop2")
        self.assertIsNone(fmth)
        self.assertIsNone(fget)
        self.assertIsNone(fput)
        self.assertIsNotNone(fputref)


class TestPatchDunder(unittest.TestCase):
    def test_patch_dunder_count(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._patch_dunder("Count")
        self.assertIn("__len__ = hints.to_dunder_len(Count)", annotator.data)

    def test_patch_dunder_item(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._patch_dunder("Item")
        self.assertIn("__call__ = hints.to_dunder_call(Item)", annotator.data)
        self.assertIn("__getitem__ = hints.to_dunder_getitem(Item)", annotator.data)
        self.assertIn("__setitem__ = hints.to_dunder_setitem(Item)", annotator.data)

    def test_patch_dunder_newenum(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._patch_dunder("_NewEnum")
        self.assertIn("__iter__ = hints.to_dunder_iter(_NewEnum)", annotator.data)

    def test_patch_dunder_other(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._patch_dunder("Other")
        # Should not add any dunder methods
        self.assertEqual(len(annotator.data), 0)


class TestDefineNamedProp(unittest.TestCase):
    def test_define_named_prop_getter_setter(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_named_prop("TestProp", "_get_TestProp", "_set_TestProp", "doc")
        self.assertIn("TestProp = hints.named_property('TestProp', _get_TestProp, _set_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_named_prop_getter_only(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_named_prop("TestProp", "_get_TestProp", None, "doc")
        self.assertIn("TestProp = hints.named_property('TestProp', _get_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_named_prop_setter_only(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_named_prop("TestProp", None, "_set_TestProp", "doc")
        self.assertIn("TestProp = hints.named_property('TestProp', fset=_set_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_named_prop_keyword(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_named_prop("class", "_get_class", "_set_class", "doc")
        self.assertIn("pass  # avoid using a keyword", annotator.data[0])

    def test_define_named_prop_no_doc(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_named_prop("TestProp", "_get_TestProp", "_set_TestProp", None)
        self.assertNotIn("doc", annotator.data[0])


class TestDefineNormalProp(unittest.TestCase):
    def test_define_normal_prop_getter_setter(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_normal_prop("TestProp", "_get_TestProp", "_set_TestProp", "doc")
        self.assertIn("TestProp = hints.normal_property(_get_TestProp, _set_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_normal_prop_getter_only(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_normal_prop("TestProp", "_get_TestProp", None, "doc")
        self.assertIn("TestProp = hints.normal_property(_get_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_normal_prop_setter_only(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_normal_prop("TestProp", None, "_set_TestProp", "doc")
        self.assertIn("TestProp = hints.normal_property(fset=_set_TestProp); \"\"\"doc\"\"\"", annotator.data[0])

    def test_define_normal_prop_keyword(self):
        annotator = typeannotator.ComMethodsAnnotator()
        annotator._define_normal_prop("class", "_get_class", "_set_class", "doc")
        self.assertIn("pass  # avoid using a keyword", annotator.data[0])


class TestGenPropGet(unittest.TestCase):
    def test_gen_prop_get_with_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        fget.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fget])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get("Prop", fget)
        # Should use named_property when there are inargs
        self.assertTrue(any("named_property" in d for d in annotator.data))

    def test_gen_prop_get_without_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        itf.extend_members([fget])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get("Prop", fget)
        # Should use normal_property when no inargs
        self.assertTrue(any("normal_property" in d for d in annotator.data))


class TestGenPropGetPut(unittest.TestCase):
    def test_gen_prop_get_put_with_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        fget.add_argument(VARIANT_type, "arg1", ["in"], None)
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        fput.add_argument(VARIANT_type, "arg2", ["in"], None)
        itf.extend_members([fget, fput])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get_put("Prop", fget, fput)
        self.assertTrue(any("named_property" in d for d in annotator.data))

    def test_gen_prop_get_put_without_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fget, fput])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get_put("Prop", fget, fput)
        self.assertTrue(any("normal_property" in d for d in annotator.data))


class TestGenPropGetPutref(unittest.TestCase):
    def test_gen_prop_get_putref(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        fputref = typedesc.ComMethod(8, 1, "Prop", HRESULT_type, ["propputref"], None)
        fputref.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fget, fputref])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get_putref("Prop", fget, fputref)
        self.assertTrue(any("put_or_putref" in d or "named_property" in d for d in annotator.data))


class TestGenPropGetPutPutref(unittest.TestCase):
    def test_gen_prop_get_put_putref(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fget = typedesc.ComMethod(2, 1, "Prop", HRESULT_type, ["propget"], "doc")
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        fputref = typedesc.ComMethod(8, 1, "Prop", HRESULT_type, ["propputref"], None)
        fputref.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fget, fput, fputref])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_get_put_putref("Prop", fget, fput, fputref)
        self.assertTrue(any("put_or_putref" in d for d in annotator.data))


class TestGenPropPut(unittest.TestCase):
    def test_gen_prop_put_with_multiple_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        fput.add_argument(VARIANT_type, "arg2", ["in"], None)
        itf.extend_members([fput])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_put("Prop", fput)
        self.assertTrue(any("named_property" in d for d in annotator.data))

    def test_gen_prop_put_with_single_inarg(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fput])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_put("Prop", fput)
        self.assertTrue(any("normal_property" in d for d in annotator.data))


class TestGenPropPutref(unittest.TestCase):
    def test_gen_prop_putref_with_multiple_inargs(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fputref = typedesc.ComMethod(8, 1, "Prop", HRESULT_type, ["propputref"], None)
        fputref.add_argument(VARIANT_type, "arg1", ["in"], None)
        fputref.add_argument(VARIANT_type, "arg2", ["in"], None)
        itf.extend_members([fputref])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_putref("Prop", fputref)
        self.assertTrue(any("named_property" in d for d in annotator.data))

    def test_gen_prop_putref_with_single_inarg(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fputref = typedesc.ComMethod(8, 1, "Prop", HRESULT_type, ["propputref"], None)
        fputref.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.extend_members([fputref])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_putref("Prop", fputref)
        self.assertTrue(any("normal_property" in d for d in annotator.data))


class TestGenPropPutPutref(unittest.TestCase):
    def test_gen_prop_put_putref(self):
        itf = typedesc.ComInterface("ITest", iunknown_type, "{}", [], None)
        fput = typedesc.ComMethod(4, 1, "Prop", HRESULT_type, ["propput"], None)
        fput.add_argument(VARIANT_type, "arg1", ["in"], None)
        fput.add_argument(VARIANT_type, "arg2", ["in"], None)
        fputref = typedesc.ComMethod(8, 1, "Prop", HRESULT_type, ["propputref"], None)
        fputref.add_argument(VARIANT_type, "arg1", ["in"], None)
        fputref.add_argument(VARIANT_type, "arg2", ["in"], None)
        itf.extend_members([fput, fputref])

        annotator = typeannotator.ComMethodsAnnotator()
        annotator._gen_prop_put_putref("Prop", fput, fputref)
        self.assertTrue(any("put_or_putref" in d for d in annotator.data))


class TestDispInterfaceMembersAnnotator(unittest.TestCase):
    def test_categorize_members(self):
        itf = typedesc.DispInterface("IDispTest", idispatch_type, "{}", [], None)
        prop = typedesc.DispProperty(1, "TestProp", BSTR_type, ["readonly"], "doc")
        mth = typedesc.DispMethod(2, 1, "TestMethod", VARIANT_type, [], None)
        mth.add_argument(VARIANT_type, "arg1", ["in"], None)
        itf.add_member(prop)
        itf.add_member(mth)

        annotator = typeannotator.DispInterfaceMembersAnnotator(itf)
        props, methods = annotator._categorize_members()
        self.assertEqual(len(props), 1)
        self.assertEqual(len(methods), 1)
        self.assertEqual(props[0].name, "TestProp")
        self.assertEqual(methods[0].name, "TestMethod")

    def test_generate_dispinterface(self):
        itf = typedesc.DispInterface("IDispTest", idispatch_type, "{}", [], None)
        prop = typedesc.DispProperty(1, "TestProp", BSTR_type, ["readonly"], "doc")
        mth = typedesc.DispMethod(2, 1, "TestMethod", VARIANT_type, [], None)
        itf.add_member(prop)
        itf.add_member(mth)

        annotator = typeannotator.DispInterfaceMembersAnnotator(itf)
        result = annotator.generate()
        self.assertIn("@property  # dispprop", result)
        self.assertIn("def TestProp(self)", result)
        self.assertIn("def TestMethod(", result)


if __name__ == "__main__":
    unittest.main()