"""Tests for ``comtypes.tools.codegenerator.modulenamer``.

The module namer decides the import path of a generated wrapper module. If it
returns the wrong name, the generated code imports a module that does not
exist.
"""

import unittest

import comtypes
from comtypes.tools.codegenerator import modulenamer


class _TypeLibWithDoc:
    def GetDocumentation(self, index):
        return ("MyLib", None, None, None)


class _TypeLibWithoutDoc:
    def GetDocumentation(self, index):
        raise comtypes.COMError(-2147467259, "no documentation", None)


class Test_name_friendly_module(unittest.TestCase):
    def test_returns_module_name_from_documentation(self):
        # The friendly module name is derived from the typelib's documentation
        # string.
        self.assertEqual(
            modulenamer.name_friendly_module(_TypeLibWithDoc()),
            "comtypes.gen.MyLib",
        )

    def test_returns_none_when_documentation_is_unavailable(self):
        # A typelib without documentation raises COMError; the namer must
        # return None instead of propagating the error, so callers can fall
        # back to the wrapper module name.
        self.assertIsNone(modulenamer.name_friendly_module(_TypeLibWithoutDoc()))


if __name__ == "__main__":
    unittest.main()
