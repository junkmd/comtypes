"""Tests for ``comtypes._tlib_version_checker``.

``_check_version`` is emitted into generated modules to detect when a cached
module no longer matches the typelib it was generated from.  These tests pin
down the three ways it can react: raising on a version mismatch, silently
returning when the typelib path is unusable, and raising when the typelib is
newer than the cached module.
"""

import unittest as ut

from comtypes._tlib_version_checker import _check_version
from comtypes.tools.codegenerator import version as required


class Test_check_version(ut.TestCase):
    def test_wrong_version_raises_import_error(self):
        # A generated module records the code generator version it was
        # produced with.  If the installed generator reports a different
        # version, the cached module must not be used, so an ImportError is
        # raised instead of silently importing a stale module.
        with self.assertRaises(ImportError):
            _check_version("a version that is not the required one")

    def test_missing_typelib_path_returns(self):
        # ``_check_version`` looks up ``typelib_path`` in the caller's
        # globals.  When it is absent, ``os.stat(None)`` raises TypeError and
        # the function returns without raising, so a module that does not
        # record a typelib path is still importable.
        namespace = {"_check_version": _check_version, "required": required}
        exec("result = _check_version(required)", namespace)
        self.assertIsNone(namespace["result"])

    def test_different_typelib_raises_import_error(self):
        # When the recorded typelib exists but no cached mtime is given, the
        # module is considered stale and an ImportError is raised so the
        # caller regenerates it.
        namespace = {
            "_check_version": _check_version,
            "required": required,
            "typelib_path": __file__,
        }
        with self.assertRaises(ImportError):
            exec("_check_version(required)", namespace)
