"""Tests for ``comtypes._meta``."""

import unittest as ut

from comtypes._meta import _coclass_from_param


class _DummyCoclass:
    # ``_coclass_from_param`` only inspects ``_com_interfaces_`` and the class
    # itself, so a minimal stand-in is enough to exercise it.
    _com_interfaces_ = [int]


class Test_coclass_from_param(ut.TestCase):
    def test_returns_object_that_is_an_interface_instance(self):
        # ``from_param`` is used by ctypes to convert arguments passed to a
        # POINTER(coclass) parameter.  An object that already is an instance
        # of the coclass' default interface is passed through unchanged.
        obj = 42
        self.assertIs(_coclass_from_param(_DummyCoclass, obj), obj)

    def test_returns_object_that_is_a_coclass_instance(self):
        # An instance of the coclass itself is also accepted unchanged.
        obj = _DummyCoclass()
        self.assertIs(_coclass_from_param(_DummyCoclass, obj), obj)

    def test_raises_type_error_for_unrelated_object(self):
        # Anything else cannot be converted to the coclass pointer, so a
        # TypeError is raised instead of silently passing a wrong value.
        with self.assertRaises(TypeError):
            _coclass_from_param(_DummyCoclass, "not a coclass")
