# comtypes._meta helper module
from ctypes import POINTER, c_void_p, cast
import sys
from typing import Any, Dict, Tuple, Type

import comtypes

################################################################
# metaclass for CoClass (in comtypes/__init__.py)


def _wrap_coclass(self):
    # We are an IUnknown pointer, represented as a c_void_p instance,
    # but we really want this interface:
    itf = self._com_interfaces_[0]
    punk = cast(self, POINTER(itf))
    result = punk.QueryInterface(itf)
    result.__dict__["__clsid"] = str(self._reg_clsid_)
    return result


def _coclass_from_param(cls, obj):
    if isinstance(obj, (cls._com_interfaces_[0], cls)):
        return obj
    raise TypeError(obj)


#
# The mro() of a POINTER(App) type, where class App is a subclass of CoClass:
#
#  POINTER(App)
#   App
#    CoClass
#     c_void_p
#      _SimpleCData
#       _CData
#        object


class _coclass_meta(type):
    # metaclass for CoClass

    # The reason this version bridge is necessary:
    # https://docs.python.org/3.13/whatsnew/3.13.html#ctypes
    if sys.version_info >= (3, 13):

        def __new__(cls, name, bases, namespace):
            return cls.__new(name, bases, namespace)

        def __init__(self, name, bases, namespace):
            self.__init(name, bases, namespace)

    else:

        def __new__(cls, name, bases, namespace):
            self = cls.__new(name, bases, namespace)
            self.__init(name, bases, namespace)
            return self

    # When a CoClass subclass is created, create a POINTER(...) type
    # for that class, with bases <coclass> and c_void_p.  Also, the
    # POINTER(...) type gets a __ctypes_from_outparam__ method which
    # will QueryInterface for the default interface: the first one on
    # the coclass' _com_interfaces_ list.
    @classmethod
    def __new(cls, name: str, bases: Tuple[Type, ...], namespace: Dict[str, Any]):
        return type.__new__(cls, name, bases, namespace)

    def __init(
        self, name: str, bases: Tuple[Type, ...], namespace: Dict[str, Any]
    ) -> None:
        if bases == (object,):
            return
        # XXX We should insist that a _reg_clsid_ is present.
        if "_reg_clsid_" in namespace:
            clsid = namespace["_reg_clsid_"]
            comtypes.com_coclass_registry[str(clsid)] = self
        PTR = _coclass_pointer_meta(
            f"POINTER({self.__name__})",
            (self, c_void_p),
            {
                "__ctypes_from_outparam__": _wrap_coclass,
                "from_param": classmethod(_coclass_from_param),
            },
        )
        from ctypes import _pointer_type_cache  # type: ignore

        _pointer_type_cache[self] = PTR


# will not work if we change the order of the two base classes!
class _coclass_pointer_meta(type(c_void_p), _coclass_meta):
    pass
