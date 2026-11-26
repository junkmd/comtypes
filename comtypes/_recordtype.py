"""Leaf module for the `SupportsRecordInfo` marker.

This module must not import any name from `comtypes` at runtime.  It is
imported from `comtypes._memberspec`, which itself is imported while
`comtypes/__init__.py` is still executing.  Importing `comtypes.recordinfo`
from there would fail, because `recordinfo` needs names (`BSTR`, `IUnknown`,
...) that `comtypes/__init__.py` has not defined yet at that point.
"""

from abc import ABCMeta
from ctypes import Structure
from typing import TYPE_CHECKING, Any, ClassVar, NoReturn, final

if TYPE_CHECKING:
    from comtypes import hints  # type: ignore


def _is_recordinfo_guids(obj: Any) -> "hints.TypeIs[tuple[str, int, int, int, str]]":
    if not isinstance(obj, tuple):
        return False
    if len(obj) != 5:
        return False
    tlibid, majorver, minver, lcid, tinfoid = obj
    if (
        isinstance(tlibid, str)
        and isinstance(majorver, int)
        and isinstance(minver, int)
        and isinstance(lcid, int)
        and isinstance(tinfoid, str)
    ):
        return True
    return False


_PyCStructType = type(Structure)


class _RecordTypeCheckMeta(_PyCStructType, ABCMeta):
    def __init_subclass__(cls, **kwargs: Any) -> None:
        raise TypeError(f"{cls.__name__}: '_RecordTypeCheckMeta' cannot be subclassed")

    def __subclasscheck__(cls, subclass: type[Any]) -> bool:
        if not isinstance(subclass, type):
            return False
        if not issubclass(subclass, Structure):
            return False
        # Check that `_recordinfo_` is a class attribute, not an instance
        # attribute.
        if "_recordinfo_" not in subclass.__dict__:
            return False
        return _is_recordinfo_guids(subclass._recordinfo_)

    def __instancecheck__(cls, instance: Any) -> bool:
        return cls.__subclasscheck__(type(instance))


@final
class SupportsRecordInfo(Structure, metaclass=_RecordTypeCheckMeta):
    """An abstract marker class for VT_RECORD types.

    "Supports RecordInfo" means that passing `_recordinfo_` to
    `GetRecordInfoFromGuids` returns a valid `IRecordInfo`. This class is not
    `IRecordInfo` itself or its protocol.

    Do not inherit from this class to create concrete types. Instead, use
    `isinstance` or `issubclass` to check if a class is a `Structure` that has
    a class attribute `_recordinfo_`. This attribute must be a 5-tuple:
    (type library GUID string, major version, minor version, locale ID,
    typeinfo GUID string).
    """

    _recordinfo_: ClassVar[tuple[str, int, int, int, str]]

    def __init__(self) -> NoReturn:
        raise TypeError("This cannot be instantiated")

    def __init_subclass__(cls, **kwargs: Any) -> None:
        raise TypeError(
            f"{cls.__name__}: 'SupportsRecordInfo' is an abstract marker class "
            "and cannot be subclassed; use 'isinstance' / 'issubclass' instead."
        )
