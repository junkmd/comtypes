import ctypes
import weakref
from ctypes import POINTER, OleDLL, Structure, byref, c_void_p, c_wchar_p
from ctypes.wintypes import BOOL, LCID, ULONG
from typing import TYPE_CHECKING, Any, TypeVar

from comtypes import BSTR, COMMETHOD, GUID, HRESULT, IUnknown, _safearray
from comtypes._automation import VARIANT
from comtypes._recordtype import SupportsRecordInfo  # noqa: F401
from comtypes.typeinfo import ITypeInfo

if TYPE_CHECKING:
    from ctypes import _Pointer

    from comtypes import hints  # type: ignore


class IRecordInfo(IUnknown):
    # C:/vc98/include/OAIDL.H 5974
    _iid_ = GUID("{0000002F-0000-0000-C000-000000000046}")

    def GetFieldNames(self, *args: Any) -> list[str | None]:
        count = ctypes.c_ulong()
        self.__com_GetFieldNames(count, None)  # type: ignore
        array = (BSTR * count.value)()
        self.__com_GetFieldNames(count, array)  # type: ignore
        result = array[:]
        # XXX Should SysFreeString the array contents. How to?
        return result

    if TYPE_CHECKING:
        # def RecordInit
        def RecordClear(self, pvExisting: hints.Incomplete) -> hints.Hresult: ...
        def RecordCopy(
            self, pvExisting: hints.Incomplete, pvNew: hints.Incomplete
        ) -> hints.Hresult: ...
        def GetGuid(self) -> GUID: ...
        def GetName(self) -> str: ...
        def GetSize(self) -> int: ...
        def GetTypeInfo(self) -> ITypeInfo: ...
        # def GetField
        # def GetFieldNoCopy
        # def PutField
        # def PutFieldNoCopy
        def IsMatchingType(self, value: "IRecordInfo") -> bool: ...
        # def RecordCreate
        def RecordCreateCopy(self, pvSource: hints.Incomplete) -> int: ...
        def RecordDestroy(self, pvRecord: hints.Incomplete) -> hints.Hresult: ...


IRecordInfo._methods_ = [
    COMMETHOD([], HRESULT, "RecordInit", (["in"], c_void_p, "pvNew")),
    COMMETHOD([], HRESULT, "RecordClear", (["in"], c_void_p, "pvExisting")),
    COMMETHOD(
        [],
        HRESULT,
        "RecordCopy",
        (["in"], c_void_p, "pvExisting"),
        (["in"], c_void_p, "pvNew"),
    ),
    COMMETHOD([], HRESULT, "GetGuid", (["out"], POINTER(GUID), "pguid")),
    COMMETHOD([], HRESULT, "GetName", (["out"], POINTER(BSTR), "pbstrName")),
    COMMETHOD([], HRESULT, "GetSize", (["out"], POINTER(ctypes.c_ulong), "pcbSize")),
    COMMETHOD(
        [], HRESULT, "GetTypeInfo", (["out"], POINTER(POINTER(ITypeInfo)), "ppTypeInfo")
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetField",
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["out"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetFieldNoCopy",
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["out"], POINTER(VARIANT), "pvarField"),
        (["out"], POINTER(c_void_p), "ppvDataCArray"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "PutField",
        (["in"], ctypes.c_ulong, "wFlags"),
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["in"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "PutFieldNoCopy",
        (["in"], ctypes.c_ulong, "wFlags"),
        (["in"], c_void_p, "pvData"),
        (["in"], c_wchar_p, "szFieldName"),
        (["in"], POINTER(VARIANT), "pvarField"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "GetFieldNames",
        (["in", "out"], POINTER(ctypes.c_ulong), "pcNames"),
        (["in"], POINTER(BSTR), "rgBstrNames"),
    ),
    COMMETHOD([], BOOL, "IsMatchingType", (["in"], POINTER(IRecordInfo))),
    COMMETHOD([], HRESULT, "RecordCreate"),
    COMMETHOD(
        [],
        HRESULT,
        "RecordCreateCopy",
        (["in"], c_void_p, "pvSource"),
        (["out"], POINTER(c_void_p), "ppvDest"),
    ),
    COMMETHOD([], HRESULT, "RecordDestroy", (["in"], c_void_p, "pvRecord")),
]


################################################################
# functions
_oleaut32 = OleDLL("oleaut32")

_GetRecordInfoFromTypeInfo = _oleaut32.GetRecordInfoFromTypeInfo
_GetRecordInfoFromTypeInfo.argtypes = [
    POINTER(ITypeInfo),
    POINTER(POINTER(IRecordInfo)),
]
_GetRecordInfoFromTypeInfo.restype = HRESULT

_GetRecordInfoFromGuids = _oleaut32.GetRecordInfoFromGuids
_GetRecordInfoFromGuids.argtypes = [
    POINTER(GUID),
    ULONG,
    ULONG,
    LCID,
    POINTER(GUID),
    POINTER(POINTER(IRecordInfo)),
]
_GetRecordInfoFromGuids.restype = HRESULT


def GetRecordInfoFromTypeInfo(tinfo: ITypeInfo) -> IRecordInfo:
    """Return an IRecordInfo pointer to the UDT described in tinfo"""
    ri = POINTER(IRecordInfo)()
    _GetRecordInfoFromTypeInfo(tinfo, byref(ri))
    return ri  # type: ignore


def GetRecordInfoFromGuids(
    rGuidTypeLib: str, verMajor: int, verMinor: int, lcid: int, rGuidTypeInfo: str
) -> IRecordInfo:
    ri = POINTER(IRecordInfo)()
    _GetRecordInfoFromGuids(
        byref(GUID(rGuidTypeLib)),
        verMajor,
        verMinor,
        lcid,
        byref(GUID(rGuidTypeInfo)),
        byref(ri),
    )
    return ri  # type: ignore


_SafeArrayGetRecordInfo = _oleaut32.SafeArrayGetRecordInfo
_SafeArrayGetRecordInfo.restype = HRESULT
_SafeArrayGetRecordInfo.argtypes = [POINTER(_safearray.SAFEARRAY), POINTER(c_void_p)]


def SafeArrayGetRecordInfo(pa: "_Pointer[_safearray.tagSAFEARRAY]") -> IRecordInfo:
    result = POINTER(IRecordInfo)()
    _SafeArrayGetRecordInfo(pa, byref(result))
    return result  # type: ignore


################################################################
# Ownership of records returned through [out] parameters
#
# A COM vtable method that returns a record through an [out] parameter
# transfers ownership of the record's managed resources (BSTR, interface
# pointers) to the caller.  The caller must release them with
# `IRecordInfo::RecordClear` when it is done with the record.
#
# comtypes does not know when a Python record object is destroyed, so it
# cannot call `RecordClear` from a destructor.  Instead, `adopt_out_record`
# moves the record's contents into a buffer that is owned by a token, and
# attaches a `weakref.finalize` to that token.  When the returned record
# (and therefore the token) is garbage collected, the finalizer calls
# `RecordClear` on the buffer, releasing the resources exactly once.
#
# The original record is zeroed so that it no longer owns the resources;
# this prevents a double release if the original is cleared as well.

_RecT = TypeVar("_RecT", bound=Structure)


class _OwnerToken:
    """A weak-referenceable marker. Its death triggers `RecordClear`."""


def _record_clear(ri: IRecordInfo, buf: Any) -> None:
    ri.RecordClear(ctypes.addressof(buf))


def adopt_out_record(rec: _RecT) -> _RecT:
    """Return a copy of `rec` that releases its resources when destroyed.

    `rec` is the record returned by an [out] parameter.  Its contents are
    moved into a new buffer, the original is zeroed, and a finalizer is
    attached to the new record so that `IRecordInfo::RecordClear` runs when
    the new record is garbage collected.
    """
    cls = type(rec)
    size = ctypes.sizeof(cls)
    buf = (ctypes.c_char * size)()
    ctypes.memmove(buf, ctypes.addressof(rec), size)
    ctypes.memset(ctypes.addressof(rec), 0, size)
    owned = cls.from_buffer(buf)
    token = _OwnerToken()
    owned._owner_token = token  # type: ignore
    ri = GetRecordInfoFromGuids(*cls._recordinfo_)  # type: ignore
    fin = weakref.finalize(token, _record_clear, ri, buf)
    fin.atexit = False
    return owned
