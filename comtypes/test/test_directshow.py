import contextlib
import unittest
from ctypes import POINTER, c_ulong
from typing import TYPE_CHECKING, Iterator

import comtypes.client
from comtypes import COMMETHOD, GUID, HRESULT, IUnknown
from comtypes.automation import VARIANT
from comtypes.persist import IPropertyBag

if TYPE_CHECKING:
    from comtypes import hints  # noqa  # type: ignore

with contextlib.redirect_stdout(None):  # supress warnings
    comtypes.client.GetModule("msvidctl.dll")
from comtypes.gen.MSVidCtlLib import IEnumMoniker, IMoniker

CLSID_AudioCompressorCategory = GUID("{33D9A761-90C8-11D0-BD43-00A0C911CE86}")
CLSID_SystemDeviceEnum = GUID("{62BE5D10-60EB-11d0-BD3B-00A0C911CE86}")


class ICreateDevEnum(IUnknown):
    _iid_ = GUID("{29840822-5B84-11D0-BD3B-00A0C911CE86}")
    _idlflags_ = []

    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "CreateClassEnumerator",
            (["in"], POINTER(GUID), "clsidDeviceClass"),
            (["out"], POINTER(POINTER(IEnumMoniker)), "ppenumMoniker"),
            (["in"], c_ulong, "dwFlags"),
        ),
    ]

    if TYPE_CHECKING:  # commembers

        def CreateClassEnumerator(
            self, clsidDeviceClass: GUID, dwFlags: int
        ) -> IEnumMoniker: ...


def iterate_moniker(enum_mon: IEnumMoniker) -> Iterator[IMoniker]:
    # workaround for RemoteNext
    (item, fetched) = enum_mon.RemoteNext(1)
    while fetched:
        yield item
        (item, fetched) = enum_mon.RemoteNext(1)


class Test(unittest.TestCase):
    def test(self):
        dev_enum = comtypes.client.CreateObject(
            CLSID_SystemDeviceEnum, interface=ICreateDevEnum
        )
        # https://learn.microsoft.com/en-us/windows/win32/api/strmif/nf-strmif-icreatedevenum-createclassenumerator
        # If flag is zero, the method enumerates every filter in the category.
        class_enum = dev_enum.CreateClassEnumerator(CLSID_AudioCompressorCategory, 0)
        for mon in iterate_moniker(class_enum):
            pb = mon.QueryInterface(IPropertyBag)
            name = pb.Read("FriendlyName", VARIANT(), None)
            with self.subTest(name=name):
                self.assertIsInstance(name, str)
            break  # A single item is enough to test the COM features.
        else:
            # If no item is detected, the test will be skipped.
            raise unittest.SkipTest("No item")
