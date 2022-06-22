from ctypes import POINTER
import sys
import unittest as ut

from comtypes import IUnknown
from comtypes.client._enhance import EnhancedImplementation
from comtypes.client import CreateObject, GetModule

GetModule("scrrun.dll")
from comtypes.gen import Scripting

if sys.version_info >= (3, 3):
    from collections.abc import ItemsView
else:
    from collections import ItemsView


class SampleDictionary(Scripting.IDictionary, EnhancedImplementation):
    def items(self):
        return ItemsView(self)

    def __delitem__(self, key):
        self.Remove(key)


class Test(ut.TestCase):
    def test(self):
        orig = CreateObject(Scripting.Dictionary)
        d = SampleDictionary.from_punk(orig)
        self.assertIsNot(orig, d)
        self.assertIsInstance(d, Scripting.IDictionary)
        self.assertIsInstance(d, POINTER(IUnknown))
        d["foo"] = 1
        d["bar"] = "spam ham"
        d["baz"] = 3.14
        del d["baz"]
        actual_items = d.items()
        self.assertIsInstance(actual_items, ItemsView)
        expected_items = [("foo", 1), ("bar", "spam ham")]
        self.assertEqual(list(actual_items), expected_items)


if __name__ == "__main__":
    ut.main()
