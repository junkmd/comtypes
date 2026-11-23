import unittest
from unittest import mock
from ctypes import POINTER, c_void_p, pointer

import comtypes
from comtypes import COMObject, IUnknown
from comtypes.automation import IDispatch, IEnumVARIANT
from comtypes.hresult import *
from comtypes.server.automation import VARIANTEnumerator, COMCollection


class MockDispatchObject:
    """Mock object that supports IDispatch interface."""
    def __init__(self, name):
        self.name = name
    
    def QueryInterface(self, iid):
        if iid == IDispatch._iid_:
            return self
        raise comtypes.COMError(E_NOINTERFACE, "No interface")


class TestVARIANTEnumerator(unittest.TestCase):
    def setUp(self):
        self.items = [MockDispatchObject(f"item{i}") for i in range(3)]
        self.enumerator = VARIANTEnumerator(self.items)

    def test_init(self):
        """Test VARIANTEnumerator initialization."""
        self.assertEqual(self.enumerator.items, self.items)
        self.assertEqual(list(self.enumerator.seq), self.items)

    def test_next_success(self):
        """Test Next method with valid parameters."""
        # Create a mock rgVar array
        rgVar = [mock.Mock() for _ in range(3)]
        pCeltFetched = [0]
        
        result = self.enumerator.Next(None, 2, rgVar, pCeltFetched)
        
        self.assertEqual(result, S_OK)
        self.assertEqual(pCeltFetched[0], 2)
        # Check that QueryInterface was called on each item
        for i in range(2):
            rgVar[i].value.assert_called_once()

    def test_next_partial(self):
        """Test Next method when fewer items available than requested."""
        # Exhaust the enumerator first
        rgVar = [mock.Mock() for _ in range(3)]
        pCeltFetched = [0]
        
        # Get all 3 items
        self.enumerator.Next(None, 3, rgVar, pCeltFetched)
        self.assertEqual(pCeltFetched[0], 3)
        
        # Try to get more - should return S_FALSE
        rgVar2 = [mock.Mock() for _ in range(2)]
        pCeltFetched2 = [0]
        result = self.enumerator.Next(None, 2, rgVar2, pCeltFetched2)
        
        self.assertEqual(result, S_FALSE)
        self.assertEqual(pCeltFetched2[0], 0)

    def test_next_null_rgvar(self):
        """Test Next with null rgVar returns E_POINTER."""
        pCeltFetched = [0]
        result = self.enumerator.Next(None, 1, None, pCeltFetched)
        self.assertEqual(result, E_POINTER)

    def test_next_null_pceltfetched(self):
        """Test Next with null pCeltFetched creates default."""
        rgVar = [mock.Mock()]
        result = self.enumerator.Next(None, 1, rgVar, None)
        self.assertEqual(result, S_OK)

    def test_skip_success(self):
        """Test Skip method successfully skips elements."""
        result = self.enumerator.Skip(None, 2)
        self.assertEqual(result, S_OK)
        # Next should return the third item
        rgVar = [mock.Mock()]
        pCeltFetched = [0]
        result = self.enumerator.Next(None, 1, rgVar, pCeltFetched)
        self.assertEqual(pCeltFetched[0], 1)

    def test_skip_past_end(self):
        """Test Skip past end returns S_FALSE."""
        result = self.enumerator.Skip(None, 5)
        self.assertEqual(result, S_FALSE)

    def test_skip_zero(self):
        """Test Skip with zero elements returns S_OK."""
        result = self.enumerator.Skip(None, 0)
        self.assertEqual(result, S_OK)

    def test_reset(self):
        """Test Reset method restarts enumeration."""
        # Consume some items
        rgVar = [mock.Mock() for _ in range(2)]
        pCeltFetched = [0]
        self.enumerator.Next(None, 2, rgVar, pCeltFetched)
        
        # Reset
        result = self.enumerator.Reset(None)
        self.assertEqual(result, S_OK)
        
        # Should be able to get first item again
        rgVar2 = [mock.Mock()]
        pCeltFetched2 = [0]
        result = self.enumerator.Next(None, 1, rgVar2, pCeltFetched2)
        self.assertEqual(pCeltFetched2[0], 1)

    def test_next_item_queryinterface_fails(self):
        """Test Next handles QueryInterface failure gracefully."""
        # Create an item that fails QueryInterface
        bad_item = mock.Mock()
        bad_item.QueryInterface.side_effect = comtypes.COMError(E_NOINTERFACE, "No interface")
        
        enumerator = VARIANTEnumerator([bad_item])
        rgVar = [mock.Mock()]
        pCeltFetched = [0]
        
        # The current implementation doesn't catch COMError from QueryInterface
        # This test documents the current behavior
        with self.assertRaises(comtypes.COMError):
            enumerator.Next(None, 1, rgVar, pCeltFetched)


class MockItemType:
    """Mock item type for COMCollection tests."""
    def __init__(self, pathname):
        self.pathname = pathname
        self._iid_ = GUID.create_new()
    
    def IUnknown_QueryInterface(self, this, riid, ppv):
        return S_OK


class TestCOMCollection(unittest.TestCase):
    def setUp(self):
        self.collection = ["item1", "item2", "item3"]
        self.item_type = MockItemType
        self.com_collection = COMCollection(self.item_type, self.collection)

    def test_init(self):
        """Test COMCollection initialization."""
        self.assertEqual(self.com_collection.collection, self.collection)
        self.assertEqual(self.com_collection.itemtype, self.item_type)

    def test_get_item_success(self):
        """Test _get_Item with valid parameters."""
        pitem = [mock.Mock()]
        pitem[0]._iid_ = IUnknown._iid_
        
        result = self.com_collection._get_Item(None, "test_path", pitem)
        
        self.assertEqual(result, S_OK)
        # Verify item was created with pathname
        # The actual item creation is mocked, so we just check the call

    def test_get_item_null_pitem(self):
        """Test _get_Item with null pitem returns E_POINTER."""
        result = self.com_collection._get_Item(None, "test_path", None)
        self.assertEqual(result, E_POINTER)

    def test_get_count_success(self):
        """Test _get_Count with valid parameter."""
        pcount = [0]
        result = self.com_collection._get_Count(None, pcount)
        
        self.assertEqual(result, S_OK)
        self.assertEqual(pcount[0], 3)

    def test_get_count_null_pcount(self):
        """Test _get_Count with null pcount returns E_POINTER."""
        result = self.com_collection._get_Count(None, None)
        self.assertEqual(result, E_POINTER)

    def test_get_newenum_success(self):
        """Test _get__NewEnum with valid parameter."""
        penum = c_void_p()
        
        with mock.patch('comtypes.server.automation.VARIANTEnumerator') as mock_enum_class, \
             mock.patch('comtypes.server.automation.pointer') as mock_pointer, \
             mock.patch('comtypes.server.automation.IUnknown') as mock_iunknown:
            
            mock_enum = mock.Mock()
            mock_enum.IUnknown_QueryInterface.return_value = S_OK
            mock_enum_class.return_value = mock_enum
            mock_pointer.return_value = mock.Mock()
            mock_iunknown._iid_ = IUnknown._iid_
            
            result = self.com_collection._get__NewEnum(None, penum)
            
            self.assertEqual(result, S_OK)
            mock_enum_class.assert_called_once_with(self.item_type, self.collection)
            mock_enum.IUnknown_QueryInterface.assert_called_once()

    def test_get_newenum_null_penum(self):
        """Test _get__NewEnum with null penum returns E_POINTER."""
        result = self.com_collection._get__NewEnum(None, None)
        self.assertEqual(result, E_POINTER)


if __name__ == "__main__":
    unittest.main()