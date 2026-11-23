import unittest
from unittest import mock
from ctypes import c_void_p, pointer
from ctypes.wintypes import DWORD

import comtypes
from comtypes import COMObject, GUID, IUnknown
from comtypes.automation import IDispatch
from comtypes.connectionpoints import IConnectionPoint
from comtypes.hresult import *
from comtypes.server.connectionpoints import ConnectionPointImpl, ConnectableObjectMixin
from _ctypes import COMError


class MockTypeInfo:
    """Mock ITypeInfo for testing."""
    def GetIDsOfNames(self, name):
        return [123]  # dispid


class MockSinkInterface:
    """Mock sink interface."""
    _iid_ = GUID.create_new()
    
    def __init__(self):
        self.Invoke = mock.Mock()


class TestConnectionPointImpl(unittest.TestCase):
    def setUp(self):
        self.sink_interface = MockSinkInterface
        self.sink_typeinfo = MockTypeInfo()
        self.conn_point = ConnectionPointImpl(self.sink_interface, self.sink_typeinfo)

    def test_init(self):
        """Test ConnectionPointImpl initialization."""
        self.assertEqual(self.conn_point._connections, {})
        self.assertEqual(self.conn_point._cookie, 0)
        self.assertEqual(self.conn_point._sink_interface, self.sink_interface)
        self.assertEqual(self.conn_point._typeinfo, self.sink_typeinfo)

    def test_advise_success(self):
        """Test successful Advise call."""
        pUnk = mock.Mock()
        mock_sink = mock.Mock()
        pUnk.QueryInterface.return_value = mock_sink
        pdwCookie = pointer(DWORD(0))
        
        result = self.conn_point.IConnectionPoint_Advise(None, pUnk, pdwCookie)
        
        self.assertEqual(result, S_OK)
        self.assertEqual(pdwCookie[0], 1)
        self.assertEqual(self.conn_point._cookie, 1)
        self.assertEqual(self.conn_point._connections[1], mock_sink)
        pUnk.QueryInterface.assert_called_once_with(self.sink_interface)

    def test_advise_null_punk(self):
        """Test Advise with null pUnk returns E_POINTER."""
        pdwCookie = pointer(DWORD(0))
        result = self.conn_point.IConnectionPoint_Advise(None, None, pdwCookie)
        self.assertEqual(result, E_POINTER)

    def test_advise_null_pdwcookie(self):
        """Test Advise with null pdwCookie returns E_POINTER."""
        pUnk = mock.Mock()
        result = self.conn_point.IConnectionPoint_Advise(None, pUnk, None)
        self.assertEqual(result, E_POINTER)

    def test_advise_comerror(self):
        """Test Advise when QueryInterface raises COMError returns CONNECT_E_CANNOTCONNECT."""
        pUnk = mock.Mock()
        pUnk.QueryInterface.side_effect = COMError(E_NOINTERFACE, "No interface")
        pdwCookie = pointer(DWORD(0))
        
        result = self.conn_point.IConnectionPoint_Advise(None, pUnk, pdwCookie)
        
        self.assertEqual(result, CONNECT_E_CANNOTCONNECT)
        self.assertEqual(self.conn_point._connections, {})

    def test_unadvise_success(self):
        """Test successful Unadvise call."""
        mock_sink = mock.Mock()
        self.conn_point._connections[5] = mock_sink
        self.conn_point._cookie = 5
        
        result = self.conn_point.IConnectionPoint_Unadvise(None, 5)
        
        self.assertEqual(result, S_OK)
        self.assertNotIn(5, self.conn_point._connections)

    def test_unadvise_not_found(self):
        """Test Unadvise with non-existent cookie returns CONNECT_E_NOCONNECTION."""
        result = self.conn_point.IConnectionPoint_Unadvise(None, 999)
        self.assertEqual(result, CONNECT_E_NOCONNECTION)

    def test_get_connection_point_container(self):
        """Test GetConnectionPointContainer returns E_NOTIMPL."""
        ppCPC = c_void_p()
        result = self.conn_point.IConnectionPoint_GetConnectionPointContainer(None, ppCPC)
        self.assertEqual(result, E_NOTIMPL)

    def test_get_connection_interface(self):
        """Test GetConnectionInterface returns E_NOTIMPL."""
        pIID = pointer(GUID())
        result = self.conn_point.IConnectionPoint_GetConnectionInterface(None, pIID)
        self.assertEqual(result, E_NOTIMPL)

    def test_call_sinks_idispatch(self):
        """Test _call_sinks with IDispatch-derived interface (uses Invoke)."""
        # Create a sink interface with Invoke attribute
        sink_with_invoke = type('SinkWithInvoke', (), {'Invoke': mock.Mock()})
        conn_point = ConnectionPointImpl(sink_with_invoke, self.sink_typeinfo)
        
        mock_sink = mock.Mock()
        mock_sink.Invoke = mock.Mock(return_value=42)
        conn_point._connections[1] = mock_sink
        conn_point._cookie = 1
        
        results = conn_point._call_sinks("TestEvent", "arg1", kwarg1="value1")
        
        self.assertEqual(results, [42])
        mock_sink.Invoke.assert_called_once()
        # Check that functools.partial was used with dispid
        call_args = mock_sink.Invoke.call_args
        self.assertEqual(call_args[0][0], 123)  # dispid

    def test_call_sinks_non_idispatch(self):
        """Test _call_sinks with non-IDispatch interface (uses getattr)."""
        sink_no_invoke = type('SinkNoInvoke', (), {})
        conn_point = ConnectionPointImpl(sink_no_invoke, self.sink_typeinfo)
        
        mock_sink = mock.Mock()
        mock_sink.TestEvent = mock.Mock(return_value=99)
        conn_point._connections[1] = mock_sink
        conn_point._cookie = 1
        
        results = conn_point._call_sinks("TestEvent", "arg1", kwarg1="value1")
        
        self.assertEqual(results, [99])
        mock_sink.TestEvent.assert_called_once_with("arg1", kwarg1="value1")

    def test_call_sink_success(self):
        """Test _call_sink yields result on success."""
        mth = mock.Mock(return_value="success")
        results = list(self.conn_point._call_sink("TestEvent", 1, mth, "arg1"))
        self.assertEqual(results, ["success"])

    def test_call_sink_rpc_server_unavailable(self):
        """Test _call_sink handles RPC_S_SERVER_UNAVAILABLE by removing connection."""
        mth = mock.Mock(side_effect=COMError(RPC_S_SERVER_UNAVAILABLE, "Server unavailable"))
        self.conn_point._connections[1] = mock.Mock()
        
        with mock.patch('comtypes.server.connectionpoints.logger') as mock_logger:
            results = list(self.conn_point._call_sink("TestEvent", 1, mth))
            
            self.assertEqual(results, [])
            self.assertNotIn(1, self.conn_point._connections)
            mock_logger.warning.assert_called_once()

    def test_call_sink_rpc_server_unavailable_keyerror(self):
        """Test _call_sink handles RPC_S_SERVER_UNAVAILABLE when key already gone."""
        mth = mock.Mock(side_effect=COMError(RPC_S_SERVER_UNAVAILABLE, "Server unavailable"))
        # Connection already removed
        self.conn_point._connections = {}
        
        with mock.patch('comtypes.server.connectionpoints.logger') as mock_logger:
            results = list(self.conn_point._call_sink("TestEvent", 1, mth))
            
            self.assertEqual(results, [])
            mock_logger.warning.assert_called_once()

    def test_call_sink_other_comerror(self):
        """Test _call_sink logs warning for other COMErrors but doesn't remove connection."""
        mth = mock.Mock(side_effect=COMError(E_FAIL, "Other error"))
        self.conn_point._connections[1] = mock.Mock()
        
        with mock.patch('comtypes.server.connectionpoints.logger') as mock_logger:
            results = list(self.conn_point._call_sink("TestEvent", 1, mth))
            
            self.assertEqual(results, [])
            self.assertIn(1, self.conn_point._connections)
            mock_logger.warning.assert_called_once()


class MockOutgoingInterface:
    """Mock outgoing interface for ConnectableObjectMixin tests."""
    _iid_ = GUID.create_new()
    
    def __init__(self):
        pass


class TestConnectableObjectMixin(unittest.TestCase):
    def setUp(self):
        # Create a test class that uses the mixin
        class TestConnectableObject(ConnectableObjectMixin, COMObject):
            _outgoing_interfaces_ = [MockOutgoingInterface]
            _reg_typelib_ = ("{00000000-0000-0000-0000-000000000000}", 1, 0)
            _com_interfaces_ = [IUnknown]
        
        self.TestClass = TestConnectableObject

    @mock.patch('comtypes.server.connectionpoints.LoadRegTypeLib')
    def test_init(self, mock_load_typelib):
        """Test ConnectableObjectMixin initialization."""
        mock_tlib = mock.Mock()
        mock_typeinfo = mock.Mock()
        mock_tlib.GetTypeInfoOfGuid.return_value = mock_typeinfo
        mock_load_typelib.return_value = mock_tlib
        
        obj = self.TestClass()
        
        self.assertIn(MockOutgoingInterface, obj._ConnectableObjectMixin__connections)
        conn_point = obj._ConnectableObjectMixin__connections[MockOutgoingInterface]
        self.assertIsInstance(conn_point, ConnectionPointImpl)
        self.assertEqual(conn_point._sink_interface, MockOutgoingInterface)
        self.assertEqual(conn_point._typeinfo, mock_typeinfo)
        mock_load_typelib.assert_called_once()
        mock_tlib.GetTypeInfoOfGuid.assert_called_once_with(MockOutgoingInterface._iid_)

    def test_enum_connection_points(self):
        """Test EnumConnectionPoints returns E_NOTIMPL."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        obj._ConnectableObjectMixin__connections = {}
        
        ppEnum = c_void_p()
        result = obj.IConnectionPointContainer_EnumConnectionPoints(None, ppEnum)
        self.assertEqual(result, E_NOTIMPL)

    def test_find_connection_point_success(self):
        """Test FindConnectionPoint finds matching interface."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        
        mock_conn = mock.Mock()
        mock_conn.IUnknown_QueryInterface.return_value = S_OK
        obj._ConnectableObjectMixin__connections = {MockOutgoingInterface: mock_conn}
        
        refiid = pointer(GUID(MockOutgoingInterface._iid_))
        ppcp = c_void_p()
        
        with mock.patch('comtypes.server.connectionpoints.logger'):
            result = obj.IConnectionPointContainer_FindConnectionPoint(None, refiid, ppcp)
        
        self.assertEqual(result, S_OK)
        mock_conn.IUnknown_QueryInterface.assert_called_once()

    def test_find_connection_point_null_ppcp(self):
        """Test FindConnectionPoint with null ppcp returns E_POINTER."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        obj._ConnectableObjectMixin__connections = {}
        
        refiid = pointer(GUID(MockOutgoingInterface._iid_))
        result = obj.IConnectionPointContainer_FindConnectionPoint(None, refiid, None)
        self.assertEqual(result, E_POINTER)

    def test_find_connection_point_not_found(self):
        """Test FindConnectionPoint with non-matching IID returns CONNECT_E_NOCONNECTION."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        obj._ConnectableObjectMixin__connections = {}
        
        other_iid = GUID.create_new()
        refiid = pointer(GUID(other_iid))
        ppcp = c_void_p()
        
        with mock.patch('comtypes.server.connectionpoints.logger') as mock_logger:
            result = obj.IConnectionPointContainer_FindConnectionPoint(None, refiid, ppcp)
        
        self.assertEqual(result, CONNECT_E_NOCONNECTION)
        mock_logger.debug.assert_any_call("No connectionpoint found")

    def test_fire_event_with_index(self):
        """Test Fire_Event with integer interface index."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        
        mock_conn = mock.Mock()
        mock_conn._call_sinks.return_value = ["result1", "result2"]
        obj._ConnectableObjectMixin__connections = {MockOutgoingInterface: mock_conn}
        
        with mock.patch('comtypes.server.connectionpoints.logger'):
            results = obj.Fire_Event(0, "TestEvent", "arg1", kwarg1="value1")
        
        self.assertEqual(results, ["result1", "result2"])
        mock_conn._call_sinks.assert_called_once_with("TestEvent", "arg1", kwarg1="value1")

    def test_fire_event_with_interface(self):
        """Test Fire_Event with interface type."""
        obj = self.TestClass.__new__(self.TestClass)
        obj._outgoing_interfaces_ = [MockOutgoingInterface]
        
        mock_conn = mock.Mock()
        mock_conn._call_sinks.return_value = ["result"]
        obj._ConnectableObjectMixin__connections = {MockOutgoingInterface: mock_conn}
        
        with mock.patch('comtypes.server.connectionpoints.logger'):
            results = obj.Fire_Event(MockOutgoingInterface, "TestEvent", "arg1")
        
        self.assertEqual(results, ["result"])
        mock_conn._call_sinks.assert_called_once_with("TestEvent", "arg1")


if __name__ == "__main__":
    unittest.main()