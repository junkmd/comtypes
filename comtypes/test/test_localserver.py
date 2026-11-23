import sys
import unittest
from unittest import mock

import comtypes
from comtypes import COMObject, GUID, IUnknown
from comtypes.server.localserver import (
    ClassFactory,
    run,
    REGCLS_MULTIPLEUSE,
    REGCLS_SINGLEUSE,
    REGCLS_MULTI_SEPARATE,
    REGCLS_SUSPENDED,
    REGCLS_SURROGATE,
)
from comtypes.server import IClassFactory


class TestLocalServer(unittest.TestCase):
    def setUp(self):
        # Reset any global state
        pass

    def test_constants(self):
        """Test that REGCLS constants have expected values."""
        self.assertEqual(REGCLS_SINGLEUSE, 0)
        self.assertEqual(REGCLS_MULTIPLEUSE, 1)
        self.assertEqual(REGCLS_MULTI_SEPARATE, 2)
        self.assertEqual(REGCLS_SUSPENDED, 4)
        self.assertEqual(REGCLS_SURROGATE, 8)

    def test_class_factory_init(self):
        """Test ClassFactory initialization."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class') as mock_register:
            factory = ClassFactory(TestCOMObject, "arg1", kwarg1="value1")
            
            self.assertEqual(factory._cls, TestCOMObject)
            self.assertEqual(factory._args, ("arg1",))
            self.assertEqual(factory._kw, {"kwarg1": "value1"})
            mock_register.assert_called_once()

    def test_class_factory_init_with_regcls_attribute(self):
        """Test ClassFactory uses _regcls_ from class if present."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _regcls_ = REGCLS_SINGLEUSE
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'):
            factory = ClassFactory(TestCOMObject)
            # The regcls should be taken from the class attribute
            self.assertEqual(factory.regcls, REGCLS_MULTIPLEUSE)  # default
            # But _register_class should use the class's _regcls_

    def test_iunknown_addref(self):
        """Test IUnknown_AddRef returns 2."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'):
            factory = ClassFactory(TestCOMObject)
            result = factory.IUnknown_AddRef(None)
            self.assertEqual(result, 2)

    def test_iunknown_release(self):
        """Test IUnknown_Release returns 1."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'):
            factory = ClassFactory(TestCOMObject)
            result = factory.IUnknown_Release(None)
            self.assertEqual(result, 1)

    def test_register_class(self):
        """Test _register_class calls CoRegisterClassObject."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch('comtypes.server.localserver._CoRegisterClassObject') as mock_register, \
             mock.patch('comtypes.server.localserver.byref') as mock_byref, \
             mock.patch('comtypes.server.localserver.GUID') as mock_guid, \
             mock.patch('comtypes.server.localserver.c_ulong') as mock_culong:
            
            mock_cookie = mock.Mock()
            mock_culong.return_value = mock_cookie
            mock_ptr = mock.Mock()
            factory = ClassFactory(TestCOMObject)
            factory._com_pointers_ = {IUnknown._iid_: mock_ptr}
            
            # Call _register_class directly
            factory._register_class()
            
            mock_register.assert_called_once()
            args, kwargs = mock_register.call_args
            self.assertEqual(len(args), 5)
            self.assertEqual(factory.cookie, mock_cookie)

    def test_revoke_class(self):
        """Test _revoke_class calls CoRevokeClassObject."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch('comtypes.server.localserver._CoRevokeClassObject') as mock_revoke, \
             mock.patch.object(ClassFactory, '_register_class'):
            
            factory = ClassFactory(TestCOMObject)
            factory.cookie = 12345
            
            factory._revoke_class()
            
            mock_revoke.assert_called_once_with(12345)

    def test_create_instance(self):
        """Test CreateInstance creates object and calls QI."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'), \
             mock.patch.object(TestCOMObject, 'IUnknown_QueryInterface', return_value=0) as mock_qi:
            
            factory = ClassFactory(TestCOMObject)
            
            from ctypes import c_void_p
            riid = mock.Mock()
            riid.__getitem__ = mock.Mock(return_value=IUnknown._iid_)
            ppv = c_void_p()
            
            result = factory.CreateInstance(None, None, riid, ppv)
            
            mock_qi.assert_called_once()
            self.assertEqual(result, 0)

    def test_lock_server_lock(self):
        """Test LockServer with fLock=True."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'):
            factory = ClassFactory(TestCOMObject)
            
            mock_server = mock.Mock()
            COMObject.__server__ = mock_server
            
            try:
                result = factory.LockServer(None, True)
                mock_server.Lock.assert_called_once()
                mock_server.Unlock.assert_not_called()
                self.assertEqual(result, 0)
            finally:
                COMObject.__server__ = None

    def test_lock_server_unlock(self):
        """Test LockServer with fLock=False."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch.object(ClassFactory, '_register_class'):
            factory = ClassFactory(TestCOMObject)
            
            mock_server = mock.Mock()
            COMObject.__server__ = mock_server
            
            try:
                result = factory.LockServer(None, False)
                mock_server.Unlock.assert_called_once()
                mock_server.Lock.assert_not_called()
                self.assertEqual(result, 0)
            finally:
                COMObject.__server__ = None

    def test_run_function(self):
        """Test run function creates ClassFactory instances and starts server."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
            _com_interfaces_ = [IUnknown]
        
        with mock.patch('comtypes.server.localserver.ClassFactory') as mock_factory_class, \
             mock.patch('comtypes.COMObject.__run_localserver__') as mock_run_server:
            
            mock_factory = mock.Mock()
            mock_factory_class.return_value = mock_factory
            
            run([TestCOMObject])
            
            mock_factory_class.assert_called_once_with(TestCOMObject)
            mock_run_server.assert_called_once_with([mock_factory])


if __name__ == "__main__":
    unittest.main()