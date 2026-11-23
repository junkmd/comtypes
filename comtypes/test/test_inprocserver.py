import sys
import unittest
from unittest import mock

import comtypes
from comtypes import COMObject, GUID, IUnknown
from comtypes.server.inprocserver import (
    ClassFactory,
    DllCanUnloadNow,
    DllGetClassObject,
    inproc_find_class,
    _clsid_to_class,
    _logging_configured,
    _setup_logging,
)
from comtypes.server import IClassFactory


class TestInprocServer(unittest.TestCase):
    def setUp(self):
        # Clear the class registry
        _clsid_to_class.clear()
        # Reset logging configured flag
        import comtypes.server.inprocserver as inproc_mod
        inproc_mod._logging_configured = False

    def test_class_factory_create_instance(self):
        """Test ClassFactory.CreateInstance creates an instance and calls QI."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        factory = ClassFactory(TestCOMObject)
        
        # Mock the QueryInterface call
        with mock.patch.object(TestCOMObject, 'IUnknown_QueryInterface', return_value=0) as mock_qi:
            from ctypes import c_void_p, pointer
            riid = pointer(GUID(IUnknown._iid_))
            ppv = c_void_p()
            
            result = factory.IClassFactory_CreateInstance(None, None, riid, ppv)
            
            mock_qi.assert_called_once()
            self.assertEqual(result, 0)

    def test_class_factory_lock_server_lock(self):
        """Test ClassFactory.LockServer with fLock=True."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        factory = ClassFactory(TestCOMObject)
        
        # Mock the server
        mock_server = mock.Mock()
        COMObject.__server__ = mock_server
        
        try:
            result = factory.IClassFactory_LockServer(None, True)
            mock_server.Lock.assert_called_once()
            mock_server.Unlock.assert_not_called()
            self.assertEqual(result, 0)
        finally:
            COMObject.__server__ = None

    def test_class_factory_lock_server_unlock(self):
        """Test ClassFactory.LockServer with fLock=False."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        factory = ClassFactory(TestCOMObject)
        
        # Mock the server
        mock_server = mock.Mock()
        COMObject.__server__ = mock_server
        
        try:
            result = factory.IClassFactory_LockServer(None, False)
            mock_server.Unlock.assert_called_once()
            mock_server.Lock.assert_not_called()
            self.assertEqual(result, 0)
        finally:
            COMObject.__server__ = None

    def test_inproc_find_class_from_cache(self):
        """Test inproc_find_class returns class from _clsid_to_class cache."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        _clsid_to_class[clsid] = TestCOMObject
        
        result = inproc_find_class(clsid)
        self.assertEqual(result, TestCOMObject)

    def test_inproc_find_class_from_registry(self):
        """Test inproc_find_class reads from registry when not in cache."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        # Mock winreg
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('sys.path', new=list(sys.path)) as mock_path, \
             mock.patch('builtins.__import__') as mock_import, \
             mock.patch('sys.modules', new=dict(sys.modules)) as mock_modules:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                (r"C:\test\path", 0),  # PythonPath
                ("test_module.TestCOMObject", 0),  # PythonClass
            ]
            
            mock_mod = mock.Mock()
            mock_mod.TestCOMObject = TestCOMObject
            mock_modules["test_module"] = mock_mod
            mock_import.return_value = None
            
            result = inproc_find_class(clsid)
            
            self.assertEqual(result, TestCOMObject)
            mock_open_key.assert_called_once()
            self.assertEqual(mock_query_value.call_count, 2)

    def test_inproc_find_class_registry_no_python_path(self):
        """Test inproc_find_class when PythonPath is not in registry."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('builtins.__import__') as mock_import, \
             mock.patch('sys.modules', new=dict(sys.modules)) as mock_modules:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            # First call raises OSError (no PythonPath), second returns PythonClass
            mock_query_value.side_effect = [
                OSError("No PythonPath"),
                ("test_module.TestCOMObject", 0),
            ]
            
            mock_mod = mock.Mock()
            mock_mod.TestCOMObject = TestCOMObject
            mock_modules["test_module"] = mock_mod
            mock_import.return_value = None
            
            result = inproc_find_class(clsid)
            
            self.assertEqual(result, TestCOMObject)

    def test_inproc_find_class_path_already_in_sys_path(self):
        """Test inproc_find_class when path is already in sys.path."""
        clsid = GUID.create_new()
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        test_path = r"C:\already\in\path"
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('sys.path', [test_path]), \
             mock.patch('builtins.__import__') as mock_import, \
             mock.patch('sys.modules', new=dict(sys.modules)) as mock_modules:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                (test_path, 0),  # PythonPath
                ("test_module.TestCOMObject", 0),  # PythonClass
            ]
            
            mock_mod = mock.Mock()
            mock_mod.TestCOMObject = TestCOMObject
            mock_modules["test_module"] = mock_mod
            mock_import.return_value = None
            
            result = inproc_find_class(clsid)
            
            self.assertEqual(result, TestCOMObject)

    def test_setup_logging_already_configured(self):
        """Test _setup_logging returns early if already configured."""
        import comtypes.server.inprocserver as inproc_mod
        inproc_mod._logging_configured = True
        
        clsid = GUID.create_new()
        _setup_logging(clsid)  # Should not raise
        
        inproc_mod._logging_configured = False

    def test_setup_logging_no_logging_key(self):
        """Test _setup_logging when registry key doesn't exist."""
        clsid = GUID.create_new()
        
        with mock.patch('winreg.OpenKey', side_effect=OSError("Key not found")):
            _setup_logging(clsid)  # Should not raise

    def test_setup_logging_with_format_and_levels(self):
        """Test _setup_logging with format and levels from registry."""
        clsid = GUID.create_new()
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('comtypes.logutil.NTDebugHandler') as mock_handler_class, \
             mock.patch('logging.Formatter') as mock_formatter_class, \
             mock.patch('logging.root') as mock_root:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                ("%(message)s", 0),  # format
                (["comtypes=DEBUG"], 0),  # levels
            ]
            
            mock_handler = mock.Mock()
            mock_handler_class.return_value = mock_handler
            mock_formatter = mock.Mock()
            mock_formatter_class.return_value = mock_formatter
            
            _setup_logging(clsid)
            
            mock_handler.setFormatter.assert_called_once_with(mock_formatter)
            mock_root.addHandler.assert_called_once_with(mock_handler)
            mock_formatter_class.assert_called_once_with("%(message)s")

    def test_setup_logging_without_format(self):
        """Test _setup_logging when format is not in registry."""
        clsid = GUID.create_new()
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('comtypes.logutil.NTDebugHandler') as mock_handler_class, \
             mock.patch('logging.Formatter') as mock_formatter_class, \
             mock.patch('logging.root') as mock_root:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                OSError("No format"),  # format
                (["comtypes=DEBUG"], 0),  # levels
            ]
            
            mock_handler = mock.Mock()
            mock_handler_class.return_value = mock_handler
            mock_formatter = mock.Mock()
            mock_formatter_class.return_value = mock_formatter
            
            _setup_logging(clsid)
            
            # Should use default format
            mock_formatter_class.assert_called_once_with("(Thread %(thread)s):%(levelname)s:%(message)s")

    def test_setup_logging_levels_reg_sz(self):
        """Test _setup_logging with REG_SZ levels value."""
        clsid = GUID.create_new()
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('comtypes.logutil.NTDebugHandler') as mock_handler_class, \
             mock.patch('logging.Formatter'), \
             mock.patch('logging.root'), \
             mock.patch('logging.getLogger') as mock_get_logger:
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                ("%(message)s", 0),  # format
                ("comtypes=DEBUG", 0),  # levels as REG_SZ
            ]
            
            mock_handler = mock.Mock()
            mock_handler_class.return_value = mock_handler
            
            mock_logger = mock.Mock()
            mock_get_logger.return_value = mock_logger
            
            _setup_logging(clsid)
            
            mock_get_logger.assert_called_once_with("comtypes")
            mock_logger.setLevel.assert_called_once()

    def test_setup_logging_levels_invalid_type(self):
        """Test _setup_logging with invalid levels registry type."""
        clsid = GUID.create_new()
        
        with mock.patch('winreg.OpenKey') as mock_open_key, \
             mock.patch('winreg.QueryValueEx') as mock_query_value, \
             mock.patch('comtypes.logutil.NTDebugHandler'), \
             mock.patch('logging.Formatter'), \
             mock.patch('logging.root'):
            
            mock_key = mock.Mock()
            mock_open_key.return_value = mock_key
            mock_query_value.side_effect = [
                ("%(message)s", 0),  # format
                (["comtypes=DEBUG"], 123),  # invalid type
            ]
            
            _setup_logging(clsid)  # Should not raise, just return early

    def test_dll_get_class_object_class_not_available(self):
        """Test DllGetClassObject returns CLASS_E_CLASSNOTAVAILABLE."""
        clsid = GUID.create_new()
        iid = IUnknown._iid_
        
        with mock.patch('comtypes.server.inprocserver.inproc_find_class', return_value=None), \
             mock.patch('comtypes.server.inprocserver._setup_logging'):
            
            result = DllGetClassObject(
                int.from_bytes(clsid._bytes_le, 'little'),
                int.from_bytes(iid._bytes_le, 'little'),
                0
            )
            
            from comtypes import hresult
            self.assertEqual(result, hresult.CLASS_E_CLASSNOTAVAILABLE)

    def test_dll_get_class_object_exception(self):
        """Test DllGetClassObject returns E_FAIL on exception."""
        clsid = GUID.create_new()
        iid = IUnknown._iid_
        
        class TestCOMObject(COMObject):
            _reg_clsid_ = clsid
            _com_interfaces_ = [IUnknown]
        
        with mock.patch('comtypes.server.inprocserver.inproc_find_class', return_value=TestCOMObject), \
             mock.patch('comtypes.server.inprocserver._setup_logging'), \
             mock.patch('comtypes.server.inprocserver.ClassFactory') as mock_factory_class, \
             mock.patch('comtypes.server.inprocserver._critical') as mock_critical:
            
            mock_factory = mock.Mock()
            mock_factory.IUnknown_QueryInterface.side_effect = Exception("Test error")
            mock_factory_class.return_value = mock_factory
            
            result = DllGetClassObject(
                int.from_bytes(clsid._bytes_le, 'little'),
                int.from_bytes(iid._bytes_le, 'little'),
                0
            )
            
            from comtypes import hresult
            self.assertEqual(result, hresult.E_FAIL)
            mock_critical.assert_called_once()

    def test_dll_can_unload_now(self):
        """Test DllCanUnloadNow returns S_FALSE."""
        # Mock the server
        mock_server = mock.Mock()
        mock_server.DllCanUnloadNow.return_value = 0
        COMObject.__server__ = mock_server
        
        try:
            with mock.patch('comtypes.COMObject.__run_inprocserver__'):
                result = DllCanUnloadNow()
                from comtypes import hresult
                self.assertEqual(result, hresult.S_FALSE)
        finally:
            COMObject.__server__ = None


if __name__ == "__main__":
    unittest.main()