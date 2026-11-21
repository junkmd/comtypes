import doctest
import sys
import unittest
from ctypes import GetLastError, pointer
from typing import Any

import comtypes.server
import comtypes.test.TestComServer
from comtypes import BSTR, errorinfo, hresult, shelllink
from comtypes.automation import VARIANT, _midlSAFEARRAY
from comtypes.client import CreateObject, GetActiveObject
from comtypes.server.register import register, unregister
from comtypes.test.find_memleak import find_memleak

try:
    from win32com.client import Dispatch

    IMPORT_PYWIN32_FAILED = False
except ImportError:
    IMPORT_PYWIN32_FAILED = True


def setUpModule():
    try:
        register(comtypes.test.TestComServer.TestComServer)
    except OSError as e:
        if e.winerror != 5:  # [Error 5] Access is denied
            raise e
        raise unittest.SkipTest(
            "This test requires the tests to be run as admin since it tries to "
            "register the test COM server."
        )


def tearDownModule():
    unregister(comtypes.test.TestComServer.TestComServer)


class BaseServerTest:
    def create_object(self) -> Any: ...

    def _find_memleak(self, func):
        bytes = find_memleak(func, target="private_usage")
        self.assertFalse(bytes, f"Leaks {bytes} bytes")  # type: ignore

    def test_mixedinout(self):
        o = self.create_object()
        self.assertEqual(o.MixedInOut(2, 4), (3, 5))  # type: ignore

    def test_getname(self):
        # This tests a tricky bug, introduced with this patch:
        # http://www.python.org/sf/1643874
        #
        # Returning a BSTR as an [out] parameter from a server
        # implementation must transfer the ownership to the caller.
        # When this is not done, the BSTR instance is SysFreeString'd
        # too early, and the memory is reused.
        obj = self.create_object()
        pb = pointer(BSTR())
        # Get the BSTR from the server:
        obj._ITestComServer__com__get_name(pb)
        # Retrieve the value, but keep the pointer to the BSTR alive:
        name = pb[0]
        # Create sme BSTR's to reuse the memory in case it has been freed:
        for i in range(10):
            BSTR("f" * len(name))
        # Make sure the pointer is still valid:
        self.assertEqual(pb[0], name)  # type: ignore

    def test_get_id(self):
        obj = self.create_object()
        self._find_memleak(lambda: obj.id)

    def test_get_name(self):
        obj = self.create_object()
        self._find_memleak(lambda: obj.name)

    def test_set_name(self):
        obj = self.create_object()

        def func():
            obj.name = "abcde"

        self._find_memleak(func)

    def test_SetName(self):
        obj = self.create_object()

        def func():
            obj.SetName("abcde")

        self._find_memleak(func)

    def test_eval(self):
        obj = self.create_object()

        def func():
            return obj.eval("(1, 2, 3)")

        self.assertEqual(func(), (1, 2, 3))  # type: ignore
        self._find_memleak(func)

    def test_get_typeinfo(self):
        obj = self.create_object()

        def func():
            obj.GetTypeInfo(0)
            obj.GetTypeInfoCount()
            obj.QueryInterface(comtypes.IUnknown)

        self._find_memleak(func)


class TestInproc(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return CreateObject(
            "TestComServerLib.TestComServer", clsctx=comtypes.CLSCTX_INPROC_SERVER
        )

    def test_eval(self):
        obj = self.create_object()

        def func():
            return obj.eval("(1, 2, 3)")

        self.assertEqual(func(), (1, 2, 3))  # type: ignore
        self._find_memleak(func)


import ctypes
from contextlib import contextmanager
from ctypes import wintypes

EXCEPTION_CONTINUE_SEARCH = 0
STATUS_ACCESS_VIOLATION = 0xC0000005


class EXCEPTION_RECORD(ctypes.Structure):
    pass


EXCEPTION_RECORD._fields_ = [
    ("ExceptionCode", wintypes.DWORD),
    ("ExceptionFlags", wintypes.DWORD),
    ("ExceptionRecord", ctypes.POINTER(EXCEPTION_RECORD)),
    ("ExceptionAddress", ctypes.c_void_p),
    ("NumberParameters", wintypes.DWORD),
    ("ExceptionInformation", ctypes.c_size_t * 15),
]


class EXCEPTION_POINTERS(ctypes.Structure):
    _fields_ = [
        ("ExceptionRecord", ctypes.POINTER(EXCEPTION_RECORD)),
        ("ContextRecord", ctypes.c_void_p),
    ]


HANDLER = ctypes.WINFUNCTYPE(wintypes.LONG, ctypes.POINTER(EXCEPTION_POINTERS))

kernel32 = ctypes.WinDLL("kernel32")
kernel32.AddVectoredExceptionHandler.restype = ctypes.c_void_p
kernel32.AddVectoredExceptionHandler.argtypes = [wintypes.ULONG, HANDLER]
kernel32.RemoveVectoredExceptionHandler.argtypes = [ctypes.c_void_p]
kernel32.RemoveVectoredExceptionHandler.restype = wintypes.ULONG


@contextmanager
def record_seh_codes():
    codes = []

    @HANDLER
    def handler(ptrs):
        codes.append(ptrs.contents.ExceptionRecord.contents.ExceptionCode)
        return EXCEPTION_CONTINUE_SEARCH  # 処理は ctypes 本来の __try/__except に任せる

    h = kernel32.AddVectoredExceptionHandler(1, handler)
    try:
        yield codes
    finally:
        kernel32.RemoveVectoredExceptionHandler(h)


STATUS_ACCESS_VIOLATION = 0xC0000005


class TestLocalServer(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return CreateObject(
            "TestComServerLib.TestComServer", clsctx=comtypes.CLSCTX_LOCAL_SERVER
        )

    def test_get_typeinfo(self):
        # `IDispatch::GetTypeInfo` on a LOCAL_SERVER returns an `ITypeInfo` that
        # has to be marshaled across the process boundary.  The `ITypeInfo`
        # marshaler (in oleaut32) has to resolve the type library on the client
        # side, and in this test setup that step faults: the client dereferences
        # a NULL pointer while the COM call is being completed.
        #
        #     OSError: exception: access violation reading 0x0000000000000000
        #
        # `comtypes.automation.IDispatch.GetTypeInfo` catches the resulting
        # `OSError` (a Windows SEH access violation) and re-raises it as
        # `comtypes.AccessViolationError`, so the failure is observable from
        # Python instead of crashing the process.
        #
        # This is a characterization test: it records the current behaviour,
        # which is not necessarily the desired one.  If the marshaling problem
        # is ever fixed, this test will start failing and should be updated.
        obj = self.create_object()
        with record_seh_codes() as o:
            with self.assertRaises(OSError) as cm:
                obj.GetTypeInfo(0)
            print(o[-1] & 0xFFFFFFFF == STATUS_ACCESS_VIOLATION)
        print(cm.exception.winerror)

        print(GetLastError())


# class TestGetTypeInfoNull(unittest.TestCase):
#     def test_null_typeinfo(self):
#         # `IDispatch::GetTypeInfo` may return `S_OK` with a NULL `ppTypeInfo`
#         # (for example when the server has no type information).  Calling
#         # `QueryInterface` on the NULL pointer would raise an access violation,
#         # so `comtypes.automation.IDispatch.GetTypeInfo` guards against it and
#         # raises a `COMError` instead.
#         #
#         # `TestComServer` never returns NULL, so this test drives the guard
#         # directly by faking the low-level `_GetTypeInfo` call.  `GetTypeInfo`
#         # is a plain Python method, so it can be called on any object that
#         # provides `_GetTypeInfo`.
#         class FakeDispatch:
#             def _GetTypeInfo(self, index, lcid):
#                 return None  # NULL `ppTypeInfo`

#         with self.assertRaises(comtypes.COMError):
#             IDispatch.GetTypeInfo(FakeDispatch(), 0)


@unittest.skipIf(IMPORT_PYWIN32_FAILED, "This depends on 'pywin32'.")
class TestInproc_win32com(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return Dispatch("TestComServerLib.TestComServer")

    if sys.version_info >= (3, 14) or sys.version_info[:2] == (3, 12):

        def test_eval(self):
            obj = self.create_object()

            def func():
                return obj.eval("(1, 2, 3)")

            self.assertEqual(func(), (1, 2, 3))  # type: ignore
            self._find_memleak(func)

    # These tests make no sense with win32com, override to disable them:
    @unittest.skip("This test make no sense with win32com.")
    def test_get_typeinfo(self):
        pass

    @unittest.skip("This test make no sense with win32com.")
    def test_getname(self):
        pass

    @unittest.skip("This test make no sense with win32com.")
    def test_mixedinout(self):
        # Not sure about this; it raise 'Invalid Number of parameters'
        # Is mixed [in], [out] args not compatible with IDispatch???
        pass


@unittest.skipIf(IMPORT_PYWIN32_FAILED, "This depends on 'pywin32'.")
class TestLocalServer_win32com(BaseServerTest, unittest.TestCase):
    def create_object(self):
        return Dispatch(
            "TestComServerLib.TestComServer", clsctx=comtypes.CLSCTX_LOCAL_SERVER
        )

    # These tests are skipped for the same reason as `TestInproc_win32com`.
    @unittest.skip("This test make no sense with win32com.")
    def test_get_typeinfo(self):
        pass

    @unittest.skip("This test make no sense with win32com.")
    def test_getname(self):
        pass

    @unittest.skip("This test make no sense with win32com.")
    def test_mixedinout(self):
        pass


class ActiveObjTest(unittest.TestCase):
    def get_active_object(self) -> comtypes.test.TestComServer.ITestComServer:
        return GetActiveObject(
            comtypes.test.TestComServer.TestComServer,
            interface=comtypes.test.TestComServer.ITestComServer,
        )

    def _doit(self, weak: bool):
        comobj = comtypes.test.TestComServer.TestComServer()
        with self.assertRaises(OSError):
            self.get_active_object()
        handle = comtypes.server.RegisterActiveObject(comobj, weak=weak)
        activeobj = self.get_active_object()
        self.assertEqual(activeobj.MixedInOut(2, 4), (3, 5))
        self.assertEqual(activeobj.AddRef(), 3)
        self.assertEqual(activeobj.Release(), 2)
        punk = comobj.QueryInterface(comtypes.test.TestComServer.ITestComServer)
        self.assertEqual(activeobj, punk)
        comtypes.server.RevokeActiveObject(handle)
        with self.assertRaises(OSError):
            self.get_active_object()

    def test_activeobj_weak_registration(self):
        self._doit(weak=True)

    def test_activeobj_strong_registration(self):
        self._doit(weak=False)


class SupportErrorInfoTest(unittest.TestCase):
    def create_object(self):
        return CreateObject(
            comtypes.test.TestComServer.TestComServer,
            interface=comtypes.test.TestComServer.ITestComServer,
        )

    def test_returns_S_OK(self):
        psei = self.create_object().QueryInterface(errorinfo.ISupportErrorInfo)
        iid = comtypes.test.TestComServer.ITestComServer._iid_
        hr = psei.InterfaceSupportsErrorInfo(iid)
        self.assertEqual(hr, hresult.S_OK)

    def test_returns_S_FALSE(self):
        psei = self.create_object().QueryInterface(errorinfo.ISupportErrorInfo)
        iid = shelllink.IShellLinkW._iid_
        hr = psei.InterfaceSupportsErrorInfo(iid)
        self.assertEqual(hr, hresult.S_FALSE)


class VariantTest(unittest.TestCase):
    def test_UDT(self):
        from comtypes.gen.TestComServerLib import MYCOLOR

        v = VARIANT(MYCOLOR(red=1.0, green=2.0, blue=3.0))
        value = v.value
        self.assertEqual(1.0, value.red)  # type: ignore
        self.assertEqual(2.0, value.green)  # type: ignore
        self.assertEqual(3.0, value.blue)  # type: ignore

        def func():
            v = VARIANT(MYCOLOR(red=1.0, green=2.0, blue=3.0))
            return v.value

        bytes = find_memleak(func, target="private_usage")
        self.assertFalse(bytes, f"Leaks {bytes} bytes")  # type: ignore


class SafeArrayTest(unittest.TestCase):
    def test_UDT(self):
        from comtypes.gen.TestComServerLib import MYCOLOR

        t = _midlSAFEARRAY(MYCOLOR)
        self.assertTrue(t is _midlSAFEARRAY(MYCOLOR))

        sa = t.from_param([MYCOLOR(0, 0, 0), MYCOLOR(1, 2, 3)])

        self.assertEqual(
            [(x.red, x.green, x.blue) for x in sa[0]],
            [(0.0, 0.0, 0.0), (1.0, 2.0, 3.0)],
        )

        def doit():
            t.from_param([MYCOLOR(0, 0, 0), MYCOLOR(1, 2, 3)])

        bytes = find_memleak(doit, target="private_usage")
        self.assertFalse(bytes, f"Leaks {bytes} bytes")  # type: ignore


class PropPutRefTest(unittest.TestCase):
    def doit(self, dynamic: bool):
        d = CreateObject("Scripting.Dictionary", dynamic=dynamic)
        s = CreateObject("TestComServerLib.TestComServer", dynamic=dynamic)
        s.name = "the value"

        # This calls propputref, since we assign an Object
        d.Item["object"] = s
        # This calls propput, since we assing a Value
        d.Item["value"] = s.name

        self.assertEqual(d.Item["object"], s)
        self.assertEqual(d.Item["object"].name, "the value")
        self.assertEqual(d.Item["value"], "the value")

        # Changing the default property of the object
        s.name = "foo bar"
        self.assertEqual(d.Item["object"], s)
        self.assertEqual(d.Item["object"].name, "foo bar")
        self.assertEqual(d.Item["value"], "the value")

        # This also calls propputref since we assign an Object
        d.Item["var"] = VARIANT(s)
        self.assertEqual(d.Item["var"], s)

    def test_earlybind(self):
        self.doit(dynamic=False)

    def test_latebind(self):
        self.doit(dynamic=True)


class TestEvents(unittest.TestCase):
    def test(self):
        import comtypes.test.test_comserver

        doctest.testmod(
            comtypes.test.test_comserver,
            verbose=False,
            optionflags=doctest.ELLIPSIS,
            raise_on_error=True,
        )


class ShowEventsExamples:
    # The following functions are never called, they only contain doctests:

    def ShowEventsFloat(self):
        """
        >>> from comtypes.client import CreateObject, ShowEvents
        >>>
        >>> o = CreateObject("TestComServerLib.TestComServer")
        >>> con = ShowEvents(o)
        # event found: ITestComServerEvents_EvalStarted
        # event found: ITestComServerEvents_EvalCompleted
        >>> result = o.eval("10. / 4")
        Event ITestComServerEvents_EvalStarted(None, '10. / 4')
        Event ITestComServerEvents_EvalCompleted(None, '10. / 4', VARIANT(vt=0x5, 2.5))
        >>> result
        2.5
        >>>
        """

    # # The following test, if enabled, works but the testsuit
    # # crashes elsewhere.  Is there s problem with SAFEARRAYs?

    # def Fails(self):
    #     """
    #     >>> from comtypes.client import CreateObject, ShowEvents
    #     >>>
    #     >>> o = CreateObject("TestComServerLib.TestComServer")
    #     >>> con = ShowEvents(o)
    #     # event found: ITestComServerEvents_EvalStarted
    #     # event found: ITestComServerEvents_EvalCompleted
    #     >>> result = o.eval("['32'] * 2")
    #     Event ITestComServerEvents_EvalStarted(None, u"['32'] * 2")
    #     Event ITestComServerEvents_EvalCompleted(None, u"['32'] * 2", VARIANT(vt=0x200c, (u'32', u'32')))
    #     >>> result
    #     (u'32', u'32')
    #     >>>
    #     """

    def GetEvents(self):
        """
        >>> from comtypes.client import CreateObject, GetEvents
        >>>
        >>> o =  CreateObject("TestComServerLib.TestComServer")
        >>> class EventHandler(object):
        ...     def EvalStarted(self, this, what):
        ...         print("EvalStarted: %s" % what)
        ...         return 0
        ...     def EvalCompleted(self, this, what, result):
        ...         print("EvalCompleted: %s = %s" % (what, result.value))
        ...         return 0
        ...
        >>>
        >>> con = GetEvents(o, EventHandler())
        >>> o.eval("2 + 3")
        EvalStarted: 2 + 3
        EvalCompleted: 2 + 3 = 5
        5
        >>> del con
        >>> o.eval("3 + 2")
        5
        >>>
        """


if __name__ == "__main__":
    unittest.main()
