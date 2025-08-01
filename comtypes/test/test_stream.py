import unittest as ut
from ctypes import HRESULT, POINTER, OleDLL, byref, c_ubyte, c_ulonglong, pointer
from ctypes.wintypes import BOOL, HGLOBAL, ULARGE_INTEGER

import comtypes.client

comtypes.client.GetModule("portabledeviceapi.dll")
from comtypes.gen.PortableDeviceApiLib import IStream

STATFLAG_DEFAULT = 0
STGC_DEFAULT = 0
STGTY_STREAM = 2
STREAM_SEEK_SET = 0
STREAM_SEEK_CUR = 1
STREAM_SEEK_END = 2

_ole32 = OleDLL("ole32")

_CreateStreamOnHGlobal = _ole32.CreateStreamOnHGlobal
_CreateStreamOnHGlobal.argtypes = [HGLOBAL, BOOL, POINTER(POINTER(IStream))]
_CreateStreamOnHGlobal.restype = HRESULT

_shlwapi = OleDLL("shlwapi")

_IStream_Size = _shlwapi.IStream_Size
_IStream_Size.argtypes = [POINTER(IStream), POINTER(ULARGE_INTEGER)]
_IStream_Size.restype = HRESULT


def _create_stream() -> IStream:
    # Create an IStream
    stream = POINTER(IStream)()  # type: ignore
    _CreateStreamOnHGlobal(None, True, byref(stream))
    return stream  # type: ignore


class Test_RemoteWrite(ut.TestCase):
    def test_RemoteWrite(self):
        stream = _create_stream()
        test_data = b"Some data"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))

        written = stream.RemoteWrite(pv, len(test_data))

        # Verification
        self.assertEqual(written, len(test_data))


class Test_RemoteRead(ut.TestCase):
    def test_RemoteRead(self):
        stream = _create_stream()
        test_data = b"Some data"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        stream.RemoteWrite(pv, len(test_data))

        # Make sure the data actually gets written before trying to read back
        stream.Commit(STGC_DEFAULT)
        # Move the stream back to the beginning
        stream.RemoteSeek(0, STREAM_SEEK_SET)

        buffer_size = 1024

        read_buffer, data_read = stream.RemoteRead(buffer_size)

        # Verification
        self.assertEqual(data_read, len(test_data))
        self.assertEqual(bytearray(read_buffer)[0:data_read], test_data)


class Test_RemoteSeek(ut.TestCase):
    def _create_sample_stream(self) -> IStream:
        stream = _create_stream()
        test_data = b"spam egg bacon ham"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        stream.RemoteWrite(pv, len(test_data))
        stream.Commit(STGC_DEFAULT)
        return stream

    def test_takes_STREAM_SEEK_SET_as_origin(self):
        stream = self._create_sample_stream()
        newpos = stream.RemoteSeek(9, STREAM_SEEK_SET)
        self.assertEqual(newpos, 9)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"bacon ham")

    def test_takes_STREAM_SEEK_CUR_as_origin(self):
        stream = self._create_sample_stream()
        stream.RemoteSeek(8, STREAM_SEEK_SET)
        newpos = stream.RemoteSeek(7, STREAM_SEEK_CUR)
        self.assertEqual(newpos, 15)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"ham")

    def test_takes_STREAM_SEEK_END_as_origin(self):
        stream = self._create_sample_stream()
        stream.RemoteSeek(8, STREAM_SEEK_SET)
        newpos = stream.RemoteSeek(-13, STREAM_SEEK_END)
        self.assertEqual(newpos, 5)
        buf, read = stream.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], b"egg bacon ham")


class Test_SetSize(ut.TestCase):
    def test_SetSize(self):
        stream = _create_stream()
        stream.SetSize(42)
        pui = pointer(c_ulonglong())
        _IStream_Size(stream, pui)
        self.assertEqual(pui.contents.value, 42)


class Test_RemoteCopyTo(ut.TestCase):
    def test_RemoteCopyTo(self):
        src = _create_stream()
        dst = _create_stream()
        test_data = b"parrot"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        src_written = src.RemoteWrite(pv, len(test_data))
        src.Commit(STGC_DEFAULT)
        src.RemoteSeek(0, STREAM_SEEK_SET)
        cpy_read, cpy_written = src.RemoteCopyTo(dst, src_written)
        self.assertEqual(cpy_read, len(test_data))
        self.assertEqual(cpy_written, len(test_data))
        dst.Commit(STGC_DEFAULT)
        dst.RemoteSeek(0, STREAM_SEEK_SET)
        dst_buf, dst_read = dst.RemoteRead(1024)
        self.assertEqual(bytearray(dst_buf)[0:dst_read], test_data)


class Test_Stat(ut.TestCase):
    # https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-istream-stat
    # https://learn.microsoft.com/en-us/windows/win32/api/objidl/ns-objidl-statstg
    def test_returns_statstg_from_no_modified_stream(self):
        stream = _create_stream()
        statstg = stream.Stat(STATFLAG_DEFAULT)
        self.assertIsNone(statstg.pwcsName)
        self.assertEqual(statstg.type, STGTY_STREAM)
        self.assertEqual(statstg.cbSize, 0)
        mt, ct, at = statstg.mtime, statstg.ctime, statstg.atime
        self.assertTrue(mt.dwLowDateTime == ct.dwLowDateTime == at.dwLowDateTime)
        self.assertTrue(mt.dwHighDateTime == ct.dwHighDateTime == at.dwHighDateTime)
        self.assertEqual(statstg.grfMode, 0)
        self.assertEqual(statstg.grfLocksSupported, 0)
        self.assertEqual(statstg.clsid, comtypes.GUID())
        self.assertEqual(statstg.grfStateBits, 0)


class Test_Clone(ut.TestCase):
    def test_Clone(self):
        orig = _create_stream()
        test_data = b"spam egg bacon ham"
        pv = (c_ubyte * len(test_data)).from_buffer(bytearray(test_data))
        orig.RemoteWrite(pv, len(test_data))
        orig.Commit(STGC_DEFAULT)
        orig.RemoteSeek(0, STREAM_SEEK_SET)
        new_stm = orig.Clone()
        buf, read = new_stm.RemoteRead(1024)
        self.assertEqual(bytearray(buf)[0:read], test_data)


if __name__ == "__main__":
    ut.main()
