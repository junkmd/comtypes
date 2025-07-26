import itertools
import unittest

from comtypes import CLSCTX_LOCAL_SERVER
from comtypes.client import CreateObject, GetModule

ComtypesCppTestSrvLib_GUID = "{07D2AEE5-1DF8-4D2C-953A-554ADFD25F99}"

try:
    GetModule((ComtypesCppTestSrvLib_GUID, 1, 0, 0))
    import comtypes.gen.ComtypesCppTestSrvLib as ComtypesCppTestSrvLib

    IMPORT_FAILED = False
except (ImportError, OSError):
    IMPORT_FAILED = True


@unittest.skipIf(IMPORT_FAILED, "This depends on the out of process COM-server.")
class TestNamedPropertyPut(unittest.TestCase):
    def _create_dispifc(self) -> "ComtypesCppTestSrvLib.IDispNamedPropertyPutTest":
        # Explicitly ask for the dispinterface of the component.
        return CreateObject(
            "Comtypes.NamedPropertyPutTest",
            clsctx=CLSCTX_LOCAL_SERVER,
            interface=ComtypesCppTestSrvLib.IDispNamedPropertyPutTest,
        )

    def test_initial_values(self):
        dispifc = self._create_dispifc()
        for i, j in itertools.product(range(2), range(3)):
            with self.subTest(i=i, j=j):
                self.assertEqual(dispifc.Value[i, j], 0)
        for i in range(2):
            with self.subTest(i=i):
                self.assertEqual(dispifc.Value[i], (0, 0, 0))
        self.assertEqual(dispifc.Value[()], ((0, 0, 0), (0, 0, 0)))
        self.assertEqual(dispifc.Value[:], ((0, 0, 0), (0, 0, 0)))

    def test_two_args(self):
        for i, j, v in [
            (0, 0, 6),
            (0, 1, 5),
            (0, 2, 4),
            (1, 0, 3),
            (1, 1, 2),
            (1, 2, 1),
        ]:
            with self.subTest(i=i, j=j, v=v):
                dispifc = self._create_dispifc()
                dispifc.Value[i, j] = v
                self.assertEqual(dispifc.Value[i, j], v)

    def test_one_arg(self):
        for index, in_, out in [
            (0, (6, 5, 4), (6, 5, 4)),
            (1, [9, 8, 7], (9, 8, 7)),
        ]:
            dispifc = self._create_dispifc()
            with self.subTest(index=index, in_=in_, out=out):
                dispifc.Value[index] = in_
                self.assertEqual(dispifc.Value[index], out)

    def test_slice(self):
        for in_ in [
            [[11, 22, 33], [44, 55, 66]],  # also check list can be set
            ((11, 22, 33), (44, 55, 66)),
        ]:
            dispifc = self._create_dispifc()
            with self.subTest(in_=in_):
                dispifc.Value[:] = in_
                self.assertEqual(dispifc.Value[:], ((11, 22, 33), (44, 55, 66)))

    def test_empty_tuple(self):
        for in_ in [
            [[101, 102, 103], [104, 105, 106]],  # also check list can be set
            ((101, 102, 103), (104, 105, 106)),
        ]:
            dispifc = self._create_dispifc()
            dispifc.Value[()] = in_
            self.assertEqual(dispifc.Value[()], ((101, 102, 103), (104, 105, 106)))

    def test_repr(self):
        dispifc = self._create_dispifc()
        self.assertIn("<bound_named_property", repr(dispifc.Value))
        self.assertIn("<named_property", repr(type(dispifc).Value))
