"""Tests for the code-generator "head" writers in ``comtypes.tools.codegenerator.heads``.

The head writers turn the parsed type descriptions into the first lines of
the generated Python classes. A wrong line here means the generated module
does not import or does not expose the expected members, so the exact output
is worth pinning down.
"""

import io
import unittest

from comtypes.tools import typedesc
from comtypes.tools.codegenerator import heads


class TestStructureHeadWriter(unittest.TestCase):
    def test_writes_structure_without_recordinfo(self):
        # A plain structure with no record info gets an empty body. The
        # generated class must still be a ``Structure`` subclass.
        struct = typedesc.Structure("Point", align=32, size=64)
        stream = io.StringIO()
        heads.StructureHeadWriter(stream).write(struct.get_head())
        self.assertEqual(stream.getvalue(), "class Point(Structure):\n    pass\n")

    def test_writes_structure_with_recordinfo(self):
        # When the structure carries record info it must be emitted as the
        # ``_recordinfo_`` class attribute, otherwise the generated wrapper
        # loses the information needed to marshal the record.
        recordinfo = ("{00000000-0000-0000-0000-000000000000}", 1, 2, 3, "Point")
        struct = typedesc.Structure("Point", align=32, size=64, recordinfo=recordinfo)
        stream = io.StringIO()
        heads.StructureHeadWriter(stream).write(struct.get_head())
        self.assertEqual(
            stream.getvalue(),
            f"class Point(Structure):\n    _recordinfo_ = {recordinfo!r}\n",
        )

    def test_writes_location_comment(self):
        # The source location is emitted as a comment so a reader can trace
        # the generated class back to the header it came from.
        struct = typedesc.Structure("Point", align=32, size=64)
        struct.location = "foo.h:1"
        stream = io.StringIO()
        heads.StructureHeadWriter(stream).write(struct.get_head())
        self.assertEqual(
            stream.getvalue(),
            "# foo.h:1\nclass Point(Structure):\n    pass\n",
        )

    def test_writes_union(self):
        # A union must be emitted as a ``Union`` subclass, not a ``Structure``.
        union = typedesc.Union("U", align=32, size=64)
        stream = io.StringIO()
        heads.StructureHeadWriter(stream).write(union.get_head())
        self.assertEqual(stream.getvalue(), "class U(Union):\n    pass\n")


class TestComInterfaceHeadWriter(unittest.TestCase):
    def test_returns_early_when_base_is_none(self):
        # IUnknown has no base interface, so the writer must not emit a class
        # for it. Emitting one would produce a duplicate/incorrect definition.
        itf = typedesc.ComInterface(
            "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", [], None
        )
        stream = io.StringIO()
        heads.ComInterfaceHeadWriter(stream).write(itf.get_head(), "IUnknown")
        self.assertEqual(stream.getvalue(), "")

    def test_writes_enum_interface_iteration_helpers(self):
        # An IEnumXXX interface that has Next/Skip/Reset/Clone gets __iter__,
        # __next__ and __getitem__ helpers so callers can iterate it directly
        # instead of calling Next() by hand.
        base = typedesc.ComInterface(
            "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", [], None
        )
        itf = typedesc.ComInterface(
            "IEnumFoo", base, "{00000000-0000-0000-0000-000000000001}", [], None
        )
        itf.extend_members(
            [
                typedesc.ComMethod(1, 0, "Next", "HRESULT", [], None),
                typedesc.ComMethod(1, 0, "Skip", "HRESULT", [], None),
                typedesc.ComMethod(1, 0, "Reset", "HRESULT", [], None),
                typedesc.ComMethod(1, 0, "Clone", "HRESULT", [], None),
            ]
        )
        stream = io.StringIO()
        heads.ComInterfaceHeadWriter(stream).write(itf.get_head(), "IUnknown")
        output = stream.getvalue()
        self.assertIn("    def __iter__(self):", output)
        self.assertIn("    def __next__(self):", output)
        self.assertIn("    def __getitem__(self, index):", output)

    def test_does_not_write_iteration_helpers_for_non_enum_interface(self):
        # A regular interface must not get the iteration helpers, even if it
        # happens to have a method named Next. Only the full IEnumXXX shape
        # (Next/Skip/Reset/Clone) qualifies.
        base = typedesc.ComInterface(
            "IUnknown", None, "{00000000-0000-0000-C000-000000000046}", [], None
        )
        itf = typedesc.ComInterface(
            "IFoo", base, "{00000000-0000-0000-0000-000000000002}", [], None
        )
        itf.extend_members([typedesc.ComMethod(1, 0, "Next", "HRESULT", [], None)])
        stream = io.StringIO()
        heads.ComInterfaceHeadWriter(stream).write(itf.get_head(), "IUnknown")
        self.assertNotIn("def __iter__", stream.getvalue())


if __name__ == "__main__":
    unittest.main()
