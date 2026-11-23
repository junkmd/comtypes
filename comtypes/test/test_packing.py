import unittest

from comtypes.tools import typedesc
from comtypes.tools.codegenerator.packing import (
    PackingError,
    _calc_packing,
    calc_packing,
    storage,
)


# Sizes, alignments and offsets are expressed in bits, matching the values
# produced by ``comtypes.tools.tlbparser`` (e.g. a 32-bit ``int``).
def _make_field(name, size, align, offset):
    typ = typedesc.FundamentalType("int", size, align)
    return typedesc.Field(name, typ, None, offset)


class CalcPackingTest(unittest.TestCase):
    def test_returns_none_for_default_packing(self):
        # A naturally laid out single-field struct needs no explicit packing.
        field = _make_field("a", 32, 32, 0)
        struct = typedesc.Structure("Good", align=32, size=32)
        struct.add_member(field)
        self.assertIsNone(calc_packing(struct, [field]))

    def test_incomplete_struct_returns_none(self):
        field = _make_field("a", 32, 32, 0)
        struct = typedesc.Structure("Incomplete", align=32, size=None)
        struct.add_member(field)
        self.assertIsNone(calc_packing(struct, [field]))

    def test_raises_packing_error_with_details_when_layout_is_inconsistent(self):
        # The declared field offset does not match the computed layout, so every
        # packing attempt raises ``PackingError`` and ``calc_packing`` must
        # re-raise with the underlying reason.
        #
        # Regression test for gh-937: the final ``raise`` referenced the
        # ``except ... as`` target after it had been cleared, raising
        # ``UnboundLocalError`` and masking the real ``PackingError``.
        field = _make_field("x", 32, 32, 64)
        struct = typedesc.Structure("Bad", align=32, size=96)
        struct.add_member(field)
        with self.assertRaises(PackingError) as ctx:
            calc_packing(struct, [field])
        message = str(ctx.exception)
        self.assertIn("PACKING FAILED", message)
        # The original failure reason must be preserved, not lost.
        self.assertIn("field x offset", message)


class CalcPackingExplicitTest(unittest.TestCase):
    def test_returns_pack_size_when_explicit_packing_is_required(self):
        # A struct whose declared alignment is smaller than the natural
        # alignment of its field cannot be laid out without an explicit
        # packing. ``calc_packing`` must report the packing in bytes.
        field = typedesc.Field("a", typedesc.FundamentalType("int", 8, 32), None, 0)
        struct = typedesc.Structure("Packed", align=8, size=8)
        struct.add_member(field)
        self.assertEqual(calc_packing(struct, [field]), 1)

    def test_returns_none_for_struct_in_dont_assert_size(self):
        # Some compiler-generated structs are known to have a size that does
        # not match their layout. They are listed in ``dont_assert_size`` and
        # must be accepted without any packing.
        field = typedesc.Field("a", typedesc.FundamentalType("int", 32, 32), None, 0)
        struct = typedesc.Structure("__si_class_type_info_pseudo", align=32, size=32)
        struct.add_member(field)
        self.assertIsNone(calc_packing(struct, [field]))

    def test_returns_minus_two_for_bit_field(self):
        # ``_calc_packing`` cannot handle bit fields and signals this with -2
        # instead of raising, so the caller can try another packing.
        field = typedesc.Field("a", typedesc.FundamentalType("int", 32, 32), 1, 0)
        struct = typedesc.Structure("Bits", align=32, size=32)
        struct.add_member(field)
        self.assertEqual(_calc_packing(struct, [field], None, True), -2)

    def test_union_size_is_max_of_member_sizes(self):
        # In a union all members start at offset 0, so the size is the largest
        # member size, not the sum of the member sizes.
        small = typedesc.Field("a", typedesc.FundamentalType("char", 8, 8), None, 0)
        large = typedesc.Field("b", typedesc.FundamentalType("int", 32, 32), None, 0)
        union = typedesc.Union("U", align=32, size=32)
        union.add_member(small)
        union.add_member(large)
        self.assertIsNone(calc_packing(union, [small, large]))

    def test_raises_packing_error_for_total_alignment_mismatch(self):
        # The computed alignment must match the declared one; otherwise the
        # struct layout is inconsistent and ``_calc_packing`` must complain.
        field = typedesc.Field("a", typedesc.FundamentalType("int", 32, 32), None, 0)
        struct = typedesc.Structure("Bad", align=16, size=32)
        struct.add_member(field)
        with self.assertRaises(PackingError):
            _calc_packing(struct, [field], None, True)

    def test_raises_packing_error_for_total_size_mismatch(self):
        # The computed size must match the declared one; otherwise the struct
        # layout is inconsistent and ``_calc_packing`` must complain.
        field = typedesc.Field("a", typedesc.FundamentalType("char", 8, 8), None, 0)
        struct = typedesc.Structure("Bad", align=8, size=16)
        struct.add_member(field)
        with self.assertRaises(PackingError):
            _calc_packing(struct, [field], None, True)


class StorageTest(unittest.TestCase):
    def test_returns_size_and_align_of_fundamental_type(self):
        # ``storage`` reports the size and alignment of a type in bits.
        typ = typedesc.FundamentalType("int", 32, 32)
        self.assertEqual(storage(typ), (32, 32))

    def test_resolves_typedef_to_underlying_type(self):
        # A typedef has no size of its own; ``storage`` must look through it
        # to the type it aliases.
        typ = typedesc.Typedef("T", typedesc.FundamentalType("int", 32, 32))
        self.assertEqual(storage(typ), (32, 32))

    def test_multiplies_array_element_size_by_element_count(self):
        # An array's size is the element size times the number of elements
        # (max - min + 1), while the alignment stays that of the element.
        typ = typedesc.ArrayType(typedesc.FundamentalType("int", 32, 32), 0, 3)
        self.assertEqual(storage(typ), (128, 32))


if __name__ == "__main__":
    unittest.main()
