import pytest

from docd import addressing
from docd.errors import DocdError


def get_text_fn(paras):
    return lambda i: paras[i]


class TestHashing:
    def test_hash_is_stable_and_short(self):
        assert addressing.para_hash("Hello world") == addressing.para_hash("Hello world")
        assert len(addressing.para_hash("Hello world")) == 4

    def test_normalization_ignores_trailing_marks(self):
        # Word paragraph text carries a trailing \r; cells carry \r\x07.
        assert addressing.para_hash("Hello\r") == addressing.para_hash("Hello")
        assert addressing.para_hash("Hello\r\x07") == addressing.para_hash("Hello")

    def test_nfc_normalization_for_hangul(self):
        composed = "한글"           # 한글 (NFC)
        decomposed = "한글"  # NFD jamo
        assert addressing.para_hash(composed) == addressing.para_hash(decomposed)


class TestResolveAnchor:
    def test_exact_match(self):
        paras = ["alpha", "beta", "gamma"]
        h = addressing.para_hash("beta")
        idx, moved = addressing.resolve_anchor(get_text_fn(paras), 3, 1, h)
        assert (idx, moved) == (1, False)

    def test_reanchors_after_insert_above(self):
        # User inserted two paragraphs above: 'beta' moved from index 1 to 3.
        paras = ["new1", "new2", "alpha", "beta", "gamma"]
        h = addressing.para_hash("beta")
        idx, moved = addressing.resolve_anchor(get_text_fn(paras), 5, 1, h)
        assert (idx, moved) == (3, True)

    def test_nearest_match_wins(self):
        paras = ["x", "dup", "x", "dup", "x"]
        h = addressing.para_hash("dup")
        idx, _ = addressing.resolve_anchor(get_text_fn(paras), 5, 2, h)
        assert idx in (1, 3)  # one index away, not the far duplicate

    def test_stale_raises_with_context(self):
        paras = ["alpha", "beta", "gamma"]
        with pytest.raises(DocdError) as exc:
            addressing.resolve_anchor(get_text_fn(paras), 3, 1, "0000")
        assert exc.value.code == "STALE_RANGE"
        context = exc.value.data["context"]
        assert [c["para"] for c in context] == [0, 1, 2]

    def test_no_expect_hash_passes_through(self):
        idx, moved = addressing.resolve_anchor(get_text_fn(["a"]), 1, 0, None)
        assert (idx, moved) == (0, False)


class TestCheckRangeHashes:
    def test_all_match(self):
        paras = ["a", "b", "c"]
        hashes = [addressing.para_hash(p) for p in paras]
        addressing.check_range_hashes(get_text_fn(paras), 3, 0, 2, hashes)

    def test_wrong_count_rejected(self):
        with pytest.raises(DocdError) as exc:
            addressing.check_range_hashes(get_text_fn(["a", "b"]), 2, 0, 1, ["1111"])
        assert exc.value.code == "STALE_RANGE"

    def test_single_mismatch_rejected(self):
        paras = ["a", "b", "c"]
        hashes = [addressing.para_hash(p) for p in paras]
        hashes[1] = "dead"
        with pytest.raises(DocdError) as exc:
            addressing.check_range_hashes(get_text_fn(paras), 3, 0, 2, hashes)
        assert "Paragraph 1" in exc.value.message

    def test_out_of_bounds_rejected(self):
        with pytest.raises(DocdError):
            addressing.check_range_hashes(get_text_fn(["a"]), 1, 0, 5, ["x"] * 6)
