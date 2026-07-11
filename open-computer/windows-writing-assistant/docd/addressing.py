"""Stable paragraph addressing: content hashes, re-anchoring, staleness.

Design doc §2.2: paragraph indices are assigned at read time; every mutating
call carries the expected content hash of its anchor paragraph(s). At execute
time the driver re-resolves the index, compares hashes, and searches ±WINDOW
indices for a moved anchor before failing with STALE_RANGE.

Pure stdlib so it is unit-testable off-Windows and shared by all drivers.
"""

import hashlib
import unicodedata

from .errors import DocdError, STALE_RANGE

REANCHOR_WINDOW = 8
HASH_LEN = 4


def normalize(text):
    """NFC-normalize and strip trailing whitespace / paragraph & cell marks."""
    return unicodedata.normalize("NFC", text).rstrip("\r\n\x07 \t")


def para_hash(text):
    """4-hex-char content hash, e.g. '3fa2' in '[p0#3fa2]'."""
    return hashlib.sha1(normalize(text).encode("utf-8")).hexdigest()[:HASH_LEN]


def snapshot_rev(hashes):
    """Whole-snapshot rev token from the per-paragraph hashes."""
    return "r" + hashlib.sha1("".join(hashes).encode("utf-8")).hexdigest()[:6]


def resolve_anchor(get_text, count, index, expect_hash, window=REANCHOR_WINDOW):
    """Resolve a paragraph anchor, absorbing index drift.

    get_text(i) -> current text of paragraph i (0-based); count = paragraph
    count. Returns (resolved_index, moved: bool). Raises STALE_RANGE with ±2
    paragraphs of context when the hash is found nowhere in the window.
    """
    if count == 0:
        raise DocdError(STALE_RANGE, "Document has no paragraphs.")
    index = max(0, min(index, count - 1))
    if expect_hash is None:
        return index, False
    if para_hash(get_text(index)) == expect_hash:
        return index, False
    # Search outward: index±1, index±2, ... so the nearest match wins.
    for delta in range(1, window + 1):
        for cand in (index - delta, index + delta):
            if 0 <= cand < count and para_hash(get_text(cand)) == expect_hash:
                return cand, True
    raise DocdError(
        STALE_RANGE,
        f"Paragraph {index} no longer matches hash #{expect_hash} "
        f"(searched ±{window}). Re-read before editing.",
        data={"context": _context(get_text, count, index)},
    )


def check_range_hashes(get_text, count, from_para, to_para, expect_hashes):
    """Validate every hash in [from_para, to_para]; raise STALE_RANGE on any mismatch."""
    if from_para < 0 or to_para >= count or from_para > to_para:
        raise DocdError(
            STALE_RANGE,
            f"Range p{from_para}..p{to_para} out of bounds (document has {count} paragraphs).",
        )
    span = to_para - from_para + 1
    if len(expect_hashes) != span:
        raise DocdError(
            STALE_RANGE,
            f"expect_hashes has {len(expect_hashes)} entries for a {span}-paragraph range.",
        )
    for offset, expected in enumerate(expect_hashes):
        i = from_para + offset
        if para_hash(get_text(i)) != expected:
            raise DocdError(
                STALE_RANGE,
                f"Paragraph {i} changed since the last read (expected #{expected}).",
                data={"context": _context(get_text, count, i)},
            )


def _context(get_text, count, index, radius=2):
    """±radius paragraphs of current text, for STALE_RANGE error data."""
    out = []
    for i in range(max(0, index - radius), min(count, index + radius + 1)):
        text = normalize(get_text(i))
        out.append({"para": i, "hash": para_hash(text), "text": text[:200]})
    return out
