"""Rendering and parsing shared by all drivers.

doc_read output format (design doc §2.1):
    [p0#3fa2] # 2026 Business Plan     <- Heading 1 rendered as '#'
    [p1#91cc] Body text ...
plus a trailing 'rev:' token.

doc_insert markdown parsing: LLMs write markdown natively, so style_map maps
the common subset to real word-processor formatting instead of leaving raw
markers in the document — headings, bullet/numbered lists, checkboxes,
> quotes, **bold**, *italic*, ***both***, `code`.
"""

import re

from .addressing import normalize, para_hash, snapshot_rev

MAX_READ_CHARS_DEFAULT = 20000

# Normalized style names (design doc §2.1 doc_apply_style).
HEADING_STYLES = {i: f"Heading {i}" for i in range(1, 10)}
BODY_STYLE = "Normal"


def style_to_md_prefix(style_name, outline_level):
    """'Heading 2' / outline level 2 -> '## '. Body text -> ''."""
    if style_name and style_name.startswith("Heading "):
        try:
            level = int(style_name.split(" ", 1)[1])
            return "#" * min(level, 9) + " "
        except ValueError:
            pass
    if outline_level and 1 <= outline_level <= 9:
        return "#" * outline_level + " "
    return ""


# Inline markdown: ***bold-italic***, **bold**, *italic*, _italic_, `code`.
# Underscore italics require non-word boundaries so snake_case survives.
_INLINE_RE = re.compile(
    r"\*\*\*(?P<bi>[^*]+?)\*\*\*"
    r"|\*\*(?P<b>.+?)\*\*"
    r"|\*(?P<i>[^*\s](?:[^*]*?[^*\s])?)\*"
    r"|(?<!\w)_(?P<iu>[^_\s](?:[^_]*?[^_\s])?)_(?!\w)"
    r"|`(?P<c>[^`]+)`"
)
_BULLET_RE = re.compile(r"^[-*+]\s+(.*)$")
_NUMBER_RE = re.compile(r"^\d{1,3}[.)]\s+(.*)$")
_CHECKBOX_RE = re.compile(r"^\[([ xX])\]\s*(.*)$")


def _parse_inline(text):
    """Strip inline markers; return (plain_text, [(start, end, fmt), ...])
    with offsets into the plain text."""
    parts, spans = [], []
    pos = out_len = 0
    for m in _INLINE_RE.finditer(text):
        parts.append(text[pos : m.start()])
        out_len += m.start() - pos
        if m.group("bi") is not None:
            content, fmt = m.group("bi"), {"bold": True, "italic": True}
        elif m.group("b") is not None:
            content, fmt = m.group("b"), {"bold": True}
        elif m.group("i") is not None:
            content, fmt = m.group("i"), {"italic": True}
        elif m.group("iu") is not None:
            content, fmt = m.group("iu"), {"italic": True}
        else:
            content, fmt = m.group("c"), {"code": True}
        # Nested emphasis inside a span isn't parsed further (rare from LLMs).
        parts.append(content)
        spans.append((out_len, out_len + len(content), fmt))
        out_len += len(content)
        pos = m.end()
    parts.append(text[pos:])
    return "".join(parts), spans


def parse_markdown(text, style_map=True):
    """Split insert text into paragraph dicts:
        {"text": plain, "style": style_or_None, "spans": [(start, end, fmt)]}

    '\\n' separates paragraphs. With style_map, block markers map to styles
    (# -> Heading N, -/* -> List Bullet, 1. -> List Number, > -> Quote,
    checkboxes -> visible box glyphs) and inline markers become format spans.
    """
    out = []
    for line in text.split("\n"):
        style = None
        work = line
        if style_map:
            stripped = line.lstrip()
            hashes = len(stripped) - len(stripped.lstrip("#"))
            m_bullet = _BULLET_RE.match(stripped)
            m_number = _NUMBER_RE.match(stripped)
            if 0 < hashes <= 9 and stripped[hashes : hashes + 1] == " ":
                style, work = HEADING_STYLES[hashes], stripped[hashes + 1 :]
            elif m_bullet:
                style, work = "List Bullet", m_bullet.group(1)
                m_box = _CHECKBOX_RE.match(work)
                if m_box:
                    glyph = "☐" if m_box.group(1) == " " else "☑"  # ☐ / ☑
                    work = f"{glyph} {m_box.group(2)}"
            elif m_number:
                style, work = "List Number", m_number.group(1)
            elif stripped.startswith(">"):
                style, work = "Quote", stripped[1:].lstrip()
        plain, spans = _parse_inline(work) if style_map else (work, [])
        out.append({"text": plain, "style": style, "spans": spans})
    return out


def flatten_payload(paras):
    """Join paragraph dicts into one '\\r'-separated payload plus format spans
    with offsets absolute within the payload (Word Range arithmetic)."""
    payload = "\r".join(p["text"] for p in paras)
    spans, offset = [], 0
    for p in paras:
        spans.extend((offset + s, offset + e, f) for s, e, f in p["spans"])
        offset += len(p["text"]) + 1  # +1 for the \r paragraph mark
    return payload, spans


def parse_styled_text(text, style_map=True):
    """Back-compat shim: [(paragraph_text, style_or_None), ...]."""
    return [(p["text"], p["style"]) for p in parse_markdown(text, style_map)]


def render_read(paragraphs, from_para=0, max_chars=None):
    """paragraphs: [{text, style, outline_level, in_table}] -> (text, hashes, rev).

    Hashes cover the rendered slice only; rev covers the same slice (drivers
    may override rev with a whole-doc token).
    """
    max_chars = max_chars or MAX_READ_CHARS_DEFAULT
    lines, hashes = [], []
    used = 0
    truncated = False
    for offset, para in enumerate(paragraphs):
        text = normalize(para["text"])
        h = para_hash(text)
        hashes.append(h)
        prefix = style_to_md_prefix(para.get("style"), para.get("outline_level"))
        marker = " (table)" if para.get("in_table") else ""
        line = f"[p{from_para + offset}#{h}]{marker} {prefix}{text}"
        if used + len(line) > max_chars:
            truncated = True
            break
        lines.append(line)
        used += len(line) + 1
    rev = snapshot_rev(hashes)
    if truncated:
        lines.append(f"... truncated at {max_chars} chars; continue with from_para={from_para + len(lines)}")
    lines.append(f"rev:{rev}")
    return "\n".join(lines), hashes, rev


def render_outline(headings):
    """headings: [{level, para, text}] -> 'H1 [p0] Overview' lines."""
    if not headings:
        return "(no headings)"
    return "\n".join(
        f"{'  ' * (h['level'] - 1)}H{h['level']} [p{h['para']}] {normalize(h['text'])}"
        for h in headings
    )


def affected(paras_with_hashes):
    """[(index, hash)] -> 'Affected: [p4#a1b2] [p9#77de]'."""
    if not paras_with_hashes:
        return "Affected: (none)"
    return "Affected: " + " ".join(f"[p{i}#{h}]" for i, h in paras_with_hashes)
