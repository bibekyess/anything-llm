"""Rendering and parsing shared by all drivers.

doc_read output format (design doc §2.1):
    [p0#3fa2] # 2026 Business Plan     <- Heading 1 rendered as '#'
    [p1#91cc] Body text ...
plus a trailing 'rev:' token.

doc_insert style_map parsing: markdown '#'-prefixes map to Heading styles.
"""

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


def parse_styled_text(text, style_map=True):
    """Split insert text into [(paragraph_text, style_or_None), ...].

    '\\n' separates paragraphs; with style_map, leading '#'-runs become
    Heading styles and are stripped from the text.
    """
    out = []
    for line in text.split("\n"):
        style = None
        if style_map:
            stripped = line.lstrip()
            hashes = len(stripped) - len(stripped.lstrip("#"))
            if 0 < hashes <= 9 and stripped[hashes : hashes + 1] == " ":
                style = HEADING_STYLES[hashes]
                line = stripped[hashes + 1 :]
        out.append((line, style))
    return out


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
