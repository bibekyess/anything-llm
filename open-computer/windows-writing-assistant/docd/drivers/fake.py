"""In-memory driver: exercises the full RPC/addressing/render stack off-Windows.

Semantics mirror WordDriver (staleness checks, re-anchoring, style_map) over a
plain list of (text, style) paragraphs, so protocol-level tests and the TS
extension can be developed without Word. Opening an existing .txt/.md file
loads its lines; a missing path starts an empty document.
"""

import os
import re

from .. import addressing, render
from ..errors import (
    DocdError, NO_SUCH_DOC, READ_ONLY, SAVE_FORMAT_UNSUPPORTED,
)
from .base import BaseDriver


class _FakeDoc:
    def __init__(self, path, read_only):
        self.path = path
        self.read_only = read_only
        self.dirty = False
        self.paras = [{"text": "", "style": None}]
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                content = f.read()
            self.paras = [
                {"text": line, "style": _style_of(line)}
                for line in content.splitlines()
            ] or self.paras


def _style_of(line):
    match = re.match(r"^(#{1,9}) ", line)
    return f"Heading {len(match.group(1))}" if match else None


def _strip_md(line):
    return re.sub(r"^#{1,9} ", "", line)


class FakeDriver(BaseDriver):
    prefix = "f"
    backend = "fake"

    def __init__(self):
        self._docs = {}
        self._counter = 0

    # ── helpers ────────────────────────────────────────────────────────
    def _doc(self, handle):
        if handle not in self._docs:
            raise DocdError(NO_SUCH_DOC, f"No open document with handle '{handle}'.")
        return self._docs[handle]

    def _writable(self, handle):
        doc = self._doc(handle)
        if doc.read_only:
            raise DocdError(READ_ONLY, f"Document '{handle}' is open read-only.")
        return doc

    @staticmethod
    def _text_at(doc):
        return lambda i: _strip_md(doc.paras[i]["text"])

    @staticmethod
    def _para_dicts(doc):
        return [
            {"text": _strip_md(p["text"]), "style": p["style"], "outline_level": None}
            for p in doc.paras
        ]

    def _hashes(self, doc, indices):
        return [[i, addressing.para_hash(_strip_md(doc.paras[i]["text"]))] for i in indices]

    # ── BaseDriver ─────────────────────────────────────────────────────
    def list_open(self):
        return {
            "docs": [
                {
                    "doc": handle,
                    "backend": self.backend,
                    "path": d.path,
                    "dirty": d.dirty,
                    "paragraphs": len(d.paras),
                    "read_only": d.read_only,
                }
                for handle, d in self._docs.items()
            ]
        }

    def open(self, path, read_only=False):
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._docs[handle] = _FakeDoc(path, read_only)
        doc = self._docs[handle]
        return {
            "doc": handle,
            "path": path,
            "paragraphs": len(doc.paras),
            "dirty": False,
            "read_only": read_only,
        }

    def read(self, doc, from_para=None, to_para=None, max_chars=None):
        d = self._doc(doc)
        count = len(d.paras)
        start = from_para or 0
        end = min(to_para if to_para is not None else count - 1, count - 1)
        text, _, rev = render.render_read(
            self._para_dicts(d)[start : end + 1], from_para=start, max_chars=max_chars
        )
        return {"text": text, "from_para": start, "to_para": end, "count": count, "rev": rev}

    def outline(self, doc):
        d = self._doc(doc)
        headings = [
            {"level": int(p["style"].split(" ")[1]), "para": i, "text": _strip_md(p["text"])}
            for i, p in enumerate(d.paras)
            if p["style"] and p["style"].startswith("Heading ")
        ]
        return {"text": render.render_outline(headings)}

    def insert(self, doc, text, where, para=None, expect_hash=None,
               bookmark=None, style_map=True):
        d = self._writable(doc)
        count = len(d.paras)
        moved = False
        if where == "end" or where == "cursor":  # fake has no live caret
            at = count
        elif where in ("before_para", "after_para"):
            if para is None:
                raise DocdError("BAD_PARAMS", f"'{where}' requires `para`.")
            idx, moved = addressing.resolve_anchor(
                self._text_at(d), count, para, expect_hash
            )
            at = idx if where == "before_para" else idx + 1
        else:
            raise DocdError("BAD_PARAMS", f"Unsupported insert anchor '{where}' on fake backend.")
        pieces = render.parse_styled_text(text, style_map)
        d.paras[at:at] = [{"text": t, "style": s} for t, s in pieces]
        d.dirty = True
        return {
            "inserted": len(pieces),
            "first_para": at,
            "affected": self._hashes(d, range(at, at + len(pieces))),
            "moved": moved,
        }

    def replace(self, doc, find, replace, regex=False, match_case=False,
                occurrence="all", scope=None):
        d = self._writable(doc)
        flags = 0 if match_case else re.IGNORECASE
        pattern = re.compile(find if regex else re.escape(find), flags)
        lo = scope["from_para"] if scope else 0
        hi = scope["to_para"] if scope else len(d.paras) - 1
        limit = (
            0 if occurrence == "all"
            else 1 if occurrence == "first"
            else int(occurrence)
        )
        seen = replaced = 0
        touched = []
        for i in range(lo, min(hi, len(d.paras) - 1) + 1):
            text = d.paras[i]["text"]
            if limit == 0:
                new, n = pattern.subn(replace, text)
                replaced += n
            else:
                new, n = text, 0
                for match in pattern.finditer(text):
                    seen += 1
                    if seen == limit:
                        new = text[: match.start()] + replace + text[match.end():]
                        n = 1
                        replaced += 1
                        break
                if replaced and n == 0:
                    break
            if n:
                d.paras[i]["text"] = new
                d.paras[i]["style"] = _style_of(new)
                touched.append(i)
            if limit and replaced:
                break
        if touched:
            d.dirty = True
        return {"replaced": replaced, "affected": self._hashes(d, touched)}

    def edit_range(self, doc, from_para, to_para, expect_hashes, new_text):
        d = self._writable(doc)
        addressing.check_range_hashes(
            self._text_at(d), len(d.paras), from_para, to_para, expect_hashes
        )
        pieces = render.parse_styled_text(new_text, True) if new_text else []
        d.paras[from_para : to_para + 1] = [{"text": t, "style": s} for t, s in pieces]
        if not d.paras:
            d.paras = [{"text": "", "style": None}]
        d.dirty = True
        return {
            "replaced": len(pieces),
            "deleted": (to_para - from_para + 1) - len(pieces) if not pieces else 0,
            "affected": self._hashes(d, range(from_para, from_para + len(pieces))),
        }

    def apply_style(self, doc, from_para, to_para=None, style=None):
        d = self._writable(doc)
        end = to_para if to_para is not None else from_para
        if from_para < 0 or end >= len(d.paras):
            raise DocdError("BAD_PARAMS", "Paragraph range out of bounds.")
        for i in range(from_para, end + 1):
            d.paras[i]["style"] = style
        d.dirty = True
        return {"styled": end - from_para + 1, "affected": self._hashes(d, range(from_para, end + 1))}

    def save(self, doc):
        d = self._writable(doc)
        self._write(d, d.path)
        d.dirty = False
        return {"saved": True, "path": d.path}

    def save_as(self, doc, path, format):
        d = self._doc(doc)
        if format not in ("txt", "md"):
            raise DocdError(
                SAVE_FORMAT_UNSUPPORTED,
                f"Fake backend can only save txt/md, not '{format}'.",
            )
        self._write(d, path)
        return {"saved": True, "path": path, "format": format}

    def close(self, doc, discard_changes=False):
        d = self._doc(doc)
        was_dirty = d.dirty
        if was_dirty and not discard_changes:
            self._write(d, d.path)
        del self._docs[doc]
        return {"closed": True, "was_dirty": was_dirty}

    @staticmethod
    def _write(d, path):
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(p["text"] for p in d.paras))
