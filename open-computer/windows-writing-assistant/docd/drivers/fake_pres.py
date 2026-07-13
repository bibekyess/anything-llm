"""In-memory presentation driver: the slide_* contract, testable on any OS.

Mirrors PptDriver semantics — slide index + hash re-anchoring, stable shape
ids, placeholder kinds — over plain dicts.
"""

from .. import addressing
from ..errors import (
    DocdError, BAD_PARAMS, NO_SUCH_DOC, READ_ONLY, SAVE_FORMAT_UNSUPPORTED,
    STALE_RANGE, UNSUPPORTED_ON_BACKEND,
)
from .base import BaseDriver

SLIDE_REANCHOR_WINDOW = 3


class _Deck:
    def __init__(self, path):
        self.path = path
        self.read_only = False
        self.dirty = False
        self.next_shape_id = 1
        self.slides = []
        self.add_slide(None, "Title Slide", None, None)

    def add_slide(self, at, layout, title, body):
        shapes = []
        for kind, text in (("title", title), ("body", body)):
            shapes.append({"id": self.next_shape_id, "placeholder": kind, "text": text or ""})
            self.next_shape_id += 1
        slide = {"layout": layout or "Title and Content", "shapes": shapes, "notes": ""}
        if at is None:
            self.slides.append(slide)
        else:
            self.slides.insert(at, slide)
        return slide


def _slide_hash(slide):
    return addressing.para_hash("\n".join(s["text"] for s in slide["shapes"] if s["text"]))


class FakePresDriver(BaseDriver):
    prefix = "q"
    backend = "fakepres"

    def __init__(self):
        self._decks = {}
        self._counter = 0

    # ── plumbing ───────────────────────────────────────────────────────
    def _deck(self, handle):
        if handle not in self._decks:
            raise DocdError(NO_SUCH_DOC, f"No open presentation with handle '{handle}'.")
        return self._decks[handle]

    def _writable(self, handle):
        deck = self._deck(handle)
        if deck.read_only:
            raise DocdError(READ_ONLY, f"Presentation '{handle}' is read-only.")
        return deck

    def _resolve_slide(self, deck, index, expect_hash=None):
        count = len(deck.slides)
        if count == 0:
            raise DocdError(BAD_PARAMS, "Presentation has no slides.")
        index = max(0, min(index, count - 1))
        if expect_hash is None or _slide_hash(deck.slides[index]) == expect_hash:
            return index
        for delta in range(1, SLIDE_REANCHOR_WINDOW + 1):
            for cand in (index - delta, index + delta):
                if 0 <= cand < count and _slide_hash(deck.slides[cand]) == expect_hash:
                    return cand
        raise DocdError(STALE_RANGE, f"Slide {index} no longer matches hash #{expect_hash}.")

    @staticmethod
    def _find_shape(slide, shape_id):
        if not shape_id or not shape_id.startswith("sh"):
            raise DocdError(BAD_PARAMS, "shape must be an id from slide_read, e.g. 'sh5'.")
        target = int(shape_id[2:])
        for s in slide["shapes"]:
            if s["id"] == target:
                return s
        raise DocdError(BAD_PARAMS, f"No shape '{shape_id}' on this slide.")

    # ── shared doc surface ─────────────────────────────────────────────
    def list_open(self):
        return {"docs": [
            {"doc": h, "backend": self.backend, "path": d.path, "dirty": d.dirty,
             "slides": len(d.slides), "read_only": d.read_only}
            for h, d in self._decks.items()
        ]}

    def open(self, path, read_only=False):
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        deck = _Deck(path)
        deck.read_only = read_only
        self._decks[handle] = deck
        return {"doc": handle, "path": path, "slides": len(deck.slides),
                "dirty": False, "read_only": read_only,
                "hint": "This is a presentation: use slide_list/slide_read, not doc_read."}

    def new_doc(self):
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._decks[handle] = _Deck(None)
        return {"doc": handle, "path": None, "slides": 1}

    def _unsupported(self, alt):
        raise DocdError(UNSUPPORTED_ON_BACKEND,
                        f"Presentations are slide/shape-oriented — use {alt} instead.")

    def read(self, doc, **kw):
        self._unsupported("slide_list + slide_read")

    def outline(self, doc, **kw):
        self._unsupported("slide_list")

    def insert(self, doc, **kw):
        self._unsupported("slide_add / slide_edit_text")

    def replace(self, doc, **kw):
        self._unsupported("slide_edit_text")

    def edit_range(self, doc, **kw):
        self._unsupported("slide_edit_text")

    def apply_style(self, doc, **kw):
        self._unsupported("slide_edit_text")

    def tables(self, doc, **kw):
        self._unsupported("slide_edit_text")

    def selection(self, doc, **kw):
        self._unsupported("slide_read")

    def save(self, doc):
        deck = self._writable(doc)
        if not deck.path:
            raise DocdError(BAD_PARAMS,
                            "This presentation has never been saved; use pres_save_as.")
        deck.dirty = False
        return {"saved": True, "path": deck.path}

    def save_as(self, doc, path, format):
        return self.pres_save_as(doc, path, format)

    def close(self, doc, discard_changes=False):
        deck = self._deck(doc)
        was_dirty = deck.dirty
        del self._decks[doc]
        return {"closed": True, "was_dirty": was_dirty}

    # ── slide_* tools ──────────────────────────────────────────────────
    def slide_list(self, doc):
        deck = self._deck(doc)
        return {"count": len(deck.slides), "slides": [
            {"slide": i, "hash": _slide_hash(sl), "layout": sl["layout"],
             "title": next((s["text"] for s in sl["shapes"] if s["placeholder"] == "title"), ""),
             "notes": bool(sl["notes"])}
            for i, sl in enumerate(deck.slides)
        ]}

    def slide_read(self, doc, slide):
        deck = self._deck(doc)
        idx = self._resolve_slide(deck, slide)
        sl = deck.slides[idx]
        return {
            "slide": idx,
            "slide_hash": _slide_hash(sl),
            "shapes": [
                {"shape": f"sh{s['id']}", "hash": addressing.para_hash(s["text"]),
                 "placeholder": s["placeholder"], "text": s["text"]}
                for s in sl["shapes"] if s["text"]
            ],
            "notes": sl["notes"],
        }

    def slide_add(self, doc, after_slide=None, layout=None, title=None,
                  body=None, duplicate_of=None):
        deck = self._writable(doc)
        if duplicate_of is not None:
            src = deck.slides[self._resolve_slide(deck, duplicate_of)]
            at = (after_slide + 1) if after_slide is not None else len(deck.slides)
            clone = {
                "layout": src["layout"], "notes": src["notes"],
                "shapes": [
                    {"id": deck.next_shape_id + k, "placeholder": s["placeholder"], "text": s["text"]}
                    for k, s in enumerate(src["shapes"])
                ],
            }
            deck.next_shape_id += len(src["shapes"])
            deck.slides.insert(at, clone)
            new, idx = clone, at
        else:
            at = (after_slide + 1) if after_slide is not None else len(deck.slides)
            new = deck.add_slide(at, layout, title, self._strip_indents(body))
            idx = at
        deck.dirty = True
        return {"slide": idx, "hash": _slide_hash(new), "layout": new["layout"],
                "count": len(deck.slides)}

    @staticmethod
    def _strip_indents(body):
        if body is None:
            return None
        return "\n".join(line.lstrip("\t") for line in body.split("\n"))

    def slide_edit_text(self, doc, slide, shape, text, expect_hash=None):
        deck = self._writable(doc)
        idx = self._resolve_slide(deck, slide)
        sh = self._find_shape(deck.slides[idx], shape)
        if expect_hash is not None and addressing.para_hash(sh["text"]) != expect_hash:
            raise DocdError(STALE_RANGE,
                            f"Shape {shape} changed since slide_read. Re-read the slide.")
        sh["text"] = self._strip_indents(text)
        deck.dirty = True
        return {"slide": idx, "shape": shape, "hash": addressing.para_hash(sh["text"])}

    def slide_notes_edit(self, doc, slide, text):
        deck = self._writable(doc)
        idx = self._resolve_slide(deck, slide)
        deck.slides[idx]["notes"] = text
        deck.dirty = True
        return {"slide": idx, "notes_chars": len(text)}

    def slide_reorder(self, doc, slide, to_index):
        deck = self._writable(doc)
        idx = self._resolve_slide(deck, slide)
        to = max(0, min(to_index, len(deck.slides) - 1))
        deck.slides.insert(to, deck.slides.pop(idx))
        deck.dirty = True
        return {"moved_from": idx, "moved_to": to}

    def slide_delete(self, doc, slide, expect_hash=None):
        deck = self._writable(doc)
        idx = self._resolve_slide(deck, slide, expect_hash)
        deck.slides.pop(idx)
        deck.dirty = True
        return {"deleted": idx, "count": len(deck.slides)}

    def slide_thumbnail(self, doc, slide, width_px=960):
        raise DocdError(UNSUPPORTED_ON_BACKEND,
                        "The fake presentation backend cannot render thumbnails.")

    def pres_save_as(self, doc, path, format):
        deck = self._deck(doc)
        if format != "pptx":
            raise DocdError(SAVE_FORMAT_UNSUPPORTED,
                            "Fake presentation backend can only save 'pptx' (a stub).")
        with open(path, "w", encoding="utf-8") as f:
            for i, sl in enumerate(deck.slides):
                f.write(f"# slide {i}\n")
                for s in sl["shapes"]:
                    if s["text"]:
                        f.write(f"{s['placeholder']}: {s['text']}\n")
        deck.path = deck.path or path
        deck.dirty = False
        return {"saved": True, "path": path, "format": format}
