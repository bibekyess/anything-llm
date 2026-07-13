"""PptDriver — PowerPoint automation via COM (design doc §6.3).

Slides are canvases of shapes, so this driver exposes the slide_* tool subset
instead of the paragraph-oriented doc_* editing surface. Addressing (§6.2):
slides by 0-based index + content hash (±3 re-anchor), shapes by PowerPoint's
stable integer Shape.Id ("sh5").

Windows-only verification via smoke/ppt_smoke.py; the cross-platform test
suite covers the same contract through FakePresDriver.
"""

import base64
import os
import tempfile

from .. import addressing
from ..errors import (
    DocdError, BAD_PARAMS, NO_SUCH_DOC, READ_ONLY, SAVE_FORMAT_UNSUPPORTED,
    STALE_RANGE, UNSUPPORTED_ON_BACKEND,
)
from .base import BaseDriver
from .word import _com, com_retry

MSO_PLACEHOLDER = 14
PP_PLACEHOLDER_TITLE = 1
PP_PLACEHOLDER_CENTER_TITLE = 13
PP_PLACEHOLDER_BODY = 2

PRES_SAVE_FORMATS = {
    "pptx": 24,  # ppSaveAsOpenXMLPresentation
    "ppt": 1,    # ppSaveAsPresentation
    "pdf": 32,   # ppSaveAsPDF
    "odp": 35,   # ppSaveAsOpenDocumentPresentation (lossy)
}

SLIDE_REANCHOR_WINDOW = 3


def _shape_text(shape):
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return shape.TextFrame.TextRange.Text.replace("\r", "\n")
    except Exception:
        pass
    return ""


def _placeholder_kind(shape):
    try:
        if shape.Type == MSO_PLACEHOLDER:
            t = shape.PlaceholderFormat.Type
            if t in (PP_PLACEHOLDER_TITLE, PP_PLACEHOLDER_CENTER_TITLE):
                return "title"
            if t == PP_PLACEHOLDER_BODY:
                return "body"
            return f"placeholder:{t}"
    except Exception:
        pass
    return None


class PptDriver(BaseDriver):
    prefix = "p"
    backend = "powerpoint"

    def __init__(self):
        self._app = None
        self._pres = {}
        self._counter = 0

    # ── app / handle plumbing (mirrors WordDriver) ─────────────────────
    def _ensure_app(self):
        client = _com()
        if self._app is not None:
            try:
                _ = self._app.Name
                return self._app
            except Exception:
                self._app = None
                self._pres.clear()
        import pywintypes
        try:
            self._app = client.GetActiveObject("PowerPoint.Application")
        except pywintypes.com_error:
            self._app = client.Dispatch("PowerPoint.Application")
        # PowerPoint has no headless mode; Visible=True is both required and
        # exactly what we want — the user watches the deck change.
        self._app.Visible = True
        return self._app

    def _p(self, handle):
        if handle not in self._pres:
            raise DocdError(NO_SUCH_DOC, f"No open presentation with handle '{handle}'.")
        pres = self._pres[handle]
        try:
            _ = pres.Name
        except Exception:
            del self._pres[handle]
            raise DocdError(NO_SUCH_DOC, f"Presentation '{handle}' was closed in PowerPoint.")
        return pres

    def _writable(self, handle):
        pres = self._p(handle)
        if pres.ReadOnly:
            raise DocdError(READ_ONLY, f"Presentation '{handle}' is open read-only.")
        return pres

    def _register(self, pres):
        for handle, existing in self._pres.items():
            try:
                if existing.FullName == pres.FullName:
                    return handle
            except Exception:
                continue
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._pres[handle] = pres
        return handle

    # ── slide addressing ───────────────────────────────────────────────
    @staticmethod
    def _slide_hash(sl):
        parts = [_shape_text(sh) for sh in sl.Shapes]
        return addressing.para_hash("\n".join(p for p in parts if p))

    def _resolve_slide(self, pres, index, expect_hash=None):
        count = pres.Slides.Count
        if count == 0:
            raise DocdError(BAD_PARAMS, "Presentation has no slides.")
        index = max(0, min(index, count - 1))
        if expect_hash is None:
            return index
        if self._slide_hash(pres.Slides(index + 1)) == expect_hash:
            return index
        for delta in range(1, SLIDE_REANCHOR_WINDOW + 1):
            for cand in (index - delta, index + delta):
                if 0 <= cand < count and self._slide_hash(pres.Slides(cand + 1)) == expect_hash:
                    return cand
        raise DocdError(
            STALE_RANGE,
            f"Slide {index} no longer matches hash #{expect_hash} "
            f"(searched ±{SLIDE_REANCHOR_WINDOW}). Run slide_list again.",
        )

    @staticmethod
    def _find_shape(sl, shape_id):
        if not shape_id or not shape_id.startswith("sh"):
            raise DocdError(BAD_PARAMS, "shape must be an id from slide_read, e.g. 'sh5'.")
        target = int(shape_id[2:])
        for sh in sl.Shapes:
            if sh.Id == target:
                return sh
        raise DocdError(BAD_PARAMS, f"No shape '{shape_id}' on this slide (re-run slide_read).")

    def _goto(self, pres, index):
        try:
            pres.Application.ActiveWindow.View.GotoSlide(index + 1)
        except Exception:
            pass  # cosmetic

    # ── shared doc_* surface (open/list/save/close route here too) ─────
    @com_retry
    def list_open(self):
        app = self._ensure_app()
        out = []
        for i in range(1, app.Presentations.Count + 1):
            pres = app.Presentations(i)
            handle = self._register(pres)
            out.append({
                "doc": handle,
                "backend": self.backend,
                "path": pres.FullName,
                "dirty": not pres.Saved,
                "slides": pres.Slides.Count,
                "read_only": bool(pres.ReadOnly),
            })
        return {"docs": out}

    @com_retry
    def open(self, path, read_only=False):
        app = self._ensure_app()
        for i in range(1, app.Presentations.Count + 1):
            if app.Presentations(i).FullName.lower() == path.lower():
                handle = self._register(app.Presentations(i))
                pres = self._pres[handle]
                break
        else:
            pres = app.Presentations.Open(
                FileName=path, ReadOnly=read_only, Untitled=False, WithWindow=True
            )
            handle = self._register(pres)
        return {
            "doc": handle,
            "path": pres.FullName,
            "slides": pres.Slides.Count,
            "dirty": not pres.Saved,
            "read_only": bool(pres.ReadOnly),
            "hint": "This is a presentation: use slide_list/slide_read, not doc_read.",
        }

    @com_retry
    def new_doc(self):
        app = self._ensure_app()
        pres = app.Presentations.Add(WithWindow=True)
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._pres[handle] = pres
        return {"doc": handle, "path": None, "slides": pres.Slides.Count}

    def _unsupported(self, alt):
        raise DocdError(
            UNSUPPORTED_ON_BACKEND,
            f"Presentations are slide/shape-oriented — use {alt} instead.",
        )

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
        self._unsupported("slide_edit_text (layouts carry the styling)")

    def tables(self, doc, **kw):
        self._unsupported("slide_edit_text")

    def selection(self, doc, **kw):
        self._unsupported("slide_read")

    @com_retry
    def save(self, doc):
        pres = self._writable(doc)
        if not pres.Path:
            raise DocdError(
                BAD_PARAMS,
                "This presentation has never been saved; use pres_save_as with a path.",
            )
        pres.Save()
        return {"saved": True, "path": pres.FullName}

    def save_as(self, doc, path, format):
        return self.pres_save_as(doc, path, format)

    @com_retry
    def close(self, doc, discard_changes=False):
        pres = self._p(doc)
        was_dirty = not pres.Saved
        if not pres.Path and not discard_changes and was_dirty:
            raise DocdError(
                BAD_PARAMS,
                "Unsaved new presentation: pres_save_as it first, or close with "
                "discard_changes=true.",
            )
        if discard_changes:
            pres.Saved = True  # suppress the save prompt
        elif was_dirty and pres.Path:
            pres.Save()
        pres.Close()
        del self._pres[doc]
        return {"closed": True, "was_dirty": was_dirty}

    def shutdown(self):
        self._pres.clear()
        self._app = None

    # ── slide_* tools ──────────────────────────────────────────────────
    @com_retry
    def slide_list(self, doc):
        pres = self._p(doc)
        slides = []
        for i in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(i)
            title = ""
            try:
                if sl.Shapes.HasTitle:
                    title = _shape_text(sl.Shapes.Title).strip()
            except Exception:
                pass
            layout = ""
            try:
                layout = sl.CustomLayout.Name
            except Exception:
                pass
            has_notes = False
            try:
                has_notes = bool(
                    sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text.strip()
                )
            except Exception:
                pass
            slides.append({
                "slide": i - 1,
                "hash": self._slide_hash(sl),
                "layout": layout,
                "title": title,
                "notes": has_notes,
            })
        return {"count": pres.Slides.Count, "slides": slides}

    @com_retry
    def slide_read(self, doc, slide):
        pres = self._p(doc)
        idx = self._resolve_slide(pres, slide)
        sl = pres.Slides(idx + 1)
        shapes = []
        for sh in sl.Shapes:
            text = _shape_text(sh)
            if not text:
                continue
            shapes.append({
                "shape": f"sh{sh.Id}",
                "hash": addressing.para_hash(text),
                "placeholder": _placeholder_kind(sh),
                "text": text,
            })
        notes = ""
        try:
            notes = sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            pass
        return {"slide": idx, "slide_hash": self._slide_hash(sl), "shapes": shapes, "notes": notes}

    @com_retry
    def slide_add(self, doc, after_slide=None, layout=None, title=None,
                  body=None, duplicate_of=None):
        pres = self._writable(doc)
        count = pres.Slides.Count
        if duplicate_of is not None:
            src = pres.Slides(self._resolve_slide(pres, duplicate_of) + 1)
            new = src.Duplicate()(1)  # Duplicate returns a SlideRange
            if after_slide is not None:
                new.MoveTo(min(after_slide, count) + 2)
        else:
            at = (min(after_slide, count - 1) + 2) if after_slide is not None else count + 1
            custom = self._pick_layout(pres, layout)
            new = pres.Slides.AddSlide(at, custom)
        if title is not None:
            try:
                new.Shapes.Title.TextFrame.TextRange.Text = title
            except Exception:
                pass
        if body is not None:
            self._fill_body(new, body)
        idx = new.SlideIndex - 1
        self._goto(pres, idx)
        return {"slide": idx, "hash": self._slide_hash(new), "layout": layout or "", "count": pres.Slides.Count}

    def _pick_layout(self, pres, layout):
        layouts = pres.SlideMaster.CustomLayouts
        if layout:
            for j in range(1, layouts.Count + 1):
                if layouts(j).Name.lower() == layout.lower():
                    return layouts(j)
        # default: second layout is conventionally "Title and Content"
        return layouts(min(2, layouts.Count))

    @staticmethod
    def _fill_body(sl, body):
        for sh in sl.Shapes:
            if _placeholder_kind(sh) == "body" and sh.HasTextFrame:
                tr = sh.TextFrame.TextRange
                lines = body.split("\n")
                tr.Text = "\r".join(line.lstrip("\t") for line in lines)
                for k, line in enumerate(lines, start=1):
                    depth = len(line) - len(line.lstrip("\t"))
                    if depth:
                        try:
                            tr.Paragraphs(k).ParagraphFormat.IndentLevel = min(depth + 1, 5)
                        except Exception:
                            pass
                return
        # No body placeholder on this layout — leave silently; caller sees text absent.

    @com_retry
    def slide_edit_text(self, doc, slide, shape, text, expect_hash=None):
        pres = self._writable(doc)
        idx = self._resolve_slide(pres, slide)
        sl = pres.Slides(idx + 1)
        sh = self._find_shape(sl, shape)
        if expect_hash is not None:
            current = addressing.para_hash(_shape_text(sh))
            if current != expect_hash:
                raise DocdError(
                    STALE_RANGE,
                    f"Shape {shape} changed since slide_read (now #{current}). Re-read the slide.",
                )
        if not sh.HasTextFrame:
            raise DocdError(BAD_PARAMS, f"Shape {shape} has no text frame.")
        tr = sh.TextFrame.TextRange
        lines = text.split("\n")
        tr.Text = "\r".join(line.lstrip("\t") for line in lines)
        for k, line in enumerate(lines, start=1):
            depth = len(line) - len(line.lstrip("\t"))
            if depth:
                try:
                    tr.Paragraphs(k).ParagraphFormat.IndentLevel = min(depth + 1, 5)
                except Exception:
                    pass
        self._goto(pres, idx)
        return {"slide": idx, "shape": shape, "hash": addressing.para_hash(_shape_text(sh))}

    @com_retry
    def slide_notes_edit(self, doc, slide, text):
        pres = self._writable(doc)
        idx = self._resolve_slide(pres, slide)
        sl = pres.Slides(idx + 1)
        try:
            sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        except Exception:
            raise DocdError(
                BAD_PARAMS,
                "Could not access the notes placeholder (customized notes master?).",
            )
        return {"slide": idx, "notes_chars": len(text)}

    @com_retry
    def slide_reorder(self, doc, slide, to_index):
        pres = self._writable(doc)
        idx = self._resolve_slide(pres, slide)
        count = pres.Slides.Count
        to = max(0, min(to_index, count - 1))
        pres.Slides(idx + 1).MoveTo(to + 1)
        self._goto(pres, to)
        return {"moved_from": idx, "moved_to": to}

    @com_retry
    def slide_delete(self, doc, slide, expect_hash=None):
        pres = self._writable(doc)
        idx = self._resolve_slide(pres, slide, expect_hash)
        pres.Slides(idx + 1).Delete()
        return {"deleted": idx, "count": pres.Slides.Count}

    @com_retry
    def slide_thumbnail(self, doc, slide, width_px=960):
        pres = self._p(doc)
        idx = self._resolve_slide(pres, slide)
        sl = pres.Slides(idx + 1)
        height = int(width_px * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
        out = os.path.join(tempfile.gettempdir(), f"docd_slide_{doc}_{idx}.png")
        sl.Export(out, "PNG", width_px, height)
        with open(out, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return {"slide": idx, "path": out, "png_base64": b64}

    @com_retry
    def pres_save_as(self, doc, path, format):
        pres = self._p(doc)
        if format == "png":
            os.makedirs(path, exist_ok=True)
            pres.Export(path, "PNG")
            return {"saved": True, "path": path, "format": "png",
                    "slides": pres.Slides.Count}
        if format not in PRES_SAVE_FORMATS:
            raise DocdError(
                SAVE_FORMAT_UNSUPPORTED,
                f"PowerPoint cannot save '{format}'. Supported: "
                f"{', '.join(PRES_SAVE_FORMATS)}, png.",
            )
        pres.SaveAs(path, PRES_SAVE_FORMATS[format])
        return {"saved": True, "path": path, "format": format}
