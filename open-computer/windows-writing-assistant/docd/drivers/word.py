"""WordDriver — MS Word automation via COM (design doc §3).

Range-based, never Selection-based: Range edits don't steal the user's caret
and don't depend on window focus, but still repaint live in the open window.

Requires pywin32; import is deferred so the module can be imported (and the
rest of the package tested) on non-Windows hosts.

Windows-only verification note: this driver is exercised by smoke/word_smoke.py
on a machine with Word installed; the shared logic (addressing/render/RPC) is
covered by the cross-platform test suite against FakeDriver.
"""

import re
import time

from .. import addressing, render
from ..errors import (
    DocdError, APP_BUSY_MODAL, APP_NOT_RUNNING, COM_ERROR, NO_SUCH_DOC,
    READ_ONLY, SAVE_FORMAT_UNSUPPORTED,
)
from .base import BaseDriver

# Word enum constants (avoid win32com.client.constants: it requires makepy).
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_FIND_STOP = 0          # Wrap: never escape the scoped range
WD_REPLACE_NONE = 0
WD_REPLACE_ONE = 1
WD_REPLACE_ALL = 2
WD_OUTLINE_BODY = 10
WD_WITHIN_TABLE = 12      # Range.Information(wdWithInTable)
WD_ALERTS_NONE = 0
WD_SELECTION_IP = 1       # wdSelectionIP: collapsed caret, nothing selected

SAVE_FORMATS = {
    "docx": 12,  # wdFormatXMLDocument
    "doc": 0,    # wdFormatDocument97
    "pdf": 17,   # wdFormatPDF
    "odt": 23,   # wdFormatOpenDocumentText
    "txt": 2,    # wdFormatText
}

# Modal dialog up -> COM calls bounce with these HRESULTs (design doc §3).
RPC_E_CALL_REJECTED = -2147418111       # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
_RETRYABLE = (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER)
_RETRIES = 5
_BACKOFF_S = 0.3


def _com():
    try:
        import pythoncom  # noqa: F401
        import win32com.client
        return win32com.client
    except ImportError as e:
        raise DocdError(
            APP_NOT_RUNNING,
            "pywin32 is not installed — WordDriver only runs on Windows "
            "(pip install pywin32). Use --backend fake elsewhere.",
        ) from e


def com_retry(fn):
    """Retry COM calls rejected while Word shows a modal dialog."""
    def wrapper(*args, **kwargs):
        import pywintypes
        last = None
        for _ in range(_RETRIES):
            try:
                return fn(*args, **kwargs)
            except pywintypes.com_error as e:
                last = e
                if e.hresult in _RETRYABLE:
                    time.sleep(_BACKOFF_S)
                    continue
                raise DocdError(
                    COM_ERROR,
                    f"Word COM call failed: {e.strerror or e}",
                    data={"hresult": e.hresult},
                ) from e
        raise DocdError(
            APP_BUSY_MODAL,
            "Word is blocked by a modal dialog (COM calls rejected 5x). "
            "Dismiss the dialog in the Word window, or use the UIA layer.",
            data={"hresult": last.hresult if last else None},
        ) from last
    return wrapper


class WordDriver(BaseDriver):
    prefix = "w"
    backend = "word"

    def __init__(self):
        self._app = None
        self._docs = {}     # handle -> COM Document
        self._counter = 0

    # ── app / handle plumbing ──────────────────────────────────────────
    def _ensure_app(self):
        client = _com()
        if self._app is not None:
            try:
                _ = self._app.Visible  # liveness probe
                return self._app
            except Exception:
                self._app = None
                self._docs.clear()
        import pywintypes
        try:
            self._app = client.GetActiveObject("Word.Application")
        except pywintypes.com_error:
            self._app = client.DispatchEx("Word.Application")
        self._app.Visible = True  # never set False: the user is watching
        self._app.DisplayAlerts = WD_ALERTS_NONE
        return self._app

    def _doc(self, handle):
        if handle not in self._docs:
            raise DocdError(NO_SUCH_DOC, f"No open document with handle '{handle}'.")
        doc = self._docs[handle]
        try:
            _ = doc.Name  # dropped COM proxy (doc closed by user)?
        except Exception:
            del self._docs[handle]
            raise DocdError(NO_SUCH_DOC, f"Document '{handle}' was closed in Word.")
        return doc

    def _writable(self, handle):
        doc = self._doc(handle)
        if doc.ReadOnly:
            raise DocdError(READ_ONLY, f"Document '{handle}' is open read-only.")
        return doc

    def _register(self, com_doc):
        for handle, existing in self._docs.items():
            if existing.FullName == com_doc.FullName:
                return handle
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._docs[handle] = com_doc
        return handle

    # ── paragraph access (0-based API over 1-based COM collections) ───
    @staticmethod
    def _para_text(doc, i):
        return doc.Paragraphs(i + 1).Range.Text

    def _text_at(self, doc):
        return lambda i: self._para_text(doc, i)

    def _hashes(self, doc, indices):
        return [
            [i, addressing.para_hash(self._para_text(doc, i))] for i in indices
        ]

    def _scroll_to(self, rng):
        try:
            self._app.ActiveWindow.ScrollIntoView(rng)
        except Exception:
            pass  # cosmetic: never fail an edit over scrolling

    def _global_index(self, doc, com_para):
        """0-based document index of a COM Paragraph (Range-start trick, §3)."""
        start = com_para.Range.Start
        return doc.Range(0, start).Paragraphs.Count if start > 0 else 0

    # ── BaseDriver ─────────────────────────────────────────────────────
    @com_retry
    def list_open(self):
        app = self._ensure_app()
        docs = []
        for i in range(1, app.Documents.Count + 1):
            com_doc = app.Documents(i)
            handle = self._register(com_doc)
            docs.append({
                "doc": handle,
                "backend": self.backend,
                "path": com_doc.FullName,
                "dirty": not com_doc.Saved,
                "paragraphs": com_doc.Paragraphs.Count,
                "read_only": bool(com_doc.ReadOnly),
                "track_changes": bool(com_doc.TrackRevisions),
            })
        return {"docs": docs}

    @com_retry
    def open(self, path, read_only=False):
        app = self._ensure_app()
        for i in range(1, app.Documents.Count + 1):  # attach if already open
            if app.Documents(i).FullName.lower() == path.lower():
                handle = self._register(app.Documents(i))
                com_doc = self._docs[handle]
                break
        else:
            com_doc = app.Documents.Open(
                FileName=path,
                ReadOnly=read_only,
                AddToRecentFiles=False,
                ConfirmConversions=False,
            )
            handle = self._register(com_doc)
        return {
            "doc": handle,
            "path": com_doc.FullName,
            "paragraphs": com_doc.Paragraphs.Count,
            "dirty": not com_doc.Saved,
            "read_only": bool(com_doc.ReadOnly),
        }

    @com_retry
    def new_doc(self):
        app = self._ensure_app()
        com_doc = app.Documents.Add()
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._docs[handle] = com_doc
        return {"doc": handle, "path": None, "paragraphs": com_doc.Paragraphs.Count}

    @com_retry
    def selection(self, doc):
        d = self._doc(doc)
        app = self._ensure_app()
        sel = app.Selection
        # The selection must belong to this document's window.
        if sel.Document.FullName != d.FullName or sel.Type == WD_SELECTION_IP:
            return {"collapsed": True, "text": "", "from_para": None,
                    "to_para": None, "hashes": []}
        rng = sel.Range
        first = self._global_index(d, rng.Paragraphs(1))
        last = first + rng.Paragraphs.Count - 1
        hashes = [addressing.para_hash(self._para_text(d, i))
                  for i in range(first, last + 1)]
        return {
            "collapsed": False,
            "text": rng.Text.rstrip("\r"),
            "from_para": first,
            "to_para": last,
            "hashes": hashes,
        }

    @com_retry
    def tables(self, doc, op, table=None, cell=None, value=None, values=None,
               at=None, para=None, expect_hash=None, replace_range=None,
               header_row=False):
        d = self._doc(doc)
        if op == "list":
            out = []
            for i in range(1, d.Tables.Count + 1):
                t = d.Tables(i)
                out.append({
                    "table": f"t{i - 1}",
                    "rows": t.Rows.Count,
                    "cols": t.Columns.Count,
                    "at_para": self._global_index(d, t.Range.Paragraphs(1)),
                })
            return {"tables": out}
        if op == "read":
            t = self._com_table(d, table)
            rows = []
            for r in range(1, t.Rows.Count + 1):
                row = []
                for c in range(1, t.Columns.Count + 1):
                    try:
                        row.append(t.Cell(r, c).Range.Text.rstrip("\r\x07"))
                    except Exception:
                        row.append("")  # merged/missing cell
                rows.append(row)
            lines = ["| " + " | ".join(r) + " |" for r in rows]
            if len(lines) > 1:
                lines.insert(1, "|" + "---|" * (t.Columns.Count))
            return {"table": table, "rows": t.Rows.Count,
                    "cols": t.Columns.Count, "text": "\n".join(lines)}
        if op == "write":
            self._writable(doc)
            t = self._com_table(d, table)
            if cell is None:
                raise DocdError("BAD_PARAMS", "tables write requires `cell`.")
            block = values if values is not None else [[value]]
            written = 0
            for r, row in enumerate(block):
                for c, val in enumerate(row):
                    rr, cc = cell["row"] + 1 + r, cell["col"] + 1 + c
                    if rr <= t.Rows.Count and cc <= t.Columns.Count:
                        t.Cell(rr, cc).Range.Text = val
                        written += 1
            self._scroll_to(t.Range)
            return {"written": written}
        if op == "create":
            self._writable(doc)
            if not values or not values[0]:
                raise DocdError("BAD_PARAMS", "tables create requires non-empty `values`.")
            return self._create_table(
                d, values, at, para, expect_hash, replace_range, header_row
            )
        raise DocdError("BAD_PARAMS", f"Unknown tables op '{op}'.")

    def _com_table(self, d, table):
        if not table or not table.startswith("t"):
            raise DocdError("BAD_PARAMS", "tables read/write requires `table` (e.g. 't0').")
        idx = int(table[1:]) + 1
        if idx > d.Tables.Count:
            raise DocdError("BAD_PARAMS", f"No table '{table}' in this document.")
        return d.Tables(idx)

    def _create_table(self, d, values, at, para, expect_hash, replace_range,
                      header_row):
        count = d.Paragraphs.Count
        deleted = 0
        if replace_range:
            # Convert-selection-to-table: verify hashes, delete the source
            # paragraphs, then build the table where they were.
            lo, hi = replace_range["from_para"], replace_range["to_para"]
            addressing.check_range_hashes(
                self._text_at(d), count, lo, hi, replace_range["expect_hashes"]
            )
            rng = d.Range(
                d.Paragraphs(lo + 1).Range.Start, d.Paragraphs(hi + 1).Range.End
            )
            if rng.End >= d.Content.End:
                rng = d.Range(rng.Start, d.Content.End - 1)
            rng.Text = ""
            deleted = hi - lo + 1
            anchor = d.Range(rng.Start, rng.Start)
        elif at in ("before_para", "after_para"):
            if para is None:
                raise DocdError("BAD_PARAMS", f"'{at}' requires `para`.")
            idx, _ = addressing.resolve_anchor(
                self._text_at(d), count, para, expect_hash
            )
            p_rng = d.Paragraphs(idx + 1).Range
            pos = p_rng.Start if at == "before_para" else p_rng.End
            anchor = d.Range(pos, pos)
        else:  # end of document
            end = d.Content.End - 1
            anchor = d.Range(end, end)

        table = d.Tables.Add(anchor, len(values), len(values[0]))
        table.Borders.Enable = True  # locale-independent (style names are localized)
        for r, row in enumerate(values, start=1):
            for c, val in enumerate(row, start=1):
                table.Cell(r, c).Range.Text = str(val)
        if header_row:
            table.Rows(1).Range.Bold = True
            table.Rows(1).HeadingFormat = True
        self._scroll_to(table.Range)
        # Table ids are collection positions; find where this one landed.
        t_index = next(
            (i for i in range(1, d.Tables.Count + 1)
             if d.Tables(i).Range.Start == table.Range.Start),
            d.Tables.Count,
        )
        return {
            "table": f"t{t_index - 1}",
            "rows": len(values),
            "cols": len(values[0]),
            "at_para": self._global_index(d, table.Range.Paragraphs(1)),
            "deleted_paras": deleted,
        }

    @com_retry
    def read(self, doc, from_para=None, to_para=None, max_chars=None):
        d = self._doc(doc)
        count = d.Paragraphs.Count
        start = from_para or 0
        end = min(to_para if to_para is not None else count - 1, count - 1)
        paras = []
        for i in range(start, end + 1):
            p = d.Paragraphs(i + 1)
            rng = p.Range
            paras.append({
                "text": rng.Text,
                "style": str(p.Style.NameLocal),
                "outline_level": None if p.OutlineLevel == WD_OUTLINE_BODY else int(p.OutlineLevel),
                "in_table": bool(rng.Information(WD_WITHIN_TABLE)),
            })
        text, _, rev = render.render_read(paras, from_para=start, max_chars=max_chars)
        return {"text": text, "from_para": start, "to_para": end, "count": count, "rev": rev}

    @com_retry
    def outline(self, doc):
        d = self._doc(doc)
        headings = []
        for i in range(d.Paragraphs.Count):
            p = d.Paragraphs(i + 1)
            level = int(p.OutlineLevel)
            if level != WD_OUTLINE_BODY:
                headings.append({"level": level, "para": i, "text": p.Range.Text})
        return {"text": render.render_outline(headings)}

    @com_retry
    def insert(self, doc, text, where, para=None, expect_hash=None,
               bookmark=None, style_map=True):
        d = self._writable(doc)
        count = d.Paragraphs.Count
        moved = False
        restore_bookmark = None

        if where == "end":
            anchor = d.Content
            anchor.Collapse(WD_COLLAPSE_END)
        elif where == "cursor":
            anchor = self._app.Selection.Range  # read position only
            anchor.Collapse(WD_COLLAPSE_END)
        elif where in ("before_para", "after_para"):
            if para is None:
                raise DocdError("BAD_PARAMS", f"'{where}' requires `para`.")
            idx, moved = addressing.resolve_anchor(
                self._text_at(d), count, para, expect_hash
            )
            anchor = d.Paragraphs(idx + 1).Range
            anchor.Collapse(
                WD_COLLAPSE_START if where == "before_para" else WD_COLLAPSE_END
            )
        elif where == "bookmark":
            if not bookmark or not d.Bookmarks.Exists(bookmark):
                raise DocdError("BAD_PARAMS", f"Bookmark '{bookmark}' not found.")
            anchor = d.Bookmarks(bookmark).Range
            anchor.Collapse(WD_COLLAPSE_END)
            restore_bookmark = bookmark  # inserting at a bookmark can delete it
        else:
            raise DocdError("BAD_PARAMS", f"Unknown insert anchor '{where}'.")

        pieces = render.parse_styled_text(text, style_map)
        payload = "\r".join(t for t, _ in pieces)
        # A standalone block needs a trailing paragraph mark unless appending
        # inline at the caret.
        if where != "cursor":
            payload += "\r"
        ins_start = anchor.Start
        anchor.InsertAfter(payload)  # extends `anchor` over the inserted text
        ins_rng = d.Range(ins_start, anchor.End)

        for k, (_, style) in enumerate(pieces, start=1):
            if style:
                try:
                    ins_rng.Paragraphs(k).Style = d.Styles(style)
                except Exception:
                    pass  # style not in template; leave as body text

        if restore_bookmark and not d.Bookmarks.Exists(restore_bookmark):
            d.Bookmarks.Add(restore_bookmark, d.Range(ins_start, ins_start))

        first = self._global_index(d, ins_rng.Paragraphs(1))
        self._scroll_to(ins_rng)
        return {
            "inserted": len(pieces),
            "first_para": first,
            "affected": self._hashes(d, range(first, first + len(pieces))),
            "moved": moved,
        }

    @com_retry
    def replace(self, doc, find, replace, regex=False, match_case=False,
                occurrence="all", scope=None):
        d = self._writable(doc)
        count = d.Paragraphs.Count
        if scope:
            lo = max(0, scope["from_para"])
            hi = min(scope["to_para"], count - 1)
            rng = d.Range(
                d.Paragraphs(lo + 1).Range.Start, d.Paragraphs(hi + 1).Range.End
            )
        else:
            rng = d.Content

        if regex:
            replaced, touched = self._replace_regex(
                d, rng, find, replace, match_case, occurrence
            )
        else:
            replaced, touched = self._replace_native(
                d, rng, find, replace, match_case, occurrence
            )
        if touched:
            self._scroll_to(d.Paragraphs(touched[-1] + 1).Range)
        return {"replaced": replaced, "affected": self._hashes(d, touched)}

    def _replace_native(self, d, rng, find, replace, match_case, occurrence):
        """Range.Find with Wrap=wdFindStop so we never escape the scope."""
        touched = set()
        if occurrence == "all":
            # Loop one-at-a-time (not wdReplaceAll) so we can record positions.
            replaced = 0
            search = d.Range(rng.Start, rng.End)
            while True:
                f = search.Find
                f.ClearFormatting()
                found = f.Execute(
                    FindText=find, MatchCase=match_case, MatchWholeWord=False,
                    MatchWildcards=False, Forward=True, Wrap=WD_FIND_STOP,
                )
                if not found or search.Start >= rng.End:
                    break
                touched.add(self._global_index(d, search.Paragraphs(1)))
                search.Text = replace
                replaced += 1
                search = d.Range(search.End, rng.End)
                if search.Start >= search.End:
                    break
            return replaced, sorted(touched)

        target = 1 if occurrence == "first" else int(occurrence)
        seen = 0
        search = d.Range(rng.Start, rng.End)
        while True:
            f = search.Find
            f.ClearFormatting()
            found = f.Execute(
                FindText=find, MatchCase=match_case, MatchWholeWord=False,
                MatchWildcards=False, Forward=True, Wrap=WD_FIND_STOP,
            )
            if not found:
                return 0, []
            seen += 1
            if seen == target:
                idx = self._global_index(d, search.Paragraphs(1))
                search.Text = replace
                return 1, [idx]
            search = d.Range(search.End, rng.End)

    def _replace_regex(self, d, rng, find, replace, match_case, occurrence):
        """Python-re emulation over paragraph text (design doc §3: full regex
        is emulated; Word wildcards only cover a subset)."""
        flags = 0 if match_case else re.IGNORECASE
        pattern = re.compile(find, flags)
        first = self._global_index(d, rng.Paragraphs(1))
        n_paras = rng.Paragraphs.Count
        limit = 0 if occurrence == "all" else (1 if occurrence == "first" else int(occurrence))
        seen = replaced = 0
        touched = []
        for k in range(n_paras):
            i = first + k
            p_rng = d.Paragraphs(i + 1).Range
            text = p_rng.Text
            body, tail = (text[:-1], text[-1]) if text.endswith("\r") else (text, "")
            if limit == 0:
                new, n = pattern.subn(replace, body)
                replaced += n
            else:
                new, n = body, 0
                for match in pattern.finditer(body):
                    seen += 1
                    if seen == limit:
                        new = body[: match.start()] + replace + body[match.end():]
                        n = 1
                        replaced += 1
                        break
            if n:
                # Rewrite paragraph body, excluding its paragraph mark.
                d.Range(p_rng.Start, p_rng.End - (1 if tail else 0)).Text = new
                touched.append(i)
            if limit and replaced:
                break
        return replaced, touched

    @com_retry
    def edit_range(self, doc, from_para, to_para, expect_hashes, new_text):
        d = self._writable(doc)
        count = d.Paragraphs.Count
        addressing.check_range_hashes(
            self._text_at(d), count, from_para, to_para, expect_hashes
        )
        rng = d.Range(
            d.Paragraphs(from_para + 1).Range.Start,
            d.Paragraphs(to_para + 1).Range.End,
        )
        # Keep the final paragraph mark of the doc intact.
        if rng.End >= d.Content.End:
            rng = d.Range(rng.Start, d.Content.End - 1)
        pieces = render.parse_styled_text(new_text, True) if new_text else []
        if not pieces:
            rng.Text = ""
            return {"replaced": 0, "deleted": to_para - from_para + 1, "affected": []}
        rng.Text = "\r".join(t for t, _ in pieces) + "\r"
        new_rng = d.Range(rng.Start, rng.End)
        for k, (_, style) in enumerate(pieces, start=1):
            if style:
                try:
                    new_rng.Paragraphs(k).Style = d.Styles(style)
                except Exception:
                    pass
        self._scroll_to(new_rng)
        return {
            "replaced": len(pieces),
            "deleted": 0,
            "affected": self._hashes(d, range(from_para, from_para + len(pieces))),
        }

    @com_retry
    def apply_style(self, doc, from_para, to_para=None, style=None):
        d = self._writable(doc)
        end = to_para if to_para is not None else from_para
        if from_para < 0 or end >= d.Paragraphs.Count:
            raise DocdError("BAD_PARAMS", "Paragraph range out of bounds.")
        try:
            style_obj = d.Styles(style)
        except Exception:
            raise DocdError("BAD_PARAMS", f"Style '{style}' not found in this document.")
        for i in range(from_para, end + 1):
            d.Paragraphs(i + 1).Style = style_obj
        self._scroll_to(d.Paragraphs(from_para + 1).Range)
        return {
            "styled": end - from_para + 1,
            "affected": self._hashes(d, range(from_para, end + 1)),
        }

    @com_retry
    def save(self, doc):
        d = self._writable(doc)
        if not d.Path:  # never saved: Save() would pop a modal SaveAs dialog
            raise DocdError(
                "BAD_PARAMS",
                "This document has never been saved; use doc_save_as with a path.",
            )
        d.Save()
        return {"saved": True, "path": d.FullName}

    @com_retry
    def save_as(self, doc, path, format):
        d = self._doc(doc)
        if format not in SAVE_FORMATS:
            raise DocdError(
                SAVE_FORMAT_UNSUPPORTED,
                f"Word cannot save '{format}'. Supported: {', '.join(SAVE_FORMATS)}. "
                "hwp/hwpx require the Hancom backend (route: save docx, open in HWP).",
            )
        d.SaveAs2(FileName=path, FileFormat=SAVE_FORMATS[format])
        return {"saved": True, "path": path, "format": format}

    @com_retry
    def close(self, doc, discard_changes=False):
        d = self._doc(doc)
        was_dirty = not d.Saved
        if not d.Path and not discard_changes and was_dirty:
            raise DocdError(
                "BAD_PARAMS",
                "Unsaved new document: doc_save_as it first, or close with "
                "discard_changes=true.",
            )
        d.Close(SaveChanges=0 if discard_changes else -1)  # wdDoNotSaveChanges / wdSaveChanges
        del self._docs[doc]
        return {"closed": True, "was_dirty": was_dirty}

    def shutdown(self):
        # Release proxies; never quit Word — the user may own other documents.
        self._docs.clear()
        self._app = None
