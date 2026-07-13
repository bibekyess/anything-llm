"""HwpDriver — Hancom Office (HWP/HWPX) automation via COM (design doc §4).

HWP's automation model is ACTION-based (HAction + HParameterSet at the cursor
position), not an object graph like Word — so position-based editing is
inherently flakier. Driver policy, per the design doc:
  - doc_replace (find-anchored) and end/cursor inserts are the reliable ops
  - index-based edit_range re-scans and hash-checks before touching anything
  - fields (누름틀) are the durable anchor: doc_insert where="bookmark"
    resolves HWP *fields* by name

UNVERIFIED-API WARNING: several call signatures below are marked
TODO(verify) — they follow Hancom's automation reference and pyhwpx usage
patterns, but this driver has not yet run against a licensed Hancom install.
Validate with smoke/hwp_smoke.py on a machine with Hancom Office and report
failures; expect to adjust constants/format strings per Hancom version.
"""

import os

from .. import addressing, render
from ..errors import (
    DocdError, APP_NOT_RUNNING, BAD_PARAMS, NO_SUCH_DOC, READ_ONLY,
    SAVE_FORMAT_UNSUPPORTED, UNSUPPORTED_ON_BACKEND,
)
from .base import BaseDriver
from .word import _com, com_retry

# hwp.MovePos moveID constants. TODO(verify): full table from the automation chm.
MOVE_TOP_OF_FILE = 2
MOVE_BOTTOM_OF_FILE = 3

SAVE_FORMATS = {
    "hwp": "HWP",
    "hwpx": "HWPX",   # TODO(verify): pre-2020 builds may need "HWPML2X"
    "pdf": "PDF",
    "txt": "TEXT",
}


class HwpDriver(BaseDriver):
    prefix = "h"
    backend = "hwp"

    def __init__(self):
        self._hwp = None
        self._docs = {}      # handle -> file path (HWP control is single-doc per window)
        self._counter = 0
        self._scan_cache = {}  # handle -> [paragraph texts] from the last read

    # ── app plumbing ───────────────────────────────────────────────────
    def _ensure_app(self):
        client = _com()
        if self._hwp is not None:
            try:
                _ = self._hwp.Version  # liveness probe; TODO(verify) property name
                return self._hwp
            except Exception:
                self._hwp = None
        try:
            self._hwp = client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception as e:
            raise DocdError(
                APP_NOT_RUNNING,
                f"Could not start Hancom Office automation ({e}). Is Hancom Office "
                "installed? (HWPFrame.HwpObject must be COM-registered.)",
            ) from e
        self._register_security_module()
        try:
            self._hwp.XHwpWindows.Item(0).Visible = True  # server starts hidden
        except Exception:
            pass
        return self._hwp

    def _register_security_module(self):
        """Suppress the per-file security popup. Needs a registry value under
        HKCU\\Software\\HNC\\HwpAutomation\\Modules (name=SecurityModule,
        data=path to Hancom's FilePathChecker DLL). We self-heal the key when
        HWP_SECURITY_DLL points at the DLL; otherwise we still try
        RegisterModule and fall back to letting the popup appear."""
        dll = os.environ.get("HWP_SECURITY_DLL", "")
        if dll and os.path.exists(dll):
            try:
                import winreg
                for branch in (r"Software\HNC\HwpAutomation\Modules",
                               r"Software\HNC\HwpCtrl\Modules"):  # older builds
                    try:
                        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, branch)
                        winreg.SetValueEx(key, "SecurityModule", 0, winreg.REG_SZ, dll)
                        winreg.CloseKey(key)
                    except OSError:
                        continue
            except Exception:
                pass
        try:
            self._hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        except Exception:
            pass  # not fatal: user may see the access-confirmation dialog

    def _doc(self, handle):
        if handle not in self._docs:
            raise DocdError(NO_SUCH_DOC, f"No open HWP document with handle '{handle}'.")
        return self._ensure_app()

    # ── text scan (doc_read backbone) ──────────────────────────────────
    def _scan_paragraphs(self, hwp):
        """InitScan/GetText loop -> list of paragraph texts (design doc §4).
        TODO(verify): scan option/Range flags across Hancom versions."""
        paras = []
        hwp.InitScan(0, 0x0077)
        try:
            while True:
                state, text = hwp.GetText()
                if state in (0, 1):  # 0=end of document, 1=error/empty
                    break
                paras.append(text.rstrip("\r\n"))
        finally:
            hwp.ReleaseScan()
        return paras or [""]

    def _refresh_cache(self, handle):
        hwp = self._doc(handle)
        self._scan_cache[handle] = self._scan_paragraphs(hwp)
        return self._scan_cache[handle]

    # ── BaseDriver ─────────────────────────────────────────────────────
    @com_retry
    def list_open(self):
        docs = []
        for handle, path in self._docs.items():
            docs.append({
                "doc": handle,
                "backend": self.backend,
                "path": path,
                "dirty": None,  # TODO(verify): IsModified property name
                "paragraphs": len(self._scan_cache.get(handle, [])) or None,
                "read_only": False,
            })
        return {"docs": docs}

    @com_retry
    def open(self, path, read_only=False):
        hwp = self._ensure_app()
        ok = hwp.Open(path, "HWP", "lock:false" if not read_only else "lock:true")
        if not ok:
            raise DocdError(BAD_PARAMS, f"Hancom Office could not open '{path}'.")
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._docs[handle] = path
        paras = self._refresh_cache(handle)
        return {
            "doc": handle,
            "path": path,
            "paragraphs": len(paras),
            "dirty": False,
            "read_only": read_only,
            "hint": "HWP editing is most reliable via doc_replace (find-anchored) "
                    "and doc_insert at end/bookmark(=HWP field). Index-based edits "
                    "may report STALE_RANGE more often than Word.",
        }

    @com_retry
    def new_doc(self):
        hwp = self._ensure_app()
        hwp.HAction.Run("FileNew")  # TODO(verify): may open a new window/tab
        self._counter += 1
        handle = f"{self.prefix}{self._counter}"
        self._docs[handle] = None
        self._scan_cache[handle] = [""]
        return {"doc": handle, "path": None, "paragraphs": 1}

    @com_retry
    def read(self, doc, from_para=None, to_para=None, max_chars=None):
        self._doc(doc)
        paras = self._refresh_cache(doc)
        count = len(paras)
        start = from_para or 0
        end = min(to_para if to_para is not None else count - 1, count - 1)
        text, _, rev = render.render_read(
            [{"text": t, "style": None, "outline_level": None} for t in paras[start:end + 1]],
            from_para=start, max_chars=max_chars,
        )
        return {"text": text, "from_para": start, "to_para": end, "count": count, "rev": rev}

    def outline(self, doc):
        # Outline levels are not part of the scan output (design doc §4);
        # style-less outline is accepted for the first HWP slice.
        raise DocdError(
            UNSUPPORTED_ON_BACKEND,
            "Outline is not available for HWP yet — use doc_read.",
        )

    @com_retry
    def insert(self, doc, text, where, para=None, expect_hash=None,
               bookmark=None, style_map=True):
        hwp = self._doc(doc)
        if where == "bookmark":
            # HWP fields are the bookmark namespace (the reliable anchor).
            if not bookmark:
                raise DocdError(BAD_PARAMS, "'bookmark' requires `bookmark` (an HWP field name).")
            plain = "\n".join(p["text"] for p in render.parse_markdown(text, style_map))
            hwp.PutFieldText(bookmark, plain)
            self._refresh_cache(doc)
            return {"inserted": 1, "first_para": None, "affected": [],
                    "moved": False, "field": bookmark}
        if where == "end":
            hwp.MovePos(MOVE_BOTTOM_OF_FILE, 0, 0)
        elif where == "cursor":
            pass  # insert at the user's caret
        else:
            raise DocdError(
                UNSUPPORTED_ON_BACKEND,
                "HWP inserts support where=end/cursor/bookmark (field). "
                "Paragraph-index anchors are not reliable on HWP.",
            )
        pieces = render.parse_markdown(text, style_map)
        for k, piece in enumerate(pieces):
            if k > 0 or where == "end":
                hwp.HAction.Run("BreakPara")
            self._insert_text(hwp, piece["text"])
        paras = self._refresh_cache(doc)
        first = max(0, len(paras) - len(pieces))
        return {
            "inserted": len(pieces),
            "first_para": first,
            "affected": [[i, addressing.para_hash(paras[i])]
                         for i in range(first, len(paras))],
            "moved": False,
            "note": "Markdown styling is not applied on HWP yet (plain text insert).",
        }

    @staticmethod
    def _insert_text(hwp, text):
        pset = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", pset.HSet)
        pset.Text = text
        hwp.HAction.Execute("InsertText", pset.HSet)

    @com_retry
    def replace(self, doc, find, replace, regex=False, match_case=False,
                occurrence="all", scope=None):
        hwp = self._doc(doc)
        if regex:
            raise DocdError(
                UNSUPPORTED_ON_BACKEND,
                "regex replace is not supported on HWP yet — use a literal find.",
            )
        if scope:
            raise DocdError(
                UNSUPPORTED_ON_BACKEND,
                "Paragraph-scoped replace is not supported on HWP — replace runs "
                "document-wide.",
            )
        before = self._scan_cache.get(doc) or self._refresh_cache(doc)
        if occurrence == "all":
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("AllReplace", pset.HSet)
            pset.FindString = find
            pset.ReplaceString = replace
            pset.IgnoreMessage = 1
            pset.MatchCase = 1 if match_case else 0
            hwp.HAction.Execute("AllReplace", pset.HSet)
        else:
            target = 1 if occurrence == "first" else int(occurrence)
            hwp.MovePos(MOVE_TOP_OF_FILE, 0, 0)
            for _ in range(target):
                pset = hwp.HParameterSet.HFindReplace
                hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                pset.FindString = find
                pset.IgnoreMessage = 1
                pset.MatchCase = 1 if match_case else 0
                found = hwp.HAction.Execute("RepeatFind", pset.HSet)
                if not found:
                    return {"replaced": 0, "affected": []}
            pset = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault("ExecReplace", pset.HSet)
            pset.FindString = find
            pset.ReplaceString = replace
            pset.IgnoreMessage = 1
            hwp.HAction.Execute("ExecReplace", pset.HSet)
        after = self._refresh_cache(doc)
        touched = [
            [i, addressing.para_hash(after[i])]
            for i in range(min(len(before), len(after)))
            if addressing.normalize(before[i]) != addressing.normalize(after[i])
        ]
        return {"replaced": len(touched), "affected": touched,
                "note": "count approximated from changed paragraphs"}

    def edit_range(self, doc, from_para, to_para, expect_hashes, new_text):
        raise DocdError(
            UNSUPPORTED_ON_BACKEND,
            "Index-based range rewrites are not reliable on HWP. Use doc_replace "
            "with the exact old text, or fields (doc_insert where=bookmark).",
        )

    def apply_style(self, doc, from_para, to_para=None, style=None):
        raise DocdError(UNSUPPORTED_ON_BACKEND, "Styles are not supported on HWP yet.")

    def tables(self, doc, op, **kw):
        raise DocdError(
            UNSUPPORTED_ON_BACKEND,
            "HWP tables are not supported yet (fields inside cells is the "
            "planned mechanism).",
        )

    def selection(self, doc):
        raise DocdError(UNSUPPORTED_ON_BACKEND, "Selection reading is not supported on HWP yet.")

    @com_retry
    def hwp_fields(self, doc):
        """List named fields (누름틀) — the durable anchors for HWP editing."""
        hwp = self._doc(doc)
        raw = hwp.GetFieldList(0, 0x01) or ""
        names = [n for n in raw.split("\x02") if n]
        return {"fields": [{"name": n, "text": hwp.GetFieldText(n)} for n in names]}

    @com_retry
    def save(self, doc):
        hwp = self._doc(doc)
        path = self._docs.get(doc)
        if not path:
            raise DocdError(BAD_PARAMS,
                            "This document has never been saved; use doc_save_as with a path.")
        hwp.Save()
        return {"saved": True, "path": path}

    @com_retry
    def save_as(self, doc, path, format):
        hwp = self._doc(doc)
        if format not in SAVE_FORMATS:
            raise DocdError(
                SAVE_FORMAT_UNSUPPORTED,
                f"HWP cannot save '{format}'. Supported: {', '.join(SAVE_FORMATS)}.",
            )
        hwp.SaveAs(path, SAVE_FORMATS[format], "")
        self._docs[doc] = self._docs[doc] or path
        return {"saved": True, "path": path, "format": format}

    @com_retry
    def close(self, doc, discard_changes=False):
        hwp = self._doc(doc)
        hwp.HAction.Run("FileClose")  # TODO(verify): unsaved-changes prompt behavior
        del self._docs[doc]
        self._scan_cache.pop(doc, None)
        return {"closed": True, "was_dirty": None}

    def shutdown(self):
        self._docs.clear()
        self._scan_cache.clear()
        self._hwp = None
