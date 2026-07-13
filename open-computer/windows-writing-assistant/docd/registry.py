"""Method dispatcher: routes doc_* RPC methods to the owning driver.

Handles are prefixed by backend ("w1" = Word, "f1" = fake); doc_open routes by
file extension (or explicit `app` override), doc_* calls route by handle
prefix. Slice 1 registers Word (or fake); HWP/UNO/PPT drivers plug in here.
"""

import os

from . import __version__
from .errors import DocdError, BAD_PARAMS, NO_SUCH_DOC, UNKNOWN_METHOD

EXT_TO_APP = {
    ".docx": "word", ".doc": "word", ".rtf": "word",
    ".txt": "fake", ".md": "fake",   # fake backend claims plain text in dev
    # slice 2+: .hwp/.hwpx -> hwp, .odt -> libreoffice, .pptx/.ppt -> powerpoint
}


class Dispatcher:
    def __init__(self, drivers):
        """drivers: {backend_name: BaseDriver instance}."""
        self.drivers = drivers
        self.methods = {
            "ping": self._ping,
            "doc_list_open": self._list_open,
            "doc_open": self._open,
            "doc_new": self._new,
            "doc_selection": self._by_handle("selection"),
            "doc_tables": self._by_handle("tables"),
            "debug_set_selection": self._by_handle("debug_set_selection"),
            "doc_read": self._by_handle("read"),
            "doc_outline": self._by_handle("outline"),
            "doc_insert": self._by_handle("insert"),
            "doc_replace": self._by_handle("replace"),
            "doc_edit_range": self._by_handle("edit_range"),
            "doc_apply_style": self._by_handle("apply_style"),
            "doc_save": self._by_handle("save"),
            "doc_save_as": self._by_handle("save_as"),
            "doc_close": self._by_handle("close"),
        }

    def handle(self, method, params):
        if method not in self.methods:
            raise DocdError(UNKNOWN_METHOD, f"Unknown method '{method}'.")
        return self.methods[method](**(params or {}))

    def shutdown(self):
        for driver in self.drivers.values():
            driver.shutdown()

    # ── routing ────────────────────────────────────────────────────────
    def _driver_for_handle(self, handle):
        for driver in self.drivers.values():
            if handle.startswith(driver.prefix):
                return driver
        raise DocdError(NO_SUCH_DOC, f"No backend for handle '{handle}'.")

    def _by_handle(self, method_name):
        def call(doc=None, **kwargs):
            if not doc:
                raise DocdError(BAD_PARAMS, f"'{method_name}' requires `doc`.")
            driver = self._driver_for_handle(doc)
            return getattr(driver, method_name)(doc, **kwargs)
        return call

    # ── methods ────────────────────────────────────────────────────────
    def _ping(self):
        return {
            "service": "docd",
            "version": __version__,
            "backends": sorted(self.drivers.keys()),
        }

    def _list_open(self):
        docs = []
        for driver in self.drivers.values():
            docs.extend(driver.list_open()["docs"])
        return {"docs": docs}

    def _new(self, app=None):
        backend = app or ("word" if "word" in self.drivers else "fake")
        if backend not in self.drivers:
            raise DocdError(
                BAD_PARAMS,
                f"Backend '{backend}' is not available in this build "
                f"(have: {', '.join(sorted(self.drivers))}).",
            )
        return self.drivers[backend].new_doc()

    def _open(self, path=None, app=None, read_only=False):
        if not path:
            raise DocdError(BAD_PARAMS, "doc_open requires `path`.")
        backend = app or EXT_TO_APP.get(os.path.splitext(path)[1].lower())
        if backend is None:
            raise DocdError(
                BAD_PARAMS,
                f"Cannot infer app for '{path}'; pass `app` explicitly.",
            )
        if backend not in self.drivers:
            raise DocdError(
                BAD_PARAMS,
                f"Backend '{backend}' is not available in this build "
                f"(have: {', '.join(sorted(self.drivers))}).",
            )
        return self.drivers[backend].open(path, read_only=read_only)
