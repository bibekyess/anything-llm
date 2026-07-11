"""Driver interface. One instance per backend; documents addressed by handle.

Method semantics follow the tool schemas in com-toolset-design.md §2.1.
Every method returns a JSON-serializable dict; errors are raised as DocdError.
"""


class BaseDriver:
    #: single-letter handle prefix ("w" = Word, "f" = fake, later "h"/"o"/"p")
    prefix = "?"
    #: backend name reported by doc_list_open
    backend = "?"

    def list_open(self):
        """-> {docs: [{doc, path, dirty, paragraphs, read_only}]}"""
        raise NotImplementedError

    def open(self, path, read_only=False):
        """-> {doc, path, paragraphs, dirty, read_only}"""
        raise NotImplementedError

    def new_doc(self):
        """Create a blank document. -> {doc, path: None, paragraphs}"""
        raise NotImplementedError

    def selection(self, doc):
        """Read the user's current selection in the document window.
        -> {collapsed, text, from_para, to_para, hashes}"""
        raise NotImplementedError

    def tables(self, doc, op, table=None, cell=None, value=None, values=None,
               at=None, para=None, expect_hash=None, replace_range=None,
               header_row=False):
        """op: list | read | write | create. See com-toolset-design.md §2.1.
        -> op-specific dict"""
        raise NotImplementedError

    def debug_set_selection(self, doc, from_para, to_para):
        """Test hook (fake backend only): simulate a user selection."""
        from ..errors import DocdError, UNSUPPORTED_ON_BACKEND
        raise DocdError(
            UNSUPPORTED_ON_BACKEND,
            "debug_set_selection is only available on the fake backend.",
        )

    def read(self, doc, from_para=None, to_para=None, max_chars=None):
        """-> {text, from_para, to_para, count, rev}"""
        raise NotImplementedError

    def outline(self, doc):
        """-> {text}"""
        raise NotImplementedError

    def insert(self, doc, text, where, para=None, expect_hash=None,
               bookmark=None, style_map=True):
        """-> {inserted, first_para, affected: [[index, hash]], moved}"""
        raise NotImplementedError

    def replace(self, doc, find, replace, regex=False, match_case=False,
                occurrence="all", scope=None):
        """-> {replaced, affected: [[index, hash]]}"""
        raise NotImplementedError

    def edit_range(self, doc, from_para, to_para, expect_hashes, new_text):
        """-> {deleted|replaced, affected: [[index, hash]]}"""
        raise NotImplementedError

    def apply_style(self, doc, from_para, to_para=None, style=None):
        """-> {styled, affected: [[index, hash]]}"""
        raise NotImplementedError

    def save(self, doc):
        """-> {saved, path}"""
        raise NotImplementedError

    def save_as(self, doc, path, format):
        """-> {saved, path, format}"""
        raise NotImplementedError

    def close(self, doc, discard_changes=False):
        """-> {closed, was_dirty}"""
        raise NotImplementedError

    def shutdown(self):
        """Release app/document references before process exit."""
