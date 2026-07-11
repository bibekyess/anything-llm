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
