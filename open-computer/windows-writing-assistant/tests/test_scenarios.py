"""The two product scenarios, end-to-end over the real stdio RPC protocol:

1. "Write a report on Korean culture" -> doc_new + styled doc_insert + save.
2. "Convert my selected text into a table" -> doc_selection + doc_tables
   create with replace_range.
"""

import json
import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class SidecarProc:
    def __init__(self):
        self.proc = subprocess.Popen(
            [sys.executable, "-m", "docd", "--backend", "fake"],
            cwd=ROOT,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
        )
        self._id = 0

    def call(self, method, params=None):
        self._id += 1
        self.proc.stdin.write(
            json.dumps({"id": str(self._id), "method": method, "params": params or {}}) + "\n"
        )
        self.proc.stdin.flush()
        return json.loads(self.proc.stdout.readline())

    def result(self, method, params=None):
        msg = self.call(method, params)
        assert "result" in msg, f"{method} failed: {msg.get('error')}"
        return msg["result"]

    def stop(self):
        self.proc.stdin.close()
        self.proc.wait(timeout=10)


@pytest.fixture()
def sidecar():
    s = SidecarProc()
    yield s
    s.stop()


def test_scenario_generate_report_into_new_doc(sidecar, tmp_path):
    """User: 'write me a report on Korean culture' -> agent generates markdown
    and dumps it into a fresh document with real heading styles."""
    doc = sidecar.result("doc_new")["doc"]

    # Unsaved doc refuses plain save with a helpful error.
    msg = sidecar.call("doc_save", {"doc": doc})
    assert msg["error"]["code"] == "BAD_PARAMS"
    assert "doc_save_as" in msg["error"]["message"]

    report = (
        "# Korean Culture\n"
        "Korea has a rich cultural heritage spanning millennia.\n"
        "## Cuisine\n"
        "Kimchi is a staple of every meal.\n"
        "## Music\n"
        "From pansori to K-pop, music is central to Korean identity."
    )
    ins = sidecar.result("doc_insert", {"doc": doc, "text": report, "where": "end"})
    assert ins["inserted"] == 6

    outline = sidecar.result("doc_outline", {"doc": doc})["text"]
    assert "H1" in outline and "Korean Culture" in outline
    assert "H2" in outline and "Cuisine" in outline

    out = tmp_path / "korean-culture.md"
    saved = sidecar.result("doc_save_as", {"doc": doc, "path": str(out), "format": "md"})
    content = out.read_text(encoding="utf-8")
    assert "Kimchi is a staple" in content


def test_scenario_convert_selection_to_table(sidecar, tmp_path):
    """User selects messy lines and asks 'make this a table' -> agent reads the
    selection, parses it, and replaces those paragraphs with a real table."""
    path = tmp_path / "notes.txt"
    path.write_text(
        "Meeting notes\n"
        "Seoul - population 9.4 million\n"
        "Busan - population 3.3 million\n"
        "Incheon - population 3.0 million\n"
        "End of notes",
        encoding="utf-8",
    )
    doc = sidecar.result("doc_open", {"path": str(path)})["doc"]

    # No selection yet -> collapsed, agent gets a clear signal.
    sel = sidecar.result("doc_selection", {"doc": doc})
    assert sel["collapsed"] is True

    # The user selects the three city lines (test hook simulates the UI).
    sidecar.result("debug_set_selection", {"doc": doc, "from_para": 1, "to_para": 3})
    sel = sidecar.result("doc_selection", {"doc": doc})
    assert sel["from_para"] == 1 and sel["to_para"] == 3
    assert "Busan" in sel["text"]
    assert len(sel["hashes"]) == 3

    # Agent (the LLM) parses the selection into rows, then swaps it for a table.
    created = sidecar.result("doc_tables", {
        "doc": doc,
        "op": "create",
        "values": [
            ["City", "Population"],
            ["Seoul", "9.4 million"],
            ["Busan", "3.3 million"],
            ["Incheon", "3.0 million"],
        ],
        "replace_range": {
            "from_para": sel["from_para"],
            "to_para": sel["to_para"],
            "expect_hashes": sel["hashes"],
        },
        "header_row": True,
    })
    assert created["rows"] == 4 and created["cols"] == 2
    assert created["deleted_paras"] == 3

    # Source lines are gone; surrounding text intact.
    text = sidecar.result("doc_read", {"doc": doc})["text"]
    assert "Seoul - population" not in text
    assert "Meeting notes" in text and "End of notes" in text

    # Table is listed, readable, and editable.
    tables = sidecar.result("doc_tables", {"doc": doc, "op": "list"})["tables"]
    assert len(tables) == 1
    read = sidecar.result("doc_tables", {"doc": doc, "op": "read", "table": "t0"})
    assert "| Busan | 3.3 million |" in read["text"]
    sidecar.result("doc_tables", {
        "doc": doc, "op": "write", "table": "t0",
        "cell": {"row": 1, "col": 1}, "value": "9.5 million",
    })
    read = sidecar.result("doc_tables", {"doc": doc, "op": "read", "table": "t0"})
    assert "9.5 million" in read["text"]


def test_table_create_stale_selection_refused(sidecar, tmp_path):
    """If the user edits the selected paragraphs before the agent acts, the
    table swap must be refused, not clobber their text."""
    path = tmp_path / "n.txt"
    path.write_text("a\nb\nc", encoding="utf-8")
    doc = sidecar.result("doc_open", {"path": str(path)})["doc"]
    sidecar.result("debug_set_selection", {"doc": doc, "from_para": 0, "to_para": 1})
    sel = sidecar.result("doc_selection", {"doc": doc})

    # User keeps typing: paragraph 1 changes after the agent read the selection.
    sidecar.result("doc_replace", {"doc": doc, "find": "b", "replace": "b edited"})

    msg = sidecar.call("doc_tables", {
        "doc": doc, "op": "create", "values": [["x"]],
        "replace_range": {
            "from_para": 0, "to_para": 1, "expect_hashes": sel["hashes"],
        },
    })
    assert msg["error"]["code"] == "STALE_RANGE"
    text = sidecar.result("doc_read", {"doc": doc})["text"]
    assert "b edited" in text  # user's edit survived


def test_table_create_anchored(sidecar, tmp_path):
    path = tmp_path / "d.txt"
    path.write_text("intro\noutro", encoding="utf-8")
    doc = sidecar.result("doc_open", {"path": str(path)})["doc"]
    created = sidecar.result("doc_tables", {
        "doc": doc, "op": "create", "values": [["only", "row"]],
        "at": "after_para", "para": 0,
    })
    assert created["at_para"] == 1
    assert created["deleted_paras"] == 0
