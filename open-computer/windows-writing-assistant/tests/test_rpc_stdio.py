"""End-to-end protocol test: spawn `python -m docd --backend fake` as a real
subprocess and speak line-delimited JSON-RPC over its stdio — exactly what the
TS extension does."""

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
        req = {"id": str(self._id), "method": method, "params": params or {}}
        self.proc.stdin.write(json.dumps(req) + "\n")
        self.proc.stdin.flush()
        line = self.proc.stdout.readline()
        assert line, "sidecar closed stdout unexpectedly"
        msg = json.loads(line)
        assert msg["id"] == str(self._id)
        return msg

    def stop(self):
        self.proc.stdin.close()
        self.proc.wait(timeout=10)


@pytest.fixture()
def sidecar():
    s = SidecarProc()
    yield s
    s.stop()


def test_ping(sidecar):
    res = sidecar.call("ping")["result"]
    assert res["service"] == "docd"
    assert "fake" in res["backends"]


def test_full_edit_session(sidecar, tmp_path):
    path = tmp_path / "story.txt"
    path.write_text("Once upon a time.\nThe end.", encoding="utf-8")

    opened = sidecar.call("doc_open", {"path": str(path)})["result"]
    handle = opened["doc"]
    assert opened["paragraphs"] == 2

    read = sidecar.call("doc_read", {"doc": handle})["result"]
    assert "[p0#" in read["text"] and "Once upon a time." in read["text"]

    rep = sidecar.call(
        "doc_replace", {"doc": handle, "find": "The end", "replace": "To be continued"}
    )["result"]
    assert rep["replaced"] == 1

    ins = sidecar.call(
        "doc_insert", {"doc": handle, "text": "A new chapter.", "where": "end"}
    )["result"]
    assert ins["inserted"] == 1

    saved = sidecar.call("doc_save", {"doc": handle})["result"]
    content = open(saved["path"], encoding="utf-8").read()
    assert "To be continued" in content and "A new chapter." in content


def test_error_envelope(sidecar):
    msg = sidecar.call("doc_read", {"doc": "w99"})
    assert msg["error"]["code"] == "NO_SUCH_DOC"

    msg = sidecar.call("no_such_method")
    assert msg["error"]["code"] == "UNKNOWN_METHOD"

    msg = sidecar.call("doc_open", {})
    assert msg["error"]["code"] == "BAD_PARAMS"


def test_bad_json_does_not_kill_sidecar(sidecar):
    sidecar.proc.stdin.write("this is not json\n")
    sidecar.proc.stdin.flush()
    line = sidecar.proc.stdout.readline()
    assert json.loads(line)["error"]["code"] == "BAD_PARAMS"
    assert sidecar.call("ping")["result"]["service"] == "docd"


def test_unknown_app_extension(sidecar, tmp_path):
    msg = sidecar.call("doc_open", {"path": str(tmp_path / "x.xyz")})
    assert msg["error"]["code"] == "BAD_PARAMS"
    # Word backend is not loaded in --backend fake builds.
    msg = sidecar.call("doc_open", {"path": str(tmp_path / "x.docx")})
    assert "not available" in msg["error"]["message"]
