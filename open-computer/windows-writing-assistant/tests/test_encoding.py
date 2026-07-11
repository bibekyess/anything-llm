"""UTF-8 pipeline safety: Korean text and typographic characters must survive
the stdio RPC round-trip even when the platform's default stdio encoding is
hostile (Windows defaults piped stdio to the ANSI code page, e.g. cp1252 —
that shredded 청자 into surrogate escapes and em-dashes into 'â€”').

We simulate the Windows condition by forcing PYTHONIOENCODING=cp1252 on the
sidecar subprocess; force_utf8_stdio() must override it.
"""

import json
import os
import subprocess
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

KOREAN = "청자: 고려 시대에 생산된 도자기입니다."
TYPOGRAPHIC = "tradition—meets—tech “quoted” café"


class SidecarProc:
    def __init__(self, hostile_encoding=None):
        env = {**os.environ}
        if hostile_encoding:
            env["PYTHONIOENCODING"] = hostile_encoding
            env.pop("PYTHONUTF8", None)
        self.proc = subprocess.Popen(
            [sys.executable, "-m", "docd", "--backend", "fake"],
            cwd=ROOT,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env,
        )
        self._id = 0

    def call(self, method, params=None):
        self._id += 1
        req = {"id": str(self._id), "method": method, "params": params or {}}
        # Raw UTF-8 bytes on the pipe — exactly what doc-tools.ts sends.
        self.proc.stdin.write((json.dumps(req, ensure_ascii=False) + "\n").encode("utf-8"))
        self.proc.stdin.flush()
        line = self.proc.stdout.readline()
        return json.loads(line.decode("utf-8"))

    def result(self, method, params=None):
        msg = self.call(method, params)
        assert "result" in msg, f"{method} failed: {msg.get('error')}"
        return msg["result"]

    def stop(self):
        self.proc.stdin.close()
        self.proc.wait(timeout=10)


@pytest.fixture(params=[None, "cp1252"], ids=["default-env", "hostile-cp1252-env"])
def sidecar(request):
    s = SidecarProc(hostile_encoding=request.param)
    yield s
    s.stop()


def test_korean_and_typographic_roundtrip(sidecar, tmp_path):
    doc = sidecar.result("doc_new")["doc"]

    ins = sidecar.result("doc_insert", {
        "doc": doc, "text": f"# 한국 문화\n{KOREAN}\n{TYPOGRAPHIC}", "where": "end",
    })
    assert ins["inserted"] == 3

    text = sidecar.result("doc_read", {"doc": doc})["text"]
    assert KOREAN in text
    assert TYPOGRAPHIC in text
    assert "â€" not in text          # the cp1252 mojibake signature
    assert "\udc90" not in text      # no surrogate escapes

    # Replace English with Korean — the failing case from the field report.
    sidecar.result("doc_insert", {"doc": doc, "text": "old western text", "where": "end"})
    rep = sidecar.result("doc_replace", {
        "doc": doc, "find": "old western text", "replace": "탈춤: 가면 무용극",
    })
    assert rep["replaced"] == 1
    text = sidecar.result("doc_read", {"doc": doc})["text"]
    assert "탈춤: 가면 무용극" in text

    # Hash-guarded edit_range with Korean, end to end.
    read = sidecar.result("doc_read", {"doc": doc})
    lines = [l for l in read["text"].splitlines() if KOREAN in l]
    idx = int(lines[0].split("#")[0][2:])
    h = lines[0].split("#")[1].split("]")[0]
    edited = sidecar.result("doc_edit_range", {
        "doc": doc, "from_para": idx, "to_para": idx,
        "expect_hashes": [h], "new_text": "백자: 조선 시대의 도자기입니다.",
    })
    assert edited["replaced"] == 1

    out = tmp_path / "korean.md"
    sidecar.result("doc_save_as", {"doc": doc, "path": str(out), "format": "md"})
    content = out.read_text(encoding="utf-8")
    assert "백자" in content and "â€" not in content


def test_long_literal_replace(sidecar, tmp_path):
    """Word's Find.Execute caps find/replace at 255 chars; the driver must
    route longer literals through paragraph rewriting. The fake driver has no
    such limit, but this locks the tool-level contract both drivers share."""
    doc = sidecar.result("doc_new")["doc"]
    long_text = "sentence " * 40  # ~360 chars
    sidecar.result("doc_insert", {"doc": doc, "text": long_text.strip(), "where": "end"})
    replacement = "문장 " * 40    # long Korean replacement with backslash-free text
    rep = sidecar.result("doc_replace", {
        "doc": doc, "find": long_text.strip(), "replace": replacement.strip(),
    })
    assert rep["replaced"] == 1
    assert "문장" in sidecar.result("doc_read", {"doc": doc})["text"]


def test_literal_replace_with_backslashes(sidecar):
    r"""Literal replacements containing '\' must not be eaten as re templates."""
    doc = sidecar.result("doc_new")["doc"]
    sidecar.result("doc_insert", {"doc": doc, "text": "path goes here", "where": "end"})
    rep = sidecar.result("doc_replace", {
        "doc": doc, "find": "goes here", "replace": r"C:\Users\bibek \g<0>",
    })
    assert rep["replaced"] == 1
    text = sidecar.result("doc_read", {"doc": doc})["text"]
    assert r"C:\Users\bibek \g<0>" in text
