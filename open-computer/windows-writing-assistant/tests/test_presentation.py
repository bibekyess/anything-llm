"""slide_* contract tests against FakePresDriver (mirrors PptDriver), plus
dispatcher routing for presentation handles over the real stdio protocol."""

import json
import os
import subprocess
import sys

import pytest

from docd.drivers.fake_pres import FakePresDriver
from docd.errors import DocdError

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture()
def driver():
    return FakePresDriver()


@pytest.fixture()
def deck(driver):
    handle = driver.new_doc()["doc"]
    driver.slide_edit_text(handle, 0, "sh1", "2026 Business Plan")
    driver.slide_add(handle, title="Background", body="point one\npoint two")
    driver.slide_add(handle, title="Budget", body="numbers")
    return handle


def test_new_deck_and_list(driver, deck):
    res = driver.slide_list(deck)
    assert res["count"] == 3
    assert res["slides"][0]["title"] == "2026 Business Plan"
    assert res["slides"][1]["title"] == "Background"
    assert all(len(s["hash"]) == 4 for s in res["slides"])


def test_slide_read_shapes_and_hashes(driver, deck):
    res = driver.slide_read(deck, 1)
    kinds = {s["placeholder"]: s for s in res["shapes"]}
    assert kinds["title"]["text"] == "Background"
    assert "point one" in kinds["body"]["text"]
    assert kinds["body"]["shape"].startswith("sh")


def test_slide_edit_text_with_hash_guard(driver, deck):
    read = driver.slide_read(deck, 1)
    body = next(s for s in read["shapes"] if s["placeholder"] == "body")
    res = driver.slide_edit_text(
        deck, 1, body["shape"], "rewritten\n\tindented", expect_hash=body["hash"]
    )
    assert res["hash"] != body["hash"]
    with pytest.raises(DocdError) as exc:
        driver.slide_edit_text(deck, 1, body["shape"], "x", expect_hash=body["hash"])
    assert exc.value.code == "STALE_RANGE"


def test_slide_reanchoring_after_insert(driver, deck):
    budget_hash = driver.slide_list(deck)["slides"][2]["hash"]
    driver.slide_add(deck, after_slide=0, title="Inserted")  # shifts Budget to 3
    idx = driver._resolve_slide(driver._deck(deck), 2, budget_hash)
    assert idx == 3


def test_notes_reorder_delete(driver, deck):
    driver.slide_notes_edit(deck, 2, "emphasize the totals")
    assert driver.slide_read(deck, 2)["notes"] == "emphasize the totals"
    driver.slide_reorder(deck, 2, 0)
    assert driver.slide_list(deck)["slides"][0]["title"] == "Budget"
    res = driver.slide_delete(deck, 0)
    assert res["count"] == 2


def test_duplicate_slide(driver, deck):
    res = driver.slide_add(deck, duplicate_of=1)
    dup = driver.slide_read(deck, res["slide"])
    assert any(s["text"] == "Background" for s in dup["shapes"])


def test_doc_tools_unsupported_on_presentation(driver, deck):
    with pytest.raises(DocdError) as exc:
        driver.read(deck)
    assert exc.value.code == "UNSUPPORTED_ON_BACKEND"
    assert "slide_" in exc.value.message


class TestRpcRouting:
    @pytest.fixture()
    def sidecar(self):
        proc = subprocess.Popen(
            [sys.executable, "-m", "docd", "--backend", "fake"],
            cwd=ROOT, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, text=True, encoding="utf-8",
        )
        counter = [0]

        def call(method, params=None):
            counter[0] += 1
            proc.stdin.write(json.dumps(
                {"id": str(counter[0]), "method": method, "params": params or {}}) + "\n")
            proc.stdin.flush()
            return json.loads(proc.stdout.readline())

        yield call
        proc.stdin.close()
        proc.wait(timeout=10)

    def test_presentation_flow_over_rpc(self, sidecar):
        handle = sidecar("doc_new", {"app": "fakepres"})["result"]["doc"]
        assert handle.startswith("q")
        sidecar("slide_add", {"doc": handle, "title": "제목", "body": "내용 한 줄"})
        listed = sidecar("slide_list", {"doc": handle})["result"]
        assert listed["count"] == 2
        assert listed["slides"][1]["title"] == "제목"

        msg = sidecar("doc_read", {"doc": handle})
        assert msg["error"]["code"] == "UNSUPPORTED_ON_BACKEND"

        # slide tools on a text-document handle are rejected cleanly too
        doc = sidecar("doc_new", {"app": "fake"})["result"]["doc"]
        msg = sidecar("slide_list", {"doc": doc})
        assert msg["error"]["code"] == "UNSUPPORTED_ON_BACKEND"

    def test_thumbnail_unsupported_on_fake(self, sidecar):
        handle = sidecar("doc_new", {"app": "fakepres"})["result"]["doc"]
        msg = sidecar("slide_thumbnail", {"doc": handle, "slide": 0})
        assert msg["error"]["code"] == "UNSUPPORTED_ON_BACKEND"
