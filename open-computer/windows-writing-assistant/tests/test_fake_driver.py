"""Driver-semantics tests against FakeDriver — the same contract WordDriver
implements, so these lock down tool-visible behavior (handles, hashes,
staleness, style_map) without needing Word."""

import re

import pytest

from docd.addressing import para_hash
from docd.drivers.fake import FakeDriver
from docd.errors import DocdError


@pytest.fixture()
def driver():
    return FakeDriver()


@pytest.fixture()
def doc(driver, tmp_path):
    path = tmp_path / "sample.md"
    path.write_text(
        "# Plan 2026\nIntro paragraph.\n## Budget\nNumbers go here.\nClosing.",
        encoding="utf-8",
    )
    return driver.open(str(path))["doc"]


def test_open_and_list(driver, doc):
    docs = driver.list_open()["docs"]
    assert len(docs) == 1
    assert docs[0]["doc"] == doc
    assert docs[0]["paragraphs"] == 5


def test_read_renders_ids_hashes_and_headings(driver, doc):
    res = driver.read(doc)
    assert res["count"] == 5
    lines = res["text"].splitlines()
    # [p0#xxxx] # Plan 2026
    assert re.match(r"^\[p0#[0-9a-f]{4}\] # Plan 2026$", lines[0])
    assert re.match(r"^\[p1#[0-9a-f]{4}\] Intro paragraph\.$", lines[1])
    assert lines[2].endswith("## Budget")
    assert lines[-1].startswith("rev:")


def test_read_range_and_truncation(driver, doc):
    res = driver.read(doc, from_para=2, to_para=3)
    assert "[p2#" in res["text"] and "[p4#" not in res["text"]
    tiny = driver.read(doc, max_chars=30)
    assert "truncated" in tiny["text"]


def test_outline(driver, doc):
    text = driver.outline(doc)["text"]
    assert "H1 [p0] Plan 2026" in text
    assert "H2 [p2] Budget" in text


def test_insert_after_para_with_reanchor(driver, doc):
    anchor_hash = para_hash("Numbers go here.")
    # Simulate the user inserting a paragraph above the anchor first.
    driver.insert(doc, "User typed this.", "before_para", para=0)
    res = driver.insert(
        doc, "AI addition.", "after_para", para=3, expect_hash=anchor_hash
    )
    assert res["moved"] is True
    assert res["first_para"] == 5
    read = driver.read(doc)["text"]
    assert "AI addition." in read


def test_insert_stale_anchor_refused(driver, doc):
    with pytest.raises(DocdError) as exc:
        driver.insert(doc, "x", "after_para", para=1, expect_hash="beef")
    assert exc.value.code == "STALE_RANGE"
    assert exc.value.data["context"]


def test_insert_style_map(driver, doc):
    res = driver.insert(doc, "## New Section\nBody text", "end")
    assert res["inserted"] == 2
    outline = driver.outline(doc)["text"]
    assert "New Section" in outline


def test_replace_all_and_occurrence(driver, doc):
    res = driver.replace(doc, "paragraph", "PARA")
    assert res["replaced"] == 1
    assert res["affected"][0][0] == 1
    driver.insert(doc, "alpha alpha alpha", "end")
    res = driver.replace(doc, "alpha", "beta", occurrence=2)
    assert res["replaced"] == 1
    assert "beta" in driver.read(doc)["text"]


def test_replace_regex_and_scope(driver, doc):
    res = driver.replace(doc, r"Numbers \w+", "Figures live", regex=True)
    assert res["replaced"] == 1
    res = driver.replace(doc, "Closing", "End", scope={"from_para": 0, "to_para": 3})
    assert res["replaced"] == 0  # out of scope


def test_edit_range_happy_path(driver, doc):
    paras = ["Intro paragraph."]
    res = driver.edit_range(
        doc, 1, 1, [para_hash(p) for p in paras], "Rewritten intro.\nExtra line."
    )
    assert res["replaced"] == 2
    text = driver.read(doc)["text"]
    assert "Rewritten intro." in text and "Extra line." in text


def test_edit_range_stale_hash_refused(driver, doc):
    with pytest.raises(DocdError) as exc:
        driver.edit_range(doc, 1, 1, ["0bad"], "nope")
    assert exc.value.code == "STALE_RANGE"
    assert "nope" not in driver.read(doc)["text"]


def test_edit_range_delete(driver, doc):
    h = para_hash("Closing.")
    res = driver.edit_range(doc, 4, 4, [h], "")
    assert res["deleted"] == 1
    assert "Closing." not in driver.read(doc)["text"]


def test_apply_style(driver, doc):
    res = driver.apply_style(doc, 4, style="Heading 2")
    assert res["styled"] == 1
    assert "H2 [p4] Closing." in driver.outline(doc)["text"]


def test_save_and_save_as(driver, doc, tmp_path):
    driver.replace(doc, "Intro", "Updated intro")
    res = driver.save(doc)
    assert "Updated intro" in open(res["path"], encoding="utf-8").read()
    out = tmp_path / "copy.txt"
    driver.save_as(doc, str(out), "txt")
    assert out.exists()
    with pytest.raises(DocdError) as exc:
        driver.save_as(doc, str(tmp_path / "x.pdf"), "pdf")
    assert exc.value.code == "SAVE_FORMAT_UNSUPPORTED"


def test_read_only_refuses_writes(driver, tmp_path):
    path = tmp_path / "ro.txt"
    path.write_text("locked", encoding="utf-8")
    handle = driver.open(str(path), read_only=True)["doc"]
    with pytest.raises(DocdError) as exc:
        driver.replace(handle, "locked", "unlocked")
    assert exc.value.code == "READ_ONLY"


def test_close_and_handle_invalidation(driver, doc):
    assert driver.close(doc)["closed"] is True
    with pytest.raises(DocdError) as exc:
        driver.read(doc)
    assert exc.value.code == "NO_SUCH_DOC"
