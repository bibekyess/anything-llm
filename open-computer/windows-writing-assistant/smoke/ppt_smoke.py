"""Manual smoke test for PptDriver — run on Windows with PowerPoint installed:

    cd open-computer\\windows-writing-assistant
    python smoke\\ppt_smoke.py

Builds a small deck in a visible PowerPoint window, walking every slide_*
method. Files land in %TEMP%.
"""

import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docd.drivers.ppt import PptDriver  # noqa: E402
from docd.errors import DocdError  # noqa: E402


def step(label, result):
    print(f"\n=== {label} ===")
    text = result if isinstance(result, str) else repr(result)
    print(text[:600])
    time.sleep(1.0)


def main():
    driver = PptDriver()

    created = driver.new_doc()
    handle = created["doc"]
    step("doc_new (presentation)", created)

    step("slide_list (fresh deck)", driver.slide_list(handle))

    step("slide_add: title+content", driver.slide_add(
        handle, title="한국 문화 개요",
        body="유구한 역사\n\t고려청자\n\t조선백자\nK-컬처의 세계화"))
    step("slide_add: second", driver.slide_add(
        handle, title="Cuisine", body="Kimchi\nBibimbap\nTteokbokki"))

    listed = driver.slide_list(handle)
    step("slide_list", listed)

    read = driver.slide_read(handle, 1)
    step("slide_read s1", read)

    body = next((s for s in read["shapes"] if s["placeholder"] == "body"), None)
    if body:
        step("slide_edit_text (hash-guarded)", driver.slide_edit_text(
            handle, 1, body["shape"], "유구한 역사\n\t고려청자\nK-컬처의 세계화 (수정)",
            expect_hash=body["hash"]))
        try:
            driver.slide_edit_text(handle, 1, body["shape"], "x", expect_hash=body["hash"])
            print("ERROR: stale hash was NOT refused")
        except DocdError as e:
            step("stale hash correctly refused", f"[{e.code}] {e.message}")

    step("slide_notes_edit", driver.slide_notes_edit(
        handle, 1, "발표 시 강조: 전통과 현대의 공존"))
    step("slide_add duplicate", driver.slide_add(handle, duplicate_of=1))
    step("slide_reorder (last -> position 1)", driver.slide_reorder(
        handle, driver.slide_list(handle)["count"] - 1, 1))

    thumb = driver.slide_thumbnail(handle, 1)
    step("slide_thumbnail", {k: (v[:40] + "..." if k == "png_base64" else v)
                             for k, v in thumb.items()})

    step("slide_delete", driver.slide_delete(handle, 1))

    out = os.path.join(tempfile.gettempdir(), "docd_smoke.pptx")
    pdf = os.path.join(tempfile.gettempdir(), "docd_smoke_deck.pdf")
    step("pres_save_as pptx", driver.pres_save_as(handle, out, "pptx"))
    step("pres_save_as pdf", driver.pres_save_as(handle, pdf, "pdf"))
    step("doc_list_open", driver.list_open())
    step("doc_close", driver.close(handle))

    print(f"\nSmoke test complete. Inspect: {out} and {pdf}")


if __name__ == "__main__":
    main()
