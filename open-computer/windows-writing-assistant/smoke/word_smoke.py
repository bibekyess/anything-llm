"""Manual smoke test for WordDriver — run on Windows with Word installed:

    cd open-computer/windows-writing-assistant
    pip install pywin32
    python smoke/word_smoke.py

Creates a scratch .docx in %TEMP%, opens it in a VISIBLE Word window, and
walks every slice-1 method so you can watch the edits land live. Nothing in
your own documents is touched; the scratch file is closed (kept on disk for
inspection) at the end.
"""

import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docd.drivers.word import WordDriver  # noqa: E402


def step(label, result):
    print(f"\n=== {label} ===")
    print(result if isinstance(result, str) else repr(result))
    time.sleep(1.0)  # so you can watch each edit land in the Word window


def main():
    driver = WordDriver()

    # Build a scratch docx through Word itself.
    app = driver._ensure_app()
    doc = app.Documents.Add()
    scratch = os.path.join(tempfile.gettempdir(), "docd_smoke.docx")
    doc.SaveAs2(FileName=scratch, FileFormat=12)  # wdFormatXMLDocument
    doc.Close()
    print(f"Scratch document: {scratch}")

    opened = driver.open(scratch)
    handle = opened["doc"]
    step("doc_open", opened)

    step("doc_insert (styled draft)", driver.insert(
        handle,
        "# Quarterly Report\nThis draft was written by the agent.\n"
        "## Findings\nRevenue grew steadily this quarter.\n"
        "## Next Steps\nExpand the pilot program.",
        where="end",
    ))

    read = driver.read(handle)
    step("doc_read", read["text"])
    step("doc_outline", driver.outline(handle)["text"])

    step("doc_replace (literal)", driver.replace(
        handle, "steadily", "faster than expected"))
    step("doc_replace (regex)", driver.replace(
        handle, r"pilot \w+", "pilot initiative", regex=True))

    # Staleness demo: a wrong hash must be refused, a right one accepted.
    from docd.addressing import para_hash
    from docd.errors import DocdError
    try:
        driver.insert(handle, "x", "after_para", para=1, expect_hash="beef")
        print("ERROR: stale hash was NOT refused")
    except DocdError as e:
        step("stale hash correctly refused", f"[{e.code}] {e.message}")

    good_hash = para_hash("This draft was written by the agent.")
    step("doc_insert with valid expect_hash", driver.insert(
        handle, "Reviewed for accuracy.", "after_para",
        para=1, expect_hash=good_hash,
    ))

    read = driver.read(handle)
    hashes = [line.split("#")[1].split("]")[0]
              for line in read["text"].splitlines() if line.startswith("[p")]
    step("doc_edit_range (rewrite p1)", driver.edit_range(
        handle, 1, 1, [hashes[1]],
        "This draft was co-written by the agent and reviewed by a human.",
    ))

    step("doc_apply_style", driver.apply_style(handle, 0, style="Title"))

    pdf = os.path.join(tempfile.gettempdir(), "docd_smoke.pdf")
    step("doc_save_as pdf", driver.save_as(handle, pdf, "pdf"))
    step("doc_save", driver.save(handle))
    step("doc_list_open", driver.list_open())
    step("doc_close", driver.close(handle))

    print(f"\nSmoke test complete. Inspect: {scratch} and {pdf}")


if __name__ == "__main__":
    main()
