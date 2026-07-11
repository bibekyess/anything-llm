"""Manual smoke test for WordDriver — run on Windows with Word installed:

    cd open-computer\\windows-writing-assistant
    pip install pywin32
    python smoke\\word_smoke.py

Opens a VISIBLE Word window and walks every slice-1/1.5 method, including the
two product scenarios (generate a styled report; convert selected text into a
table). Nothing in your own documents is touched — everything happens in a
scratch document saved to %TEMP%.
"""

import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docd.addressing import para_hash  # noqa: E402
from docd.drivers.word import WordDriver  # noqa: E402
from docd.errors import DocdError  # noqa: E402


def step(label, result):
    print(f"\n=== {label} ===")
    print(result if isinstance(result, str) else repr(result))
    time.sleep(1.0)  # so you can watch each edit land in the Word window


def main():
    driver = WordDriver()

    # ── Scenario 1: generate a styled report into a fresh document ────
    created = driver.new_doc()
    handle = created["doc"]
    step("doc_new", created)

    step("doc_insert (styled report — scenario 1)", driver.insert(
        handle,
        "# Korean Culture\n"
        "Korea has a rich cultural heritage spanning millennia.\n"
        "## Cuisine\n"
        "Kimchi is a staple of every meal.\n"
        "## Music\n"
        "From pansori to K-pop, music is central to Korean identity.",
        where="end",
    ))
    step("doc_outline", driver.outline(handle)["text"])

    scratch = os.path.join(tempfile.gettempdir(), "docd_smoke.docx")
    step("doc_save_as docx", driver.save_as(handle, scratch, "docx"))

    # ── Editing primitives ─────────────────────────────────────────────
    step("doc_replace (literal)", driver.replace(
        handle, "staple", "beloved staple"))
    step("doc_replace (regex)", driver.replace(
        handle, r"K-?pop", "K-Pop", regex=True))

    try:
        driver.insert(handle, "x", "after_para", para=1, expect_hash="beef")
        print("ERROR: stale hash was NOT refused")
    except DocdError as e:
        step("stale hash correctly refused", f"[{e.code}] {e.message}")

    good = para_hash("Korea has a rich cultural heritage spanning millennia.")
    step("doc_insert with valid expect_hash", driver.insert(
        handle, "Reviewed for accuracy.", "after_para", para=1, expect_hash=good))

    step("doc_apply_style", driver.apply_style(handle, 0, style="Title"))

    # ── Scenario 2: convert text into a real Word table ────────────────
    step("insert city lines for the table demo", driver.insert(
        handle,
        "Seoul - population 9.4 million\n"
        "Busan - population 3.3 million\n"
        "Incheon - population 3.0 million",
        where="end",
    ))

    print("\n>>> In the Word window: SELECT the three city lines with your "
          "mouse, then press Enter here (or just press Enter to use their "
          "known positions).")
    input()

    sel = driver.selection(handle)
    step("doc_selection", sel)
    if sel["collapsed"]:
        # Fall back to locating the lines by content.
        read = driver.read(handle)
        lines = [l for l in read["text"].splitlines() if "population" in l]
        first = int(lines[0].split("#")[0][2:])
        sel = {
            "from_para": first,
            "to_para": first + 2,
            "hashes": [l.split("#")[1].split("]")[0] for l in lines],
        }
        print(f"(no selection made — using p{sel['from_para']}..p{sel['to_para']})")

    step("doc_tables create (replace selection — scenario 2)", driver.tables(
        handle, "create",
        values=[
            ["City", "Population"],
            ["Seoul", "9.4 million"],
            ["Busan", "3.3 million"],
            ["Incheon", "3.0 million"],
        ],
        replace_range={
            "from_para": sel["from_para"],
            "to_para": sel["to_para"],
            "expect_hashes": sel["hashes"],
        },
        header_row=True,
    ))
    step("doc_tables list", driver.tables(handle, "list"))
    step("doc_tables read", driver.tables(handle, "read", table="t0")["text"])
    step("doc_tables write (update a cell)", driver.tables(
        handle, "write", table="t0", cell={"row": 1, "col": 1},
        value="9.5 million"))

    # ── Save / list / close ────────────────────────────────────────────
    pdf = os.path.join(tempfile.gettempdir(), "docd_smoke.pdf")
    step("doc_save_as pdf", driver.save_as(handle, pdf, "pdf"))
    step("doc_save", driver.save(handle))
    step("doc_list_open", driver.list_open())
    step("doc_close", driver.close(handle))

    print(f"\nSmoke test complete. Inspect: {scratch} and {pdf}")


if __name__ == "__main__":
    main()
