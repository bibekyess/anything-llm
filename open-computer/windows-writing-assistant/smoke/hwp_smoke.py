"""Manual smoke test for HwpDriver — run on Windows with Hancom Office:

    cd open-computer\\windows-writing-assistant
    python smoke\\hwp_smoke.py

IMPORTANT: the HWP driver contains TODO(verify) API calls that have not yet
run against a real Hancom install (see docd/drivers/hwp.py). This script IS
the verification: run it, and report every step that fails with its output —
each failure pinpoints a call signature to adjust for your Hancom version.

If Hancom shows a security popup on open, either click allow, or set
HWP_SECURITY_DLL to the path of Hancom's FilePathChecker DLL (from the
automation SDK) before running, which registers the suppression module.
"""

import os
import sys
import tempfile
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docd.drivers.hwp import HwpDriver  # noqa: E402


def step(label, fn):
    print(f"\n=== {label} ===")
    try:
        result = fn()
        print((result if isinstance(result, str) else repr(result))[:600])
        time.sleep(1.0)
        return result
    except Exception as e:
        print(f"FAILED: {type(e).__name__}: {e}")
        traceback.print_exc(limit=2)
        return None


def main():
    driver = HwpDriver()

    created = step("doc_new", driver.new_doc)
    if not created:
        print("\nCannot continue without a document — is Hancom Office installed?")
        return
    handle = created["doc"]

    step("doc_insert at end", lambda: driver.insert(
        handle,
        "한국 문화 보고서\n"
        "청자: 고려 시대의 대표적인 도자기입니다.\n"
        "판소리: 전통 서사 노래입니다.",
        where="end"))

    step("doc_read", lambda: driver.read(handle)["text"])

    step("doc_replace (all)", lambda: driver.replace(
        handle, "도자기", "청자 도자기"))
    step("doc_replace (first occurrence)", lambda: driver.replace(
        handle, "전통", "한국 전통", occurrence="first"))

    step("hwp_fields (list)", lambda: driver.hwp_fields(handle))

    hwpx = os.path.join(tempfile.gettempdir(), "docd_smoke.hwpx")
    hwp = os.path.join(tempfile.gettempdir(), "docd_smoke.hwp")
    pdf = os.path.join(tempfile.gettempdir(), "docd_smoke_hwp.pdf")
    step("doc_save_as hwpx", lambda: driver.save_as(handle, hwpx, "hwpx"))
    step("doc_save_as hwp", lambda: driver.save_as(handle, hwp, "hwp"))
    step("doc_save_as pdf", lambda: driver.save_as(handle, pdf, "pdf"))

    step("re-open the saved hwpx", lambda: driver.open(hwpx))
    step("doc_close", lambda: driver.close(handle))

    print(f"\nSmoke run finished. Files (if saves succeeded): {hwpx}, {hwp}, {pdf}")
    print("Report any FAILED steps with their output — each maps to a "
          "TODO(verify) in docd/drivers/hwp.py.")


if __name__ == "__main__":
    main()
