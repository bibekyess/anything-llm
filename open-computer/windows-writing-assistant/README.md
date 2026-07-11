# Windows Writing Assistant — slice 1 prototype

Prototype of the document-automation layer designed in
[`../docs/windows-writing-assistant/com-toolset-design.md`](../docs/windows-writing-assistant/com-toolset-design.md):
an AI writing assistant that edits documents **directly on the user's Windows
machine**, through each office app's own object model (no VM, no screenshots,
no coordinate clicking). The user watches edits land live in their real Word
window; everything is undoable with Ctrl+Z.

**Slice 1 scope:** the `docd` sidecar skeleton (JSON-RPC over stdio, single
COM worker thread), the stable-addressing core (paragraph hashes, re-anchoring,
staleness), the **WordDriver** (MS Word via COM), a `fake` in-memory backend
for cross-platform testing, and the `doc-tools.ts` pi extension.
PowerPoint / Hancom HWP / LibreOffice are slice 2+ — they plug in as additional
driver classes behind the same dispatcher.

## Layout

```
windows-writing-assistant/
├── docd/                  # Python sidecar (long-running child process)
│   ├── __main__.py        # python -m docd [--backend word|fake]
│   ├── rpc.py             # line-delimited JSON-RPC over stdio
│   ├── registry.py        # method dispatch, handle routing (w1 = Word doc 1)
│   ├── addressing.py      # paragraph content hashes, ±8 re-anchoring, STALE_RANGE
│   ├── render.py          # [p0#3fa2]-style read rendering, markdown↔style mapping
│   ├── errors.py          # closed error-code enum shared with the TS layer
│   └── drivers/
│       ├── base.py        # driver interface
│       ├── word.py        # MS Word via COM (pywin32) — Windows only
│       └── fake.py        # in-memory backend for tests/dev on any OS
├── extension/doc-tools.ts # pi extension: doc_* tools -> sidecar JSON-RPC client
├── tests/                 # cross-platform: pytest, no Word needed
└── smoke/word_smoke.py    # manual end-to-end check on Windows + Word
```

## Running the tests (any OS)

```bash
cd open-computer/windows-writing-assistant
python -m pytest tests/ -v
```

The suite covers the addressing core, the driver contract (against `fake`),
and the full stdio RPC protocol via a real subprocess.

## Verifying against real Word (Windows)

```powershell
cd open-computer\windows-writing-assistant
pip install pywin32
python smoke\word_smoke.py
```

Opens a scratch `.docx` in a visible Word window and walks every slice-1
method (insert with heading styles, literal + regex replace, staleness
refusal, hash-guarded edit, style application, save-as PDF) with a 1s pause
between steps so you can watch.

## Design invariants (from the design docs)

- **Direct on the host, not a VM.** open-computer's QEMU layer exists to
  isolate an untrusted agent; here the agent must edit the user's real
  documents. Safety comes from the object model instead: Range-based edits
  (never the user's Selection), Word's undo stack, content-hash staleness
  checks, and approval gates on destructive ops.
- **Range-based, never Selection-based** — doesn't steal the caret or focus.
- **Never hide, never quit** the app; attach to a running instance first
  (`GetActiveObject`), launch (`DispatchEx`) only if needed.
- **Stale edits are refused, not clobbered**: every mutating call can carry
  `expect_hash`(es) from the last `doc_read`; index drift up to ±8 paragraphs
  is auto-absorbed, content changes raise `STALE_RANGE` with fresh context.
- **All COM on one STA thread** (`rpc.py`); the stdin reader never touches COM.
- **Modal dialogs** surface as `APP_BUSY_MODAL` after a retry loop, never a
  hang or crash.

## Slice 2 candidates

- `doc_track_changes` (agent edits as redlined suggestions — recommended
  default for a writing assistant), `doc_comments`, `doc_format`, `doc_tables`,
  `doc_screenshot_page`
- PowerPoint driver (`slide_*` tools), Hancom HWP driver (field-anchored),
  LibreOffice UNO driver
- Dialog watchdog notifications + UIA sidecar integration
  (see `../docs/windows-writing-assistant/uia-port-design.md`)
