# Windows Writing Assistant — document automation prototype

Prototype of the document-automation layer designed in
[`../docs/windows-writing-assistant/com-toolset-design.md`](../docs/windows-writing-assistant/com-toolset-design.md):
an AI writing assistant that edits documents **directly on the user's Windows
machine**, through each office app's own object model (no VM, no screenshots,
no coordinate clicking). The user watches edits land live in their real Word
window; everything is undoable with Ctrl+Z.

**Implemented (slices 1 + 1.5):** the `docd` sidecar (JSON-RPC over stdio,
single COM worker thread), stable paragraph addressing (content hashes,
re-anchoring, staleness refusal), the **WordDriver** (MS Word via COM), a
`fake` in-memory backend for cross-platform testing, and the `doc-tools.ts`
pi extension with 14 tools:

`doc_new` · `doc_open` · `doc_list_open` · `doc_read` · `doc_outline` ·
`doc_selection` · `doc_insert` · `doc_replace` · `doc_edit_range` ·
`doc_apply_style` · `doc_tables` (list/read/write/create) · `doc_save` ·
`doc_save_as` · `doc_close`

These cover the two target workflows end-to-end:

1. **"Write me a report on X"** → the LLM writes markdown → `doc_new` +
   `doc_insert` (markdown `#`/`##` become real Word Heading styles) →
   `doc_save_as`.
2. **"Turn my selected text into a table"** → `doc_selection` (reads what the
   user highlighted, with paragraph hashes) → the LLM parses it into rows →
   `doc_tables create` with `replace_range` swaps those paragraphs for a real
   Word table — refused with `STALE_RANGE` if the user edited them meanwhile.

PowerPoint / Hancom HWP / LibreOffice are next slices — they plug in as
additional driver classes behind the same dispatcher.

---

## 🚀 Quick start on your Windows PC

### What you need

- Windows 10/11 with **Microsoft Word installed** (desktop version, any recent one)
- **Python 3.10+** — from [python.org](https://www.python.org/downloads/) or
  `winget install Python.Python.3.12` (tick *Add python to PATH* if using the installer)

### 1. Get the code

```powershell
git clone -b claude/open-computer-writing-assistant-cmntky https://github.com/bibekyess/anything-llm
cd anything-llm\open-computer\windows-writing-assistant
```

### 2. Install dependencies

```powershell
pip install pywin32 pytest
```

### 3. Run the cross-platform tests (no Word touched)

```powershell
python -m pytest tests\ -v
```

Expected: **37 passed**. These exercise the addressing core, the driver
contract, the stdio RPC protocol, and both product scenarios against the
in-memory backend.

### 4. Run the live Word smoke test 👀

Close anything important in Word first (it doesn't touch your documents, but
you'll want to watch), then:

```powershell
python smoke\word_smoke.py
```

What you should see: a Word window opens with a blank document, then — step by
step, with 1-second pauses — a styled "Korean Culture" report types itself in
(real Heading styles, check the outline), text gets find-replaced, a
deliberately stale edit is refused, and three plain text lines are converted
into a real bordered Word table with a bold header row. Midway the script
pauses and asks you to **select the three city lines in Word with your mouse**
and press Enter — that exercises `doc_selection`, the "user highlights text
and says *make this a table*" flow. (Just pressing Enter without selecting
also works; it falls back to locating the lines by content.)

At the end it saves `docd_smoke.docx` and `docd_smoke.pdf` into `%TEMP%` and
prints their paths so you can inspect them.

**Everything the script did is in Word's undo stack — Ctrl+Z steps back
through it.**

### 5. Try the sidecar by hand (optional)

The sidecar is just a process speaking JSON lines — you can drive it from any
terminal:

```powershell
python -m docd
```

then paste (one line at a time):

```json
{"id":"1","method":"ping","params":{}}
{"id":"2","method":"doc_new","params":{}}
{"id":"3","method":"doc_insert","params":{"doc":"w1","text":"# Hello\nWritten via JSON-RPC.","where":"end"}}
{"id":"4","method":"doc_read","params":{"doc":"w1"}}
```

Watch Word obey after each line. `Ctrl+C` (or closing stdin) exits. This is
exactly the interface your app's agent loop will use.

### Troubleshooting

| Symptom | Fix |
|---|---|
| `ModuleNotFoundError: win32com` | `pip install pywin32` (in the same Python you're running) |
| `python` not found | Reinstall Python with *Add to PATH*, or use `py` instead of `python` |
| Calls hang / `APP_BUSY_MODAL` | A dialog box is open in Word — dismiss it and retry |
| Word starts but stays invisible | Kill orphaned `WINWORD.EXE` in Task Manager and rerun |
| Tests pass but smoke fails at `doc_new` | Word isn't installed / not activated for COM — open Word manually once first |

---

## Wiring it into your app

Your app hosts the **agent loop** (any tool-calling LLM: Claude, GPT, or a
local model via the open-computer patterns); the loop declares the `doc_*`
tools and forwards calls to the sidecar:

- **Using pi** (open-computer's harness): register
  `extension/doc-tools.ts` as an extension — same pattern as
  open-computer's `desktop-apps.ts`. Set `DOCD_PYTHON`/`DOCD_CWD` env vars if
  the sidecar isn't at the default location.
- **Using your own loop**: spawn `python -m docd`, keep it alive, send
  `{"id","method","params"}` lines, read `{"id","result"|"error"}` lines. The
  tool schemas to give your LLM are in `extension/doc-tools.ts`.

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

## Design invariants (from the design docs)

- **Direct on the host, not a VM.** open-computer's QEMU layer exists to
  isolate an untrusted agent; here the agent must edit the user's real
  documents. Safety comes from the object model instead: Range-based edits
  (never the user's Selection), Word's undo stack, content-hash staleness
  checks, and approval gates on destructive ops.
- **Range-based, never Selection-based** — doesn't steal the caret or focus
  (`doc_selection` only *reads* the user's selection).
- **Never hide, never quit** the app; attach to a running instance first
  (`GetActiveObject`), launch (`DispatchEx`) only if needed.
- **Stale edits are refused, not clobbered**: every mutating call can carry
  `expect_hash`(es) from the last `doc_read`/`doc_selection`; index drift up
  to ±8 paragraphs is auto-absorbed, content changes raise `STALE_RANGE`.
- **All COM on one STA thread** (`rpc.py`); the stdin reader never touches COM.
- **Modal dialogs** surface as `APP_BUSY_MODAL` after a retry loop, never a
  hang or crash.

## Next slices

- `doc_track_changes` (agent edits as redlined suggestions — recommended
  default for a writing assistant), `doc_comments`, `doc_format`,
  `doc_screenshot_page`
- PowerPoint driver (`slide_*` tools), Hancom HWP driver (field-anchored),
  LibreOffice UNO driver
- Dialog watchdog notifications + UIA sidecar integration
  (see `../docs/windows-writing-assistant/uia-port-design.md`)
