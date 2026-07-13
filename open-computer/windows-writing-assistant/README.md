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

### 5. Run the full AI assistant 🤖 (pi + OpenRouter / Ollama)

This is the complete loop: you chat, the **LLM generates the content and
decides the edits**, and Word obeys. It uses the same pi agent harness as
open-computer, pointed at any OpenAI-compatible endpoint.

```powershell
# one-time: install Node.js (https://nodejs.org), then the pi agent
npm install -g --ignore-scripts @earendil-works/pi-coding-agent

# configure your LLM endpoint
copy agent\.env.example agent\.env
notepad agent\.env    # fill in the OpenRouter block (or Ollama later)

# launch the assistant
python agent\run_agent.py
```

For **OpenRouter** (quick check): get a key at https://openrouter.ai/keys and
pick any tool-calling-capable model (e.g. `anthropic/claude-sonnet-4.5`,
`openai/gpt-4o-mini`). For **Ollama** later: uncomment the Ollama block in
`agent/.env` and use a tool-calling model like `qwen2.5:14b` — nothing else
changes.

Then try your two scenarios in the chat:

- *"Write a one-page report on Korean culture in a new Word document."*
  → watch the LLM `doc_new` + `doc_insert` a styled report into Word.
- Open a document with some list-like text, select it, then:
  *"Convert my selected text into a table."*
  → the LLM reads your selection and swaps it for a real Word table.

`python agent\run_agent.py --dry-run` prints the resolved endpoint/model and
the exact pi command without launching (useful to verify your `.env`).

> How it's wired: `run_agent.py` writes your endpoint into
> `~/.pi/agent/models.json` (same provider schema open-computer generates in
> `interface-service/pi/process.js`), then runs
> `pi --provider writing-assistant --model <id> --extension extension/doc-tools.ts`
> with a writing-assistant system prompt. pi hosts the agent loop; doc-tools.ts
> forwards tool calls to the docd sidecar; docd drives Word over COM.

### 6. Run the desktop app 🪟 (glass chat UI)

A frameless Electron app (electron-vite + React) that replaces the pi
terminal: Win11 acrylic glass chat window, streaming replies, live tool
progress chips ("✍️ Writing…"), in-chat Yes/No cards for save confirmations,
persistent history sidebar, **Alt+Space** to summon/hide from anywhere, and
minimize turns into a small floating orb.

```powershell
# prerequisites: steps 2 + 5 above (pywin32, pi installed, agent\.env filled)
cd app
npm install
npm run dev        # launches the app with hot reload
```

The app spawns `pi --mode rpc` under the hood (same config as
`run_agent.py` — it reads `agent\.env`) and streams its events into the chat.
If pi's RPC event format ever changes, the only file to fix is
`app/src/main/pi-session.ts`.

**Debugging — when the chat does nothing, read these two files:**

- `%APPDATA%\writing-assistant-app\logs\main.log` — everything the app does:
  config resolution, the exact pi spawn command, every prompt sent, every pi
  event received (type-level), pi stderr, exit codes, renderer errors.
- `%APPDATA%\writing-assistant-app\pi-raw.log` — the raw JSON event lines
  from pi, plus its stderr verbatim.

In `npm run dev` the same log lines also stream to the terminal.

### 7. Try the sidecar by hand (optional)

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
| `pi` not found after npm install | Reopen the terminal (PATH refresh), or check `npm config get prefix` is on PATH |
| Agent chats but never edits Word | The model doesn't support tool calling — pick one that does (see `.env.example`) |
| OpenRouter 401 | Wrong/expired `OPENAI_API_KEY` in `agent/.env` |
| Korean text errors (`surrogates not allowed`) or `â€"`-style garbage in documents | Fixed — `git pull`. Root cause: Windows Python decoded the UTF-8 RPC pipe with the ANSI code page; the sidecar now forces UTF-8 stdio |

---

## Wiring it into your app

`agent/run_agent.py` is the reference wiring: pi hosts the agent loop, any
OpenAI-compatible endpoint provides the LLM, `extension/doc-tools.ts` bridges
to the sidecar. To embed in your own app instead, you have two options:

- **Ship pi inside your app** and drive it in RPC mode (`--mode rpc`, JSON
  events over stdio) the way open-computer's interface-service does — you get
  the loop, retries, and session handling for free.
- **Write your own loop**: spawn `python -m docd`, keep it alive, send
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
├── agent/
│   ├── run_agent.py       # launch pi + doc-tools against any OpenAI-compatible LLM
│   └── .env.example       # OpenRouter / Ollama / LM Studio endpoint configs
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
