# COM/UNO Document-Automation Toolset — Design

Windows writing-assistant desktop app, open-computer style agent. Target applications:

| App | Backend API | Driver |
|---|---|---|
| MS Word (docx/doc) | Word COM object model (`Word.Application`) via pywin32 | `WordDriver` |
| Hancom Office Hwp (hwp/hwpx) | `HWPFrame.HwpObject` COM automation via pywin32 (optionally wrapped by **pyhwpx**) | `HwpDriver` |
| LibreOffice Writer (odt) | UNO bridge (pyuno) for live editing; **odfpy** for closed-file manipulation | `OdtDriver` |
| MS PowerPoint (pptx/ppt) | PowerPoint COM object model (`PowerPoint.Application`) via pywin32 | `PptDriver` |
| Hancom Show (show/pptx) | No usable automation OM (see §6.4) — pptx interchange via `PptDriver`/python-pptx | *(routed)* |
| LibreOffice Impress (odp) | UNO draw-page model; **python-pptx** for closed-file pptx | `ImpressDriver` |

Design goals, in order:

1. All reads/edits go through the application **object model** — never synthetic keystrokes or
   pixel clicking — while the real application window stays visible so the user watches the
   document change live.
2. One **unified tool surface** — ~15 `doc_*` tools for text documents plus a dedicated
   9-tool `slide_*` subset for presentations (§6) — whose semantics are identical across
   backends; the agent should not need to know which suite is hosting the document.
3. A **stable addressing scheme** (paragraph indices + content hashes) so the LLM can target
   edits reliably and staleness is detected, mirroring open-computer's a11y harvest-id concept.
4. Safety: approval gates for destructive operations, modal-dialog detection, and honest errors.

This document follows the tool-contract style of
`open-computer/services/extensions/desktop-apps.ts` (`pi.registerTool` + TypeBox schemas +
`execute()` returning `{content:[{type:"text",text}], details}`) and the sidecar-script pattern
those tools use (`a11y-harvest` / `a11y-action` shell-outs). Dialog handling details live in the
sibling doc `uia-port-design.md` (UIA port of the a11y layer); this doc cross-references it
rather than duplicating it.

---

## 1. Architecture

```
┌────────────────────────────── Electron / Node host ──────────────────────────────┐
│  pi agent  ──►  extensions/doc-tools.ts                                          │
│                  • registerTool("doc_read" …) with TypeBox schemas               │
│                  • approval-gate hook (ask_user) for destructive ops             │
│                  • JSON-RPC client over the sidecar's stdio                      │
└──────────────────────────────────┬───────────────────────────────────────────────┘
                                   │ line-delimited JSON-RPC (stdin/stdout)
┌──────────────────────────────────▼───────────────────────────────────────────────┐
│  docd.py — Python automation sidecar (long-running)                              │
│                                                                                  │
│   stdin reader thread ──queue──► COM worker thread (STA, CoInitialize)          │
│                                   │                                              │
│                       DriverRegistry (routes by doc handle / file ext)           │
│         ┌──────────┬─────────┼──────────────┬───────────────┬─────────────┐      │
│    WordDriver  HwpDriver  OdtDriver(UNO)  PptDriver     ImpressDriver     │      │
│    win32com    win32com   pyuno socket    win32com      pyuno draw pages  │      │
│    Word.App    HWPFrame.  soffice         PowerPoint.   (+ python-pptx /  │      │
│                HwpObject  --accept=…      Application   odfpy closed-file)│      │
└──────────────────────────────────────────────────────────────────────────────────┘
```

### 1.1 Why a persistent sidecar (not one-shot scripts like `a11y-harvest`)

`desktop-apps.ts` shells out to one-shot Python scripts per call. That works for AT-SPI because
the bus is stateless. COM is not:

- COM object references (the `Word.Application` proxy, open `Document` objects, the HWP control)
  must live in the **same apartment** across calls; re-dispatching per call costs 200–800 ms and
  loses state (scan positions, undo grouping).
- HWP's `RegisterModule` security handshake should happen once per process.
- The UNO bridge connection setup is expensive (~1 s).

So `docd.py` is a **long-running child process**. The TS extension spawns it lazily on first
`doc_*` call, keeps it alive for the session, and restarts it on crash (all open-doc handles are
then invalid; the agent gets a clear "sidecar restarted, re-open documents" error).

### 1.2 JSON-RPC contract

Line-delimited JSON over stdio (same framing the pi RPC in
`interface-service/pi/process.js` uses for prompts):

```jsonc
// request  (TS ──► docd.py)
{"id": "42", "method": "doc_read", "params": {"doc": "w1", "from_para": 0, "to_para": 40}}
// response (docd.py ──► TS)
{"id": "42", "result": {"paragraphs": [...], "rev": "r7"}}
{"id": "42", "error": {"code": "STALE_RANGE", "message": "...", "data": {...}}}
// unsolicited notification (dialog watchdog, see §7)
{"event": "modal_dialog", "app": "word", "title": "Microsoft Word", "buttons": ["Yes","No","Cancel"]}
```

Error codes are a closed enum the TS layer maps to agent-readable text: `NO_SUCH_DOC`,
`STALE_RANGE`, `READ_ONLY`, `APP_BUSY_MODAL`, `APP_NOT_RUNNING`, `UNSUPPORTED_ON_BACKEND`,
`SAVE_FORMAT_UNSUPPORTED`, `COM_ERROR` (with HRESULT), `TIMEOUT`.

### 1.3 COM apartment & threading

- The COM worker thread calls `pythoncom.CoInitialize()` (STA — the default for
  `CoInitialize`; Word and the HWP control are STA servers, and marshaling their proxies across
  apartments is slow and fragile). **All** COM traffic happens on this one thread.
- The stdin reader thread only parses JSON and enqueues jobs; it never touches COM. Responses are
  written from the worker thread (stdout writes are serialized with a lock).
- If we ever need parallel drivers (e.g. Word busy on a long save while HWP is queried), each
  driver gets its **own** STA thread with its own `CoInitialize()`; COM proxies are never shared
  between threads. Phase 1 uses a single worker thread — simpler, and document edits are
  inherently serial from the agent's point of view.
- UNO (OdtDriver) is not COM; it can live on any thread, but we keep it on the same worker
  thread for uniform timeout/watchdog handling.
- On shutdown: release all COM references (`doc = None; app = None; gc.collect()`) **before**
  `pythoncom.CoUninitialize()`.

### 1.4 Attach vs launch, and keeping the window visible

Rule: **attach to a running instance first, launch only if needed**, and always leave the app
window visible and un-minimized — the user is watching.

- **Word**: `win32com.client.GetActiveObject("Word.Application")` finds a running instance via
  the ROT; on `com_error` fall back to `win32com.client.DispatchEx("Word.Application")`
  (DispatchEx forces a fresh instance; plain `Dispatch` may or may not reuse). After
  attach/launch: `app.Visible = True`. Never toggle `app.Visible = False` — that hides windows
  the user may own.
- **HWP**: the HWP control does not reliably register in the ROT under a stable moniker across
  versions, so default to `Dispatch("HWPFrame.HwpObject")` (which creates/loads the automation
  server) and then make its frame visible: `hwp.XHwpWindows.Item(0).Visible = True`.
  TODO(verify): whether recent Hancom builds (2020+) register a ROT moniker that
  `GetActiveObject` can attach to; if yes, prefer attach.
- **LibreOffice**: soffice must be started with a listening socket to be scriptable:
  `soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo`. If Writer is
  already running *without* the socket, UNO cannot attach — we detect this (connection refused +
  `soffice.bin` process present) and tell the agent/user that LibreOffice must be restarted by
  the sidecar (approval-gated, since it may hold unsaved docs). Documents are loaded with
  `Hidden=False` (default) so the window shows.
- **Live visual feedback**: after every mutating call, the driver scrolls the changed range into
  view (Word: `app.ActiveWindow.ScrollIntoView(rng)`; HWP: cursor movement already scrolls;
  UNO: `doc.CurrentController.ViewCursor.gotoRange(cursor, False)`), so the user literally sees
  the edit land.

---

## 2. Unified tool surface

15 `doc_*` tools for text documents, plus 9 `slide_*`/`pres_*` tools for presentations (§6.1 —
slides are shape-oriented, not paragraph-oriented, so they get their own subset instead of
overloading `doc_*`). All are registered in `extensions/doc-tools.ts` in the `desktop-apps.ts`
style; each `execute()` forwards to the sidecar and formats the result as text. Common
conventions:

- `doc` — the **document handle** returned by `doc_open`/`doc_list_open` (e.g. `"w1"`, `"h2"`,
  `"o1"`, `"p1"` for a presentation; prefix encodes the backend). Every tool takes it.
  `doc_open`, `doc_list_open`, `doc_save`, and `doc_close` are shared across text documents and
  presentations; a presentation handle is used with the `slide_*` tools (`doc_read` etc. on a
  presentation handle returns `UNSUPPORTED_ON_BACKEND` pointing at `slide_read`).
- Addressing (see §2.2): paragraph indices are 0-based, assigned at read time; mutating tools
  accept an `expect_hash` to detect staleness.
- Every mutating tool's result restates the affected paragraphs (index + new hash) so the agent's
  model of the doc stays current without a full re-read.

### 2.1 Tool schemas

```ts
// ─── doc_list_open ───
parameters: Type.Object({});
// → "Open documents:\n  [w1] word  C:\Users\...\report.docx  (dirty, 42 paras, track-changes ON)\n  [h1] hwp  제안서.hwp ..."

// ─── doc_open ───
parameters: Type.Object({
  path: Type.String({ description: "Absolute path to a .docx/.doc/.hwp/.hwpx/.odt/.pptx/.ppt/.odp file. The hosting app is chosen by extension (docx→Word, hwp/hwpx→Hancom, odt→LibreOffice, pptx/ppt→PowerPoint, odp→Impress) unless `app` overrides it. Presentation handles are used with the slide_* tools." }),
  app: Type.Optional(Type.Union([Type.Literal("word"), Type.Literal("hwp"), Type.Literal("libreoffice"), Type.Literal("powerpoint")], { description: "Force a hosting app (e.g. open a .docx in LibreOffice, or a .pptx in Impress via 'libreoffice')." })),
  read_only: Type.Optional(Type.Boolean({ description: "Open without write access (default false)." })),
});
// → handle + doc summary (title, page count, paragraph count, track-changes state)

// ─── doc_read ───
parameters: Type.Object({
  doc: Type.String(),
  from_para: Type.Optional(Type.Number({ description: "First paragraph index (0-based). Omit for whole document." })),
  to_para: Type.Optional(Type.Number({ description: "Last paragraph index, inclusive." })),
  max_chars: Type.Optional(Type.Number({ description: "Truncate output (default 20000). Long docs: read the outline first, then ranges." })),
});
// → markdown-ish text, one block per paragraph:
//    [p0#3fa2] # 2026 사업 계획          ← heading style rendered as #
//    [p1#91cc] 본 문서는 …               ← body text
//    [p2#e001] | 항목 | 금액 |            ← table rows rendered as pipe rows, [t0] table id
//    plus a trailing `rev:` token for the whole snapshot

// ─── doc_outline ───
parameters: Type.Object({ doc: Type.String() });
// → headings tree with paragraph indices: "H1 [p0] 개요 / H2 [p14] 배경 …"

// ─── doc_insert ───
parameters: Type.Object({
  doc: Type.String(),
  text: Type.String({ description: "Text to insert. '\\n' starts a new paragraph. Markdown headings (#, ##) map to heading styles when style_map is true." }),
  where: Type.Union([
    Type.Literal("end"), Type.Literal("cursor"),
    Type.Literal("before_para"), Type.Literal("after_para"),
    Type.Literal("bookmark"),
  ], { description: "Anchor. 'bookmark' uses `bookmark` (Word bookmark / HWP field(필드) / ODT bookmark)." }),
  para: Type.Optional(Type.Number({ description: "Paragraph index for before_para/after_para." })),
  expect_hash: Type.Optional(Type.String({ description: "Hash of the anchor paragraph from the last doc_read; edit is refused if it no longer matches." })),
  bookmark: Type.Optional(Type.String()),
  style_map: Type.Optional(Type.Boolean({ description: "Map markdown '#'-prefixes to Heading styles (default true)." })),
});

// ─── doc_replace ───
parameters: Type.Object({
  doc: Type.String(),
  find: Type.String(),
  replace: Type.String(),
  regex: Type.Optional(Type.Boolean({ description: "Treat `find` as a regex (backend-native where possible, see notes; default false = literal)." })),
  match_case: Type.Optional(Type.Boolean()),
  occurrence: Type.Optional(Type.Union([Type.Literal("all"), Type.Literal("first"), Type.Number()], { description: "'all' (default), 'first', or the 1-based Nth occurrence." })),
  scope: Type.Optional(Type.Object({ from_para: Type.Number(), to_para: Type.Number() }, { description: "Restrict to a paragraph range." })),
});
// → "Replaced 3 occurrence(s). Affected: [p4#a1b2] [p9#77de] [p31#0c0c]"

// ─── doc_edit_range ───  (surgical rewrite of a paragraph range)
parameters: Type.Object({
  doc: Type.String(),
  from_para: Type.Number(),
  to_para: Type.Number(),
  expect_hashes: Type.Array(Type.String(), { description: "Hashes of every paragraph in the range, from the last doc_read. All must match or the edit is refused as stale." }),
  new_text: Type.String({ description: "Replacement text; '\\n' separates paragraphs. Empty string deletes the range." }),
});

// ─── doc_apply_style ───
parameters: Type.Object({
  doc: Type.String(),
  from_para: Type.Number(),
  to_para: Type.Optional(Type.Number()),
  style: Type.String({ description: "Paragraph style name, normalized: 'Heading 1'…'Heading 9', 'Normal', 'Title', 'Quote', 'List Bullet', 'List Number'. Backend maps to native names (e.g. HWP '개요 1', ODT 'Heading 1')." }),
});

// ─── doc_format ───  (character formatting on a text span)
parameters: Type.Object({
  doc: Type.String(),
  target: Type.Union([
    Type.Object({ para: Type.Number(), start: Type.Optional(Type.Number()), end: Type.Optional(Type.Number()) }),
    Type.Object({ find: Type.String(), occurrence: Type.Optional(Type.Number()) }),
  ], { description: "Either a character span within a paragraph, or the Nth match of a search string." }),
  bold: Type.Optional(Type.Boolean()), italic: Type.Optional(Type.Boolean()),
  underline: Type.Optional(Type.Boolean()),
  font: Type.Optional(Type.String()), size_pt: Type.Optional(Type.Number()),
  color: Type.Optional(Type.String({ description: "Hex RGB like '#CC0000'." })),
});

// ─── doc_comments ───
parameters: Type.Object({
  doc: Type.String(),
  op: Type.Union([Type.Literal("list"), Type.Literal("add"), Type.Literal("reply"), Type.Literal("resolve"), Type.Literal("delete")]),
  comment_id: Type.Optional(Type.String({ description: "From a previous list (e.g. 'c3'). Required for reply/resolve/delete." })),
  text: Type.Optional(Type.String({ description: "Comment body for add/reply." })),
  anchor: Type.Optional(Type.Object({ para: Type.Number(), find: Type.Optional(Type.String()) }, { description: "For add: paragraph (optionally a substring within it) to attach the comment to." })),
});

// ─── doc_track_changes ───
parameters: Type.Object({
  doc: Type.String(),
  op: Type.Union([Type.Literal("status"), Type.Literal("enable"), Type.Literal("disable"),
                  Type.Literal("list"), Type.Literal("accept"), Type.Literal("reject")]),
  revision_id: Type.Optional(Type.String({ description: "From list; omit with accept/reject to act on ALL (approval-gated)." })),
});

// ─── doc_tables ───
parameters: Type.Object({
  doc: Type.String(),
  op: Type.Union([Type.Literal("list"), Type.Literal("read"), Type.Literal("write")]),
  table: Type.Optional(Type.String({ description: "Table id from doc_read/list, e.g. 't0'." })),
  cell: Type.Optional(Type.Object({ row: Type.Number(), col: Type.Number() }, { description: "0-based. For write." })),
  value: Type.Optional(Type.String()),
  values: Type.Optional(Type.Array(Type.Array(Type.String()), { description: "2-D block write starting at `cell`." })),
});

// ─── doc_save ───
parameters: Type.Object({ doc: Type.String() });   // approval-gated: overwrites the file

// ─── doc_save_as ───
parameters: Type.Object({
  doc: Type.String(),
  path: Type.String(),
  format: Type.Union([Type.Literal("docx"), Type.Literal("doc"), Type.Literal("hwp"),
                      Type.Literal("hwpx"), Type.Literal("odt"), Type.Literal("pdf"),
                      Type.Literal("txt")]),
});

// ─── doc_screenshot_page ───  (optional visual check; same "only when needed" rule as app_screenshot)
parameters: Type.Object({
  doc: Type.String(),
  page: Type.Optional(Type.Number({ description: "1-based page (default: the page containing the cursor/last edit)." })),
});
// → base64 PNG content block, captured from the live window (or PDF-rendered page as fallback)

// ─── doc_close ───
parameters: Type.Object({
  doc: Type.String(),
  discard_changes: Type.Optional(Type.Boolean({ description: "Close without saving (approval-gated when the doc is dirty)." })),
});
```

**Format conversion matrix** for `doc_save_as` (✓ native, ⚠ lossy/partial, ✗ unsupported →
sidecar returns `SAVE_FORMAT_UNSUPPORTED` with a suggested route):

| host \ target | docx | doc | hwp | hwpx | odt | pdf | txt |
|---|---|---|---|---|---|---|---|
| Word | ✓ | ✓ | ✗ (route: open in HWP¹) | ✗ | ✓ (`wdFormatOpenDocumentText`, ⚠ fidelity) | ✓ | ✓ |
| HWP | ⚠ TODO(verify)² | ⚠ | ✓ | ✓ | ✗ | ✓ | ✓ |
| LibreOffice | ✓ (`MS Word 2007 XML`, ⚠) | ✓ (`MS Word 97`) | ✗³ | ✗ | ✓ (`writer8`) | ✓ (`writer_pdf_Export`) | ✓ (`Text`) |

¹ Hancom Office can *open* docx; the routed conversion is: save docx from Word → `doc_open` in
HWP → `doc_save_as` hwp. The sidecar suggests this route in the error, it does not do it
implicitly.
² Hancom advertises MS Word export (`SaveAs(..., "MSWORD"?)`); exact FileFormat string differs
by version — TODO(verify) against the installed Hancom automation reference.
³ LibreOffice ships a **read** filter for legacy HWP 5.x (`writer_MIZI_Hwp_97`) but no writer.

Presentation formats (pptx / ppt / odp / show / pdf / png) have their own matrix under
`pres_save_as` in §6.4.

### 2.2 Stable addressing: paragraph indices + content hashes

The equivalent of open-computer's harvest ids. Problem: paragraph indices shift as soon as the
user (or the agent) inserts/deletes paragraphs, and the user can type at any time — the doc is a
shared mutable resource.

- `doc_read` assigns 0-based indices in document order and computes, per paragraph,
  `hash = sha1(normalize(text))[:4]` (normalize = NFC, strip trailing whitespace/para mark). The
  read also returns a whole-snapshot `rev` token: for Word `doc.Content.End` + save-count; for
  HWP a hash of the scan output; for UNO the modified-state + paragraph count.
- Mutating tools take `expect_hash`/`expect_hashes`. The driver re-resolves the paragraph at the
  given index **at execute time** and compares hashes. Mismatch → `STALE_RANGE` error listing the
  current text of ±2 paragraphs around the index, so the agent can usually re-anchor without a
  full re-read.
- Hashes are content-based, not position-based, so if paragraphs merely shifted (insert above),
  the driver **searches ±8 indices** for a hash match before failing, and reports the corrected
  index in the result ("anchor moved p14→p16"). This absorbs the common "user typed a line above"
  case.
- Backend mapping of "paragraph i":
  - Word: `doc.Paragraphs(i+1).Range` (1-based COM collection). Cheap random access.
  - UNO: text-enumeration order (tables count as one enumeration element; their paragraphs are
    addressed via `doc_tables`).
  - HWP: no cheap random access — the driver keeps a per-read cache of `(para_index → set-pos
    tuple)` from the last scan (`GetPos`-style position captured during `InitScan`). This cache
    is exactly what goes stale, hence §4's strong recommendation to anchor HWP edits on
    **fields** or **find**, not raw indices.

Presentations use the analogous scheme — slide index + shape id + text hash — defined in §6.2.

### 2.3 Extension-side execute() shape

Identical to `desktop-apps.ts`; one representative example:

```ts
pi.registerTool({
  name: "doc_replace",
  label: "Find & Replace in Document",
  description: "Find and replace text in an open document (Word/HWP/LibreOffice). " +
    "Runs through the app's own object model; the user sees the change live. " +
    "Prefer this over doc_edit_range for small textual changes.",
  parameters: /* schema above */,
  async execute(_id, params, signal, _onUpdate, ctx) {
    try {
      const res = await sidecar.call("doc_replace", params, { timeoutMs: 30_000, signal });
      return { content: [{ type: "text", text: formatReplaceResult(res) }], details: res };
    } catch (err) {
      return { content: [{ type: "text", text: sidecarErrorToText(err) }], details: {} };
    }
  },
});
```

---

## 3. WordDriver specifics

Everything is `Range`-based, never `Selection`-based, for two reasons: `Selection` steals the
user's caret (they're watching, maybe even holding the mouse), and Range edits don't depend on
window focus. Range-based edits still repaint live.

Core object-model calls:

- **Open/attach**: `app = GetActiveObject("Word.Application")` → fallback
  `DispatchEx("Word.Application")`; `app.Visible = True`.
  `doc = app.Documents.Open(FileName=path, ReadOnly=ro, AddToRecentFiles=False,
  ConfirmConversions=False)`. Enumerate open docs via `app.Documents` and match
  `doc.FullName`.
- **Read**: `doc.Paragraphs.Count`; `p = doc.Paragraphs(i)`; `p.Range.Text` (strip trailing
  `\r`); `p.Style.NameLocal` and `p.OutlineLevel` (1–9 vs `wdOutlineLevelBodyText=10`) for the
  outline. Tables: `doc.Tables(j).Cell(r,c).Range.Text` — strip the trailing `"\r\x07"` cell
  marker. A paragraph inside a table has `p.Range.Information(wdWithInTable)` = True; doc_read
  renders the whole table as pipe rows at the position of its first paragraph.
- **Find/replace**: on a Range so scope control is exact:

  `rng.Find.Execute(FindText=..., MatchCase=..., MatchWholeWord=False, MatchWildcards=...,
  Forward=True, Wrap=wdFindStop, ReplaceWith=..., Replace=wdReplaceOne|wdReplaceAll)`

  `Wrap=wdFindStop (0)` is essential — `wdFindContinue` escapes the scoped range. For
  `occurrence=N`, loop `Execute(Replace=wdReplaceNone)` collapsing `rng.Start = found.End` until
  the Nth hit, then replace that Range. `regex:true` maps to `MatchWildcards=True` **only for
  the wildcard-safe subset**; otherwise the driver falls back to reading paragraph text, running
  Python `re`, and rewriting via `doc_edit_range` mechanics (documented in the tool description
  so the agent knows regex may be emulated).
- **Insert**: `anchor = doc.Paragraphs(i+1).Range`; `anchor.Collapse(wdCollapseEnd)` /
  `wdCollapseStart`; then `anchor.InsertAfter(text)` or `anchor.Text = text` for in-place
  replacement. New paragraphs: include `"\r"`; then style them:
  `doc.Range(insStart, insEnd).Paragraphs(k).Style = doc.Styles("Heading 2")`. Bookmark anchor:
  `doc.Bookmarks("name").Range`; `doc.Bookmarks.Add("name", rng)` to create (note: inserting *at*
  a bookmark can delete it — re-add after insert).
- **Styles/format**: `rng.Style = doc.Styles(styleName)` (paragraph), `rng.Bold = True`,
  `rng.Italic`, `rng.Font.Name` / `.NameFarEast` (set **both** for Korean text so Hangul gets
  the intended font), `rng.Font.Size`, `rng.Font.Color = RGB`.
- **Comments**: `c = doc.Comments.Add(Range=rng, Text=body)`; list via `doc.Comments` (`c.Index`,
  `c.Author`, `c.Range.Text`, `c.Scope.Text`, `c.Done`). Resolve: `c.Done = True` (Word 2013+).
  Replies: modern Word threads expose `c.Ancestor`; there is no clean public `Replies.Add` in
  the COM OM — TODO(verify) `Comment.Replies` availability per Word version; fallback is adding
  a comment on the same `Scope` range, which Word threads visually.
- **Track changes**: `doc.TrackRevisions = True/False`; list `doc.Revisions` (`r.Type`,
  `r.Author`, `r.Range.Text`); `r.Accept()` / `r.Reject()`; all:
  `doc.Revisions.AcceptAll()` / `.RejectAll()` (approval-gated).
- **Save**: `doc.Save()`; `doc.SaveAs2(FileName, FileFormat=...)` with
  `wdFormatXMLDocument=12` (docx), `wdFormatDocumentDefault=16`, `wdFormatPDF=17`,
  `wdFormatOpenDocumentText=23`, `wdFormatDocument97=0` (doc), `wdFormatText=2`.
- **UI liveness**: leave `app.ScreenUpdating = True` in normal operation — the whole point is
  that the user watches. For bulk operations (replace-all across a 200-page doc) temporarily set
  `ScreenUpdating = False` and restore in a `finally`, then `ScrollIntoView` the last change.
  `app.DisplayAlerts = wdAlertsNone (0)` while the sidecar operates, restored on idle, to reduce
  modal prompts; file-conflict and macro dialogs can still appear (§8).
- **Modal dialogs**: while a modal dialog is up, incoming COM calls fail with
  `RPC_E_CALL_REJECTED (0x80010001)` or `RPC_E_SERVERCALL_RETRYLATER (0x8001010A)`. The driver
  wraps every call in a retry loop (5 attempts, 300 ms backoff); persistent rejection triggers
  the dialog probe (§8) and an `APP_BUSY_MODAL` error naming the dialog.
- **Screenshot**: preferred — capture the Word window rect (via `win32gui.GetWindowRect` +
  `PrintWindow`); fallback — `doc.ExportAsFixedFormat(..., Range=wdExportFromTo, From=p, To=p)`
  to a temp PDF and rasterize page 1.

---

## 4. HwpDriver specifics

The HWP automation model is **action-based**: instead of a rich object graph you mostly execute
named actions (`HAction`) with parameter sets (`HParameterSet`), operating at the current cursor
position. This makes position-based editing inherently flakier than Word — plan accordingly.

- **Dispatch & security module**: 
  `hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")`, then immediately
  `hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")` to suppress the per-file security
  popup ("이 파일에 접근을 허용하시겠습니까?"). This only works if a security-check DLL is
  registered in the registry first: a value under
  `HKCU\Software\HNC\HwpAutomation\Modules` (older builds:
  `HKCU\Software\HNC\HwpCtrl\Modules`) whose name is the module name passed as the second
  argument (`"SecurityModule"`) and whose data is the absolute path to the DLL (Hancom's sample
  `FilePathCheckerModuleExample.dll` from the automation SDK, or a vendored equivalent).
  TODO(verify): exact registry branch per Hancom Office version (2018/2020/2022/2024) — pyhwpx
  handles this registration automatically and is the reference implementation to crib from.
  The installer for our app should write this key at install time; the sidecar re-checks and
  self-heals it at startup.
- **pyhwpx**: a maintained Python wrapper over `HWPFrame.HwpObject` that covers dispatch,
  security-module registration, insert/find-replace/field/table helpers. Recommendation: depend
  on pyhwpx for Phase 2 speed, but keep our `HwpDriver` interface thin over it so we can drop to
  raw `HAction` calls where pyhwpx is incomplete.
- **Visibility**: `hwp.XHwpWindows.Item(0).Visible = True` right after dispatch (the automation
  server starts hidden).
- **Open/save**: `hwp.Open(path, "HWP", "")` (format arg may be `""` to autodetect);
  `hwp.Save()`; `hwp.SaveAs(path, Format, arg)` with Format strings `"HWP"`, `"HWPX"`, `"PDF"`
  (TODO(verify): `"HWPX"` string on pre-2020 builds — older automation exposed `"HWPML2X"` for
  the XML format lineage; and the exact string for MS Word export).
- **Insert text** (at cursor): the canonical action triple —

  ```python
  pset = hwp.HParameterSet.HInsertText
  hwp.HAction.GetDefault("InsertText", pset.HSet)
  pset.Text = text            # BSTR — Hangul-safe, no IME involved
  hwp.HAction.Execute("InsertText", pset.HSet)
  ```

  Paragraph breaks are inserted with the `"BreakPara"` action between lines (embedding `"\r\n"`
  in `Text` behaves inconsistently across versions — TODO(verify)).
- **Cursor movement**: `hwp.MovePos(moveID, para, pos)` — `moveTopOfFile=2`,
  `moveBottomOfFile=3` for begin/end anchors (TODO(verify) full moveID table against the
  automation chm); `hwp.GetPos()` / `hwp.SetPos(list, para, pos)` to capture and restore
  positions — these tuples are what the paragraph-index cache (§2.2) stores, and they are
  invalidated by any edit above them. Hence:
- **Fields (필드/누름틀) are the reliable anchor**. `hwp.PutFieldText("name", text)` writes into a
  named field anywhere in the doc without cursor math; `hwp.GetFieldText("name")` reads it;
  `hwp.GetFieldList(0, 0x01)` enumerates field names (returns a `\x02`-separated string);
  `hwp.CreateField(direction, memo, name)` creates a click-here field at the cursor
  (TODO(verify) argument order — pyhwpx `create_field` wraps it). Driver policy: when the agent
  asks to insert at `bookmark` on an HWP doc, we resolve HWP **fields** as the bookmark
  namespace; `doc_insert` can also *create* a field at the insertion point so subsequent edits
  have a durable anchor. Templates authored with named fields are the golden path for
  form-filling workflows.
- **Find/replace**: parameter-set driven, no dialog shown when `IgnoreMessage` is set:

  ```python
  pset = hwp.HParameterSet.HFindReplace
  hwp.HAction.GetDefault("AllReplace", pset.HSet)   # or "ExecReplace" (one) / "RepeatFind"
  pset.FindString, pset.ReplaceString = find, replace
  pset.IgnoreMessage = 1        # suppress the "N개를 바꾸었습니다" popup
  pset.MatchCase = 1 if match_case else 0
  hwp.HAction.Execute("AllReplace", pset.HSet)
  ```

  Occurrence targeting: loop `"RepeatFind"` to position the cursor at the Nth hit, then
  `"ExecReplace"` once. Regex: HWP find has its own limited pattern options
  (`pset.FindRegExp` — TODO(verify) exact member name); default to the read-modify-write
  emulation path for `regex:true`.
- **Text extraction** (doc_read): the scan loop —

  ```python
  hwp.InitScan(option=0, Range=0x0077)   # scan whole doc incl. controls; TODO(verify) flags
  paras = []
  while True:
      state, text = hwp.GetText()        # returns (state, text)
      if state in (0, 1): break          # 0=end, 1=error/empty
      paras.append(text)
  hwp.ReleaseScan()
  ```

  Heading detection: paragraph shape / outline level is not part of the scan output; the driver
  reads style via `hwp.GetCurFieldName`-adjacent APIs is unreliable — instead use
  `ParaShape`/outline actions per paragraph (expensive) or accept style-less outline for HWP in
  Phase 2, improving later. TODO(verify): `hwp.GetTextFile("TEXT"| "HWPML2X", "saveblock")` as a
  faster whole-doc extraction path, with HWPML parsed for structure — pyhwpx uses this trick.
- **Tables**: navigate into a table via find/field anchors, then `hwp.HAction` table actions, or
  address cells by putting **fields inside cells** (best), or via `GetTextFile("HWPML2X")`
  parsing for reads. Honest assessment: HWP table write без fields is the flakiest operation in
  this whole design; doc_tables on HWP ships in "read + field-cell write" form first.
- **Flakiness policy (be honest with the agent)**: the HWP driver's tool results and the system
  prompt note that on HWP, paragraph-index edits may be refused as stale more often, and that
  `doc_replace` (find-anchored) and field-based `doc_insert` are strongly preferred. The
  `expect_hash` machinery still applies — the scan is re-run over the target neighborhood before
  every index-based edit.

---

## 5. OdtDriver specifics

### 5.1 Live editing via UNO

- **Launch/connect**: sidecar starts
  `soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo` (visible, not
  `--headless` — the user watches) and connects:

  ```python
  localCtx = uno.getComponentContext()
  resolver = localCtx.ServiceManager.createInstanceWithContext(
      "com.sun.star.bridge.UnoUrlResolver", localCtx)
  ctx = resolver.resolve(
      "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
  desktop = ctx.ServiceManager.createInstanceWithContext(
      "com.sun.star.frame.Desktop", ctx)
  doc = desktop.loadComponentFromURL(to_file_url(path), "_blank", 0, ())
  ```

  Python is the bundled LibreOffice python (its pyuno is version-locked) — the sidecar shells
  UNO work into `instdir/program/python.exe`, or we vendor a matching pyuno; decide in Phase 3.
- **Read**: `text = doc.getText()`; `enum = text.createEnumeration()`; each element supporting
  `com.sun.star.text.Paragraph` yields `el.getString()` and `el.ParaStyleName` (headings are
  `"Heading 1"…`); elements supporting `com.sun.star.text.TextTable` are tables
  (`tbl.getCellByName("A1").getString()`, `tbl.Rows.Count`, `tbl.Columns.Count`).
- **Insert/edit**: `cur = text.createTextCursorByRange(par.getStart())`;
  `text.insertString(cur, s, False)`;
  `text.insertControlCharacter(cur, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)`;
  replace a paragraph by `cur.gotoStartOfParagraph`… span selection then
  `cur.setString(new_text)`. Character formatting: cursor properties `CharWeight`
  (`com.sun.star.awt.FontWeight.BOLD`), `CharPosture`, `CharHeight`, `CharFontName`, `CharColor`.
- **Find/replace**: native and regex-capable —

  ```python
  rd = doc.createReplaceDescriptor()          # XReplaceable
  rd.setSearchString(find); rd.setReplaceString(replace)
  rd.SearchCaseSensitive = match_case
  rd.SearchRegularExpression = use_regex      # ICU regex — real regex support
  n = doc.replaceAll(rd)
  ```

  Occurrence targeting via `doc.findFirst(sd)` / `doc.findNext(range, sd)` with a
  `createSearchDescriptor()`, then `found.setString(replace)`.
- **Comments**: `ann = doc.createInstance("com.sun.star.text.textfield.Annotation")`;
  `ann.Author`, `ann.Content`; `text.insertTextContent(cursor, ann, False)`; enumerate via
  `doc.getTextFields().createEnumeration()`. Resolve flag: `ann.Resolved` (LO 6.4+,
  TODO(verify)).
- **Track changes**: `doc.setPropertyValue("RecordChanges", True)`; list via the
  `com.sun.star.document.RedlinesSupplier` interface (`doc.getRedlines()`); accept/reject all via
  DispatchHelper: `.uno:AcceptAllTrackedChanges` / `.uno:RejectAllTrackedChanges` on
  `doc.getCurrentController().getFrame()`.
- **Save**: `doc.store()`; `doc.storeToURL(url, props)` with `FilterName` `PropertyValue`:
  `"writer8"` (odt), `"MS Word 2007 XML"` (docx), `"MS Word 97"` (doc),
  `"writer_pdf_Export"` (pdf), `"Text"` (txt).
- **Scroll-into-view**: `doc.getCurrentController().getViewCursor().gotoRange(cur, False)`.

### 5.2 odfpy fallback (closed files)

When the user hasn't got the doc open and no live view is needed (batch conversion, quick text
extraction), `OdtDriver` can operate on the closed file with **odfpy**
(`odf.opendocument.load`, walk `text:p`/`text:h` elements, `teletype.extractText`, save). Rules:
never touch a file with odfpy while the same file is open in soffice (lock-file check:
`.~lock.<name>#`), and doc handles are tagged `o1(closed)` so the agent knows there is no live
window. doc_read/doc_replace/doc_insert/doc_save_as(odt only) are supported in closed mode;
styles/comments/track-changes are UNO-only.

---

## 6. Presentation automation (PowerPoint / Hancom Show / Impress)

Slides are canvases of **shapes**, not paragraph streams, so presentations get a dedicated
tool subset instead of shoehorning them into `doc_*`. The architecture is unchanged: the same
sidecar, same JSON-RPC contract, same STA worker thread; `PptDriver` and `ImpressDriver` are
just two more entries in the `DriverRegistry`, and `doc_open`/`doc_save`/`doc_close` are shared.

### 6.1 The `slide_*` tool subset

```ts
// ─── slide_list ───  (the presentation outline — cheap, call this first)
parameters: Type.Object({ doc: Type.String() });
// → "12 slides:\n  [s0] layout='Title Slide'  \"2026 사업 계획\"\n  [s1] layout='Title and Content' \"추진 배경\" (notes: yes) …"

// ─── slide_read ───  (all text-bearing shapes on one slide, plus notes)
parameters: Type.Object({
  doc: Type.String(),
  slide: Type.Number({ description: "0-based slide index from slide_list." }),
});
// → per shape: [s3/sh5#a1c9] placeholder=title "추진 배경"
//              [s3/sh7#2e00] body • bullet 1 • bullet 2   (bullets rendered as '•' lines)
//              notes: "발표 시 강조할 것: …"

// ─── slide_add ───
parameters: Type.Object({
  doc: Type.String(),
  after_slide: Type.Optional(Type.Number({ description: "Insert after this index; omit to append at end." })),
  layout: Type.Optional(Type.String({ description: "Layout name from the template, e.g. 'Title and Content'. Default: the deck's most common body layout." })),
  title: Type.Optional(Type.String()),
  body: Type.Optional(Type.String({ description: "Body text; '\\n' = new bullet, leading tabs = indent level." })),
});

// ─── slide_edit_text ───  (rewrite one shape's text)
parameters: Type.Object({
  doc: Type.String(),
  slide: Type.Number(),
  shape: Type.String({ description: "Shape id from slide_read, e.g. 'sh5'." }),
  expect_hash: Type.Optional(Type.String({ description: "Hash from slide_read; refused as stale on mismatch." })),
  text: Type.String({ description: "New text. '\\n' = new paragraph/bullet, leading tabs = indent level. Empty string clears the shape." }),
});

// ─── slide_notes_edit ───
parameters: Type.Object({
  doc: Type.String(), slide: Type.Number(),
  text: Type.String({ description: "Replaces the speaker notes for the slide." }),
});

// ─── slide_reorder ───
parameters: Type.Object({
  doc: Type.String(), slide: Type.Number(),
  to_index: Type.Number({ description: "New 0-based position." }),
});

// ─── slide_delete ───   (approval-gated)
parameters: Type.Object({ doc: Type.String(), slide: Type.Number(),
  expect_hash: Type.Optional(Type.String({ description: "Hash of the slide title/first shape, as a safety check." })) });
// slide_add with `duplicate_of: Type.Optional(Type.Number())` covers duplication.

// ─── slide_thumbnail ───  (slides are inherently visual — this is the doc_screenshot_page analog,
//                           and unlike text docs it IS routinely worth calling after layout edits)
parameters: Type.Object({
  doc: Type.String(), slide: Type.Number(),
  width_px: Type.Optional(Type.Number({ description: "Default 960." })),
});
// → base64 PNG content block of the rendered slide

// ─── pres_save_as ───
parameters: Type.Object({
  doc: Type.String(), path: Type.String(),
  format: Type.Union([Type.Literal("pptx"), Type.Literal("ppt"), Type.Literal("odp"),
                      Type.Literal("pdf"), Type.Literal("png")],
    { description: "'png' exports every slide as an image into a directory." }),
});
```

Also under `pres_*` semantics: `slide_add`/`slide_edit_text` results echo the affected shapes
(id + new hash), and every mutating call makes the edited slide the current view
(PowerPoint: `app.ActiveWindow.View.GotoSlide(i+1)`; Impress: set the controller's
`CurrentPage`) so the user watches the deck change live.

### 6.2 Stable addressing for slides

Consistent with §2.2 but two-level:

- **Slide**: 0-based index at read time. Slides shift on insert/delete/reorder; each slide also
  carries a hash of its title + shape-text concatenation, and mutating calls that carry
  `expect_hash` get the ±3-slide re-anchoring search before failing `STALE_RANGE`.
- **Shape**: PowerPoint gives every shape a stable integer `Shape.Id` (survives reorder and
  most edits) — `sh5` maps to `Shape.Id == 5`, resolved via iterating `slide.Shapes`; this is
  the strongest anchor in the whole design. UNO/Impress shapes have **no** stable id — the
  driver uses `shape.Name` when non-empty, else `index@slide` with the text hash as the real
  identity check. python-pptx (closed files) exposes `shape.shape_id` — same semantics as COM.
- Shape text hash: `sha1(NFC(all runs' text))[:4]`, same normalization as paragraphs.

### 6.3 PptDriver — PowerPoint COM specifics

- **Attach/launch**: `GetActiveObject("PowerPoint.Application")` → fallback
  `Dispatch("PowerPoint.Application")` (PowerPoint is single-instance; both usually land on the
  same process). `app.Visible = True`. Quirk: PowerPoint has **no true headless mode** —
  setting `Visible = False` on the application raises an error, and
  `Presentations.Open(..., WithWindow=msoFalse)` merely opens the presentation without a
  document window while the app process stays visible. For us this is a feature, not a bug —
  the user is watching — so we always open `WithWindow=msoTrue`.
- **Open/enumerate**: `pres = app.Presentations.Open(FileName=path, ReadOnly=ro, Untitled=False,
  WithWindow=True)`; `pres.Slides.Count`; `sl = pres.Slides(i+1)` (1-based); `sl.SlideIndex`.
- **Add/layout**: modern API — `pres.Slides.AddSlide(Index,
  pres.SlideMaster.CustomLayouts(j))`, picking the `CustomLayout` whose `.Name` matches the
  requested layout; legacy `Slides.Add(Index, Layout=ppLayoutText)` as fallback for old
  templates. Duplicate: `sl.Duplicate()`; reorder: `sl.MoveTo(toPos)`; delete: `sl.Delete()`.
- **Shapes/text**: iterate `sl.Shapes`; a shape carries text iff `shape.HasTextFrame` and
  `shape.TextFrame.HasText`; text via `shape.TextFrame.TextRange.Text` (paragraphs:
  `TextRange.Paragraphs(k)`, indent level `ParagraphFormat.IndentLevel`). Placeholders:
  `shape.Type == msoPlaceholder (14)` + `shape.PlaceholderFormat.Type`
  (`ppPlaceholderTitle=1`, `ppPlaceholderBody=2`); title shortcut `sl.Shapes.Title`. Writes set
  `TextRange.Text = ...` (Hangul-safe BSTR, per §7) and per-paragraph `IndentLevel` for bullet
  depth. Character formatting reuses `doc_format` semantics on `TextRange.Font`
  (`.Bold/.Italic/.Name/.NameFarEast/.Size/.Color.RGB`).
- **Notes**: `sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text` (placeholder 2 is
  the notes body; guard for decks whose notes master was customized).
- **Theme/template**: `pres.ApplyTemplate(path_to_potx_or_pptx)`;
  `pres.ApplyTheme(path_to_thmx)` (2007+; also per-slide `sl.ApplyTheme`). Exposed to the agent
  only via an optional `template` param at `doc_open` time or a post-Phase-2 follow-up tool —
  wholesale re-theming is visually drastic, so it is approval-gated.
- **Save/export**: `pres.SaveAs(path, FileFormat=ppSaveAsOpenXMLPresentation=24)`;
  `ppSaveAsPDF=32`; legacy `ppSaveAsPresentation=1` (.ppt). Thumbnails:
  `sl.Export(png_path, "PNG", ScaleWidth, ScaleHeight)` per slide — this powers
  `slide_thumbnail`; whole-deck `pres.Export(dir, "PNG", w, h)` powers `pres_save_as png`.
  ODP: `ppSaveAsOpenDocumentPresentation=35` (⚠ fidelity).
- Same modal-dialog retry/watchdog machinery as Word (§8); `app.DisplayAlerts =
  ppAlertsNone (1)` while operating. TODO(verify): `ppAlertsNone` constant value.

### 6.4 Hancom Show and Impress backends

- **Hancom Show** (`.show`, Hancom's PowerPoint counterpart): unlike HWP, Show's COM automation
  surface is thin, near-undocumented, and version-unstable — there is no HAction-grade public
  API to build on (TODO(verify): whether current Hancom Office exposes a scriptable
  `ShowFrame.*` object at all; treat any claim as unverified until tested against a licensed
  install). **Pragmatic path**: Hancom Show reads and writes pptx, so the sidecar routes Show
  workflows through pptx interchange — edit with `PptDriver` (if PowerPoint is installed) or
  **python-pptx** closed-file editing (no Office needed), then let the user open the result in
  Show. `.show` files themselves: ask the user (via Show, manually or one UIA-driven Save-As —
  see `uia-port-design.md`) to convert to pptx first; the tool error message explains this.
- **Impress (UNO)**: same connection as §5.1; an Impress document implements
  `com.sun.star.drawing.XDrawPagesSupplier` — `pages = doc.getDrawPages()`;
  `page = pages.getByIndex(i)`; shapes via `page.getCount()` / `page.getByIndex(k)`; text on
  shapes supporting `com.sun.star.drawing.Text` (`shape.getString()`, or a shape-text cursor
  for runs); placeholders identified by service names `com.sun.star.presentation.TitleTextShape`
  / `OutlineTextShape` / `NotesTextShape`. Insert slide: `pages.insertNewByIndex(i)` + apply
  layout via the page's `Layout` property. Notes: each draw page implements
  `com.sun.star.presentation.XPresentationPage` → `page.getNotesPage()`. Save filters:
  `impress8` (odp), `"Impress MS PowerPoint 2007 XML"` (pptx), `impress_pdf_Export` (pdf);
  thumbnails via the `com.sun.star.drawing.GraphicExportFilter` service targeted at one page
  (`MediaType image/png`).
- **python-pptx (closed-file)**: full pptx read/edit without any office app — slides,
  placeholders, text frames, notes (`slide.notes_slide.notes_text_frame`), reorder via XML
  `sldIdLst` manipulation (⚠ not first-class in the library — TODO(verify) stability), but **no
  rendering**, so `slide_thumbnail` is unavailable in closed mode (fallback: LibreOffice
  headless `--convert-to png`). Handles are tagged `p2(closed)` like odfpy's.

**Presentation conversion matrix** (`pres_save_as`):

| host \ target | pptx | ppt | show | odp | pdf | png |
|---|---|---|---|---|---|---|
| PowerPoint | ✓ (24) | ✓ (1) | ✗ (route: open pptx in Show) | ⚠ (35) | ✓ (32) | ✓ (`Export`) |
| Impress | ✓ (`Impress MS PowerPoint 2007 XML`, ⚠) | ⚠ (`MS PowerPoint 97`) | ✗ | ✓ (`impress8`) | ✓ | ✓ (per-page graphic export) |
| python-pptx (closed) | ✓ | ✗ | ✗ | ✗ | ✗ (route via LO) | ✗ (route via LO) |

---

## 7. Korean text / IME note

Every insertion in this design goes through object-model string parameters — COM `BSTR`
(UTF-16) for Word/HWP, UNO `string` for LibreOffice. **No synthetic keystrokes are ever sent**,
so the Windows IME is completely bypassed: Hangul (and any Unicode) arrives intact regardless of
the user's current IME composition state. This is a hard requirement — keystroke injection with
an active Korean IME corrupts input (jamo recomposition) and races with the user's own typing.

Clipboard-paste fallback, only for spots with no object-model path (e.g. some HWP dialog-bound
controls or an embedded object): `win32clipboard` → save current clipboard → set
`CF_UNICODETEXT` → invoke the app's own paste (Word `rng.Paste()`; HWP `HAction.Run("Paste")`;
UNO `.uno:Paste` dispatch) → restore the user's clipboard in a `finally`. This is still not a
keystroke, and remains Hangul-safe. It is last-resort because it disturbs the user's clipboard
and is observable; each use is logged in the tool result ("used clipboard fallback").

Also set both `Font.Name` and `Font.NameFarEast` in Word when formatting (§3), or Korean glyphs
silently keep the previous East-Asian font.

---

## 8. Concurrency, errors, safety

- **Serialization**: one COM worker thread ⇒ all doc ops are serialized; the TS extension also
  queues calls per sidecar so `signal` (abort) can cancel *queued* jobs. A COM call already
  in-flight cannot be safely aborted mid-call — abort marks the job "orphaned"; its result is
  discarded on return.
- **Timeouts**: TS-side per-tool timeout (default 30 s; save/convert 120 s). The sidecar runs a
  watchdog thread: if the worker thread hasn't completed a job within its budget, the watchdog
  (a) responds with `TIMEOUT`, (b) probes for a modal dialog. It never kills the worker thread
  (that would poison the apartment); a persistently wedged worker ⇒ sidecar restart (handles
  invalidated, agent informed).
- **Modal dialogs**: the number-one wedge cause (file-recovery prompts, "keep in same format?",
  macro warnings, HWP security popup if the module registration broke). Detection: worker
  blocked or COM returning `RPC_E_CALL_REJECTED`/`RPC_E_SERVERCALL_RETRYLATER` + an enabled
  owned dialog window on the app's process (`EnumWindows` + `GetWindowThreadProcessId` +
  `WS_DLGFRAME`/`#32770` class). The sidecar emits a `modal_dialog` event with title/buttons.
  **Dismissal goes through the UIA layer**, not this sidecar — see the sibling doc
  `uia-port-design.md` (UIA port of a11y-harvest/a11y-action) for enumerating and pressing
  dialog buttons; policy: the agent is told about the dialog and must either handle it via the
  UIA tools or `ask_user`. Prevention: `DisplayAlerts` off in Word, `IgnoreMessage=1` in HWP
  psets, `Interaction=False` in UNO load/store `MediaDescriptor`.
- **Read-only documents**: detected at open (`doc.ReadOnly`, HWP `hwp.IsPrivateInfoProtected`/
  edit-mode check TODO(verify), UNO `doc.isReadonly()`); all mutating tools return `READ_ONLY`
  with the reason (file attribute, sharing lock, protected view). Word Protected View
  (`app.ProtectedViewWindows`) requires an explicit, approval-gated "enable editing".
- **Unsaved-changes protection & approval gates**: mirroring open-computer's `ask_user`/plan-
  review philosophy, the TS extension — not the sidecar — gates destructive operations. Gated:
  `doc_save` (overwrite), `doc_save_as` onto an existing path, `doc_close` with
  `discard_changes`, `doc_track_changes accept/reject` without a `revision_id` (accept-all), any
  `doc_replace` with `occurrence:"all"` whose match count exceeds a threshold (default 50 —
  reported first, then confirmed). Mechanism: the tool's `execute()` first performs a **dry-run
  sidecar call** (`{dry_run:true}` returns match counts / dirty state / target path exists),
  then calls the host's `ask_user` flow with a one-line consequence summary; only on approval
  does the real call run. A per-session "user pre-approved bulk edits on this doc" flag avoids
  nagging.
- **User-concurrent editing**: the `expect_hash` staleness machinery (§2.2) is the main defense.
  Additionally each mutating call re-checks the snapshot `rev` and, when it moved, the result
  warns "document changed since last read (rev r7→r9)". No locking is attempted — the user owns
  the document; the agent adapts.
- **Crash containment**: Word/HWP crashing takes the COM proxies with it (`com_error`
  `RPC_E_DISCONNECTED 0x80010108` / "The RPC server is unavailable"). Driver translates to
  `APP_NOT_RUNNING`, drops handles for that app only, and the agent may `doc_open` again (the
  apps' own file-recovery will show — which the dialog machinery then reports).

---

## 9. Code skeletons

### 8.1 Sidecar entry point + JSON contract (`docd.py`)

```python
#!/usr/bin/env python3
"""docd — document automation sidecar. Line-delimited JSON-RPC on stdio."""
import sys, json, threading, queue, traceback

class Sidecar:
    def __init__(self):
        self.jobs = queue.Queue()
        self.out_lock = threading.Lock()
        self.drivers = {}          # handle -> driver instance
        self.registry = {"word": WordDriver, "hwp": HwpDriver, "odt": OdtDriver,
                         "ppt": PptDriver, "impress": ImpressDriver}

    def send(self, obj):
        with self.out_lock:
            sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
            sys.stdout.flush()

    def reader(self):
        for line in sys.stdin:
            try:
                self.jobs.put(json.loads(line))
            except json.JSONDecodeError:
                self.send({"id": None, "error": {"code": "BAD_JSON", "message": line[:200]}})
        self.jobs.put(None)        # EOF -> shutdown

    def com_worker(self):
        import pythoncom
        pythoncom.CoInitialize()   # STA — all COM lives on this thread
        try:
            while (req := self.jobs.get()) is not None:
                rid = req.get("id")
                try:
                    result = self.dispatch(req["method"], req.get("params", {}))
                    self.send({"id": rid, "result": result})
                except DocError as e:            # typed errors -> stable codes
                    self.send({"id": rid, "error": {"code": e.code, "message": str(e), "data": e.data}})
                except Exception:
                    self.send({"id": rid, "error": {"code": "INTERNAL", "message": traceback.format_exc(limit=4)}})
        finally:
            for d in self.drivers.values(): d.release()
            pythoncom.CoUninitialize()

    def dispatch(self, method, params):
        if method == "doc_open":
            drv = self.registry[pick_backend(params)].attach_or_launch()
            handle = drv.open(params["path"], read_only=params.get("read_only", False))
            self.drivers[handle] = drv
            return drv.summary(handle)
        drv = self.drivers.get(params.get("doc")) or fail("NO_SUCH_DOC", params.get("doc"))
        return getattr(drv, method)(**params)     # doc_read / doc_replace / ...

if __name__ == "__main__":
    s = Sidecar()
    threading.Thread(target=s.reader, daemon=True).start()
    s.com_worker()                 # COM thread is the main thread
```

### 8.2 WordDriver — replace + insert

```python
import win32com.client, pythoncom, time

WD_FIND_STOP, WD_REPLACE_ONE, WD_REPLACE_ALL = 0, 1, 2
WD_COLLAPSE_END, WD_COLLAPSE_START = 0, 1
RETRYABLE = {-2147418111, -2147417846}   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER

def com_call(fn, *a, **kw):
    for attempt in range(5):
        try:
            return fn(*a, **kw)
        except pythoncom.com_error as e:
            if e.hresult in RETRYABLE and attempt < 4:
                time.sleep(0.3 * (attempt + 1)); continue
            raise map_com_error(e)     # -> DocError(APP_BUSY_MODAL / APP_NOT_RUNNING / COM_ERROR)

class WordDriver(BaseDriver):
    @classmethod
    def attach_or_launch(cls):
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except pythoncom.com_error:
            app = win32com.client.DispatchEx("Word.Application")
        app.Visible = True
        return cls(app)

    def doc_replace(self, doc, find, replace, regex=False, match_case=False,
                    occurrence="all", scope=None, dry_run=False, **_):
        d = self._doc(doc)
        rng = self._scope_range(d, scope)          # doc.Range() or para range
        if regex and not wildcard_safe(find):
            return self._regex_emulated_replace(d, rng, find, replace, occurrence)
        if occurrence == "all":
            if dry_run:
                return {"matches": self._count_matches(rng, find, match_case)}
            n = 0
            while com_call(rng.Find.Execute, FindText=find, MatchCase=match_case,
                           MatchWildcards=regex, Forward=True, Wrap=WD_FIND_STOP,
                           ReplaceWith=replace, Replace=WD_REPLACE_ONE):
                n += 1
                rng.Collapse(WD_COLLAPSE_END); rng.End = d.Content.End
            return {"replaced": n, "affected": self._touched_paras(d)}
        # first / Nth occurrence: walk with Replace=0, then replace the hit range
        target = self._nth_match(rng, find, match_case, regex,
                                 1 if occurrence == "first" else int(occurrence))
        com_call(setattr, target, "Text", replace)  # target.Text = replace
        self.app.ActiveWindow.ScrollIntoView(target)
        return {"replaced": 1, "affected": [self._para_ref(target)]}

    def doc_insert(self, doc, text, where, para=None, expect_hash=None,
                   bookmark=None, style_map=True, **_):
        d = self._doc(doc)
        if where == "bookmark":
            anchor = d.Bookmarks(bookmark).Range
        elif where == "end":
            anchor = d.Content; anchor.Collapse(WD_COLLAPSE_END)
        else:                                       # before_para / after_para
            anchor = self._resolve_para(d, para, expect_hash)   # raises STALE_RANGE
            anchor.Collapse(WD_COLLAPSE_START if where == "before_para" else WD_COLLAPSE_END)
        start = anchor.End
        com_call(anchor.InsertAfter, to_word_text(text))        # "\n" -> "\r"
        ins = d.Range(start, anchor.End)
        if style_map:
            apply_md_heading_styles(d, ins)         # '# ' prefixes -> Styles("Heading n")
        self.app.ActiveWindow.ScrollIntoView(ins)
        return {"inserted_paras": self._para_refs(ins)}
```

### 8.3 HwpDriver — insert_text + field write

```python
class HwpDriver(BaseDriver):
    @classmethod
    def attach_or_launch(cls):
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # Security module must already be registered under
        # HKCU\Software\HNC\HwpAutomation\Modules  (name "SecurityModule" -> DLL path).
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.XHwpWindows.Item(0).Visible = True
        return cls(hwp)

    def insert_text(self, text):
        """Insert at current cursor position. Hangul-safe: BSTR, no IME."""
        for i, line in enumerate(text.split("\n")):
            if i:
                self.hwp.HAction.Run("BreakPara")
            pset = self.hwp.HParameterSet.HInsertText
            self.hwp.HAction.GetDefault("InsertText", pset.HSet)
            pset.Text = line
            self.hwp.HAction.Execute("InsertText", pset.HSet)

    def doc_insert(self, doc, text, where, bookmark=None, **kw):
        if where == "bookmark":                     # HWP "bookmarks" == named fields (필드)
            if bookmark not in self._field_names():
                fail("NO_SUCH_ANCHOR", bookmark)
            self.hwp.PutFieldText(bookmark, text)   # most reliable HWP write path
            return {"wrote_field": bookmark}
        if where == "end":
            self.hwp.MovePos(3, 0, 0)               # moveBottomOfFile  TODO(verify) id table
        elif where in ("before_para", "after_para"):
            self._goto_para(kw["para"], kw.get("expect_hash"))   # scan-cache; may STALE_RANGE
        self.insert_text(text)
        return {"inserted": True, "note": "position-based; prefer field anchors on HWP"}

    def _field_names(self):
        raw = self.hwp.GetFieldList(0, 0x01)        # TODO(verify) option flags
        return [f for f in raw.split("\x02") if f]

    def doc_replace(self, doc, find, replace, occurrence="all", match_case=False, **_):
        act = "AllReplace" if occurrence == "all" else "ExecReplace"
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault(act, pset.HSet)
        pset.FindString, pset.ReplaceString = find, replace
        pset.IgnoreMessage = 1                       # no result popup
        pset.MatchCase = 1 if match_case else 0
        if occurrence not in ("all", "first"):       # position at Nth hit first
            self._repeat_find(find, int(occurrence) - 1)
        self.hwp.HAction.Execute(act, pset.HSet)
        return {"replaced": "all" if occurrence == "all" else 1}
```

### 8.4 OdtDriver — UNO connect + find/replace

```python
import uno
from com.sun.star.beans import PropertyValue

class OdtDriver(BaseDriver):
    @classmethod
    def attach_or_launch(cls, port=2002):
        ensure_soffice_listening(port)   # spawn: soffice --accept="socket,host=localhost,
                                         #   port=2002;urp;" --norestore --nologo   (visible!)
        localCtx = uno.getComponentContext()
        resolver = localCtx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localCtx)
        ctx = resolver.resolve(
            f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        return cls(ctx, desktop)

    def open(self, path, read_only=False):
        props = [mkprop("ReadOnly", read_only), mkprop("Interaction", False)]
        doc = self.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(path), "_blank", 0, tuple(props))
        return self._register(doc)

    def doc_replace(self, doc, find, replace, regex=False, match_case=False,
                    occurrence="all", **_):
        d = self._doc(doc)
        if occurrence == "all":
            rd = d.createReplaceDescriptor()
            rd.setSearchString(find); rd.setReplaceString(replace)
            rd.SearchCaseSensitive = match_case
            rd.SearchRegularExpression = regex       # real ICU regex
            return {"replaced": d.replaceAll(rd)}
        sd = d.createSearchDescriptor()
        sd.setSearchString(find); sd.SearchCaseSensitive = match_case
        sd.SearchRegularExpression = regex
        hit, n = d.findFirst(sd), 1
        want = 1 if occurrence == "first" else int(occurrence)
        while hit and n < want:
            hit, n = d.findNext(hit.getEnd(), sd), n + 1
        if not hit: fail("NOT_FOUND", find)
        hit.setString(replace)
        d.getCurrentController().getViewCursor().gotoRange(hit, False)
        return {"replaced": 1}

def mkprop(name, value):
    p = PropertyValue(); p.Name, p.Value = name, value; return p
```

---

## 10. Phased implementation plan

**Phase 1 — Word (2–3 wk).** Sidecar skeleton + JSON-RPC framing + TS extension with all 15
tool registrations (non-Word backends return `UNSUPPORTED_ON_BACKEND`). Full WordDriver: open/
read/outline/insert/replace/edit_range/apply_style/format/save/save_as; then comments, track
changes, tables, screenshot. Approval gate + dialog watchdog land here (they're
backend-agnostic).
*Tests*: pytest suite driving a real Word instance on a Windows CI runner (self-hosted or
GitHub Actions `windows-latest` + Office image) against fixture docs — golden-file assertions by
re-extracting text; staleness tests that mutate the doc between read and edit; dialog test that
force-opens a modal (`app.Dialogs(wdDialogFileOpen).Show` on a helper thread — TODO(verify)
practicality) and asserts `APP_BUSY_MODAL`. Korean fixtures throughout (Hangul + mixed-script).

**Phase 2 — PowerPoint (1–2 wk).** PptDriver + the `slide_*` tool subset. PowerPoint's object
model is as mature as Word's and reuses Phase 1's COM plumbing wholesale (retry loop, dialog
watchdog, approval gates), so this phase is mostly driver code: open/list/read → edit_text/
notes → add/reorder/delete/duplicate → thumbnails (`Slide.Export`) → save_as/pdf/png. Shape-id
addressing lands here and is validated against reorder-heavy decks.
*Tests*: pytest against real PowerPoint on the Windows runner; thumbnail golden-image diffs
(perceptual hash, not byte-equal); Korean deck fixtures; a python-pptx closed-mode smoke suite
that runs without Office.

**Phase 3 — HWP (2–3 wk).** Security-module registration in installer + self-heal; HwpDriver on
pyhwpx with raw-HAction escape hatch. Order: open/save/save_as → PutFieldText/GetFieldText +
field enumeration → insert_text/replace → scan-based doc_read + hash cache → outline
(best-effort) → tables (read + field-cell write). Explicit test pass for every TODO(verify) in
§4 against the target Hancom version(s); document actual behavior in a
`hwp-compat.md` matrix. Hancom Show is scoped **out** here except the pptx-interchange routing
message and a spike to test whether a scriptable Show automation object exists (§6.4).
*Tests*: same pytest pattern; requires a licensed Hancom Office install on the runner (see
risks). Round-trip test: agent fills a field template → PDF export → text-layer assertion.

**Phase 4 — ODT + ODP (1–2 wk).** soffice lifecycle management (socket flag, lock-file
detection, restart-with-consent flow), UNO OdtDriver for the full text surface, ImpressDriver
for the `slide_*` surface on odp, odfpy/python-pptx closed-file modes.
*Tests*: easiest to CI (LibreOffice is free & headless-capable — run the same suite twice,
visible and `--headless`, since CI has no display).

**Cross-cutting final pass**: conversion-matrix integration tests (docx→pdf, hwp field-fill→pdf,
odt→docx, pptx→pdf/png, pptx round-trip through Impress), 3-hour soak test of the sidecar
(handle leaks, Word/PowerPoint memory), agent-level eval tasks ("이 보고서의 2장을 요약해
덧붙여 줘", "이 발표자료 5번 슬라이드 뒤에 요약 슬라이드를 추가해 줘" against all backends).

---

## 11. Risks and open questions

1. **Hancom automation licensing/redistribution.** The HWP automation control ships with Hancom
   Office (user must own a license — we do not redistribute it). The security-module DLL
   (`FilePathCheckerModuleExample.dll`) comes from Hancom's automation SDK; redistribution terms
   must be checked with Hancom legal — worst case the user installs the SDK, or we implement our
   own module against the documented interface. Open question: minimum supported Hancom version
   (2018 vs 2020+) — action names and `SaveAs` format strings drift.
2. **HWP API opacity.** Official docs are a Korean-language CHM; several details here carry
   TODO(verify) markers (moveID table, `InitScan` flags, `CreateField` arg order, HWPX format
   string, regex member). Mitigation: pyhwpx source as living documentation + the Phase 2
   verification pass. Position-based editing remains best-effort; the product answer is
   field-anchored templates.
3. **Hancom Show automation gap.** The design assumes Show has no usable automation OM and
   routes presentations through pptx interchange (§6.4). If Korean-market users predominantly
   keep `.show` files, the only in-place path is UIA-driven UI automation (slow, fragile) —
   this is a product-level exposure to validate early with real user data; the Phase 3 spike
   answers whether any scriptable surface exists.
4. **Word/PowerPoint interop variance.** Perpetual Office 2016/2019 vs Microsoft 365 differ in the comments
   OM (threaded replies, `Done`), and Protected View/AutoSave (OneDrive) changes save semantics
   (`doc.Save()` on an AutoSaved cloud doc is near-instant but versioned). Office delivered via
   the Microsoft Store historically had broken COM registration — detect and message it. Word/
   PowerPoint must be *installed*; if absent, route docx to LibreOffice Writer and pptx to
   Impress or python-pptx, with a fidelity warning.
5. **LibreOffice bundling size.** Full LO is ~350 MB installed; bundling balloons our installer.
   Options: (a) detect an existing install, (b) optional download-on-demand component, (c) LO
   "portable"/stripped Writer-only build (~200 MB). Also: pyuno version-locking to the bundled
   LO vs system LO needs a decision in Phase 3.
6. **Single-instance apps & user contention.** Word/HWP automation shares the instance with the
   user's other documents; a modal in *their* document blocks *our* COM calls. The dialog
   watchdog covers detection, but UX policy (when may the agent auto-dismiss vs must ask) needs
   product sign-off — see `uia-port-design.md`.
7. **Regex parity.** Native regex only on UNO; Word wildcards are a subset; HWP unclear. The
   emulated read-modify-write path breaks character formatting inside rewritten paragraphs.
   Tool description must warn: "regex replacements may reset character formatting in affected
   paragraphs on Word/HWP".
8. **hwpx roadmap.** Hancom's newer HWPX (OWPML, zip/XML) can be manipulated file-level like
   odfpy does ODF — a future closed-file HwpxDriver could reduce dependence on the COM control
   for batch work. Not scheduled; noted as an escape hatch if COM flakiness dominates.
9. **Accessibility of the sidecar itself.** Long-running Python on user machines: needs
   watchdog restart, log rotation (mirror `writeTrace`-style JSONL in the app's log dir), and a
   kill switch when the user closes the assistant.
