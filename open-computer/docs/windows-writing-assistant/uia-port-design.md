# Windows UIA Port of the AT-SPI Control Layer (`uia-harvest` / `uia-action`)

Design for porting open-computer's Linux accessibility control layer
(`master/setup/a11y-harvest.py`, `master/setup/a11y-action.py`, exposed as pi tools in
`services/extensions/desktop-apps.ts`) to Windows UI Automation, for the Windows
writing-assistant desktop app. The pattern stays identical: **harvest structured UI
state → act by element label/ID, no screenshots in the hot path.**

Sibling doc: `com-toolset-design.md` (COM automation of Word / Hancom HWP document
content). This doc covers the *UI chrome* layer; document *content* editing in HWP and
Word should go through COM whenever possible (see §7).

---

## 1. Technology choice

### Candidates

| | Python `uiautomation` | `pywinauto` (backend="uia") | C# FlaUI (UIA3) / System.Windows.Automation |
|---|---|---|---|
| UIA backend | UIA COM (`UIAutomationClient`) via comtypes | UIA COM via comtypes, heavy wrapper layer | FlaUI: raw `UIAutomationClient` COM interop (UIA3). SWA: managed UIA2, older, slower, missing newer patterns |
| CacheRequest support | Not exposed (walks are property-by-property, one COM round trip each) | Not exposed in the public API | Fully exposed (`CacheRequest`, `TreeScope`, cached properties/patterns) — the single biggest lever for large trees |
| Full-tree walk on Word-sized apps | Slow (seconds); every property read is a cross-process COM call | Slowest of the three; `ElementInfo` wrapping adds large constant factors; known to take 10s+ on Office ribbons | Fast: one `FindAll` with a condition + cache = one cross-process bulk fetch. Sub-second for scoped walks |
| COM threading | Initializes COM per thread; MTA possible but fiddly | STA by default; MTA needs care | You choose: run the client in an **MTA** thread (recommended by MS for UIA clients; STA client + STA provider in-proc can deadlock). Trivial in C#: no `[STAThread]` on `Main` |
| Distribution inside a desktop app | Ship an embeddable CPython (~15–50 MB) + comtypes; pip supply chain; AV false positives on bundled python.exe are common | Same, plus a much larger dependency tree | .NET 8 self-contained single-file exe (~15 MB trimmed, or ~1 MB if you target the .NET runtime you already ship with the app). Signs cleanly, no interpreter |
| Ecosystem fit | Fine for prototyping | Mature but aimed at test automation, not sidecars | FlaUI is MIT, active, thin; SWA (managed UIA2) is legacy — avoid |

### Recommendation

**Primary stack: C# on .NET 8 with FlaUI (`FlaUI.Core` + `FlaUI.UIA3`), built as a
self-contained, single-file console exe (`uia-tool.exe`).**

Rationale, in priority order:

1. **Performance on Word/Hancom.** Only FlaUI/C# gives us `CacheRequest` + condition-based
   `FindAll`, which turns O(elements × properties) cross-process COM calls into a handful
   of bulk fetches. This is the difference between 200 ms and 10 s on the Word ribbon.
2. **Packaging.** A signed single-file exe drops into the desktop app's install dir. No
   Python runtime, no comtypes cache directory, no AV heuristics on an embedded interpreter.
3. **Threading control.** UIA clients must not run on the app's UI STA thread. A separate
   exe (or an MTA worker thread if we later in-proc it) sidesteps the whole class of
   STA re-entrancy deadlocks.
4. **Escape hatches.** C# gives clean P/Invoke for `SendInput`, `WM_MOUSEWHEEL`,
   clipboard, and `IAccessible` (MSAA) when we need the Hancom fallbacks.

Do **not** use `System.Windows.Automation` (managed UIA2): it lags the COM API (no
`TextPattern2`, weaker caching, worse perf) and is effectively frozen.

Python `uiautomation` remains acceptable for *prototyping* the role mapping and for quick
experiments, since its API names map 1:1 onto UIA concepts — but it is not the shipping
sidecar because of the CacheRequest gap and packaging cost.

---

## 2. The harvest port: `uia-tool.exe harvest`

Mirrors `a11y-harvest.py` exactly at the contract level.

### 2.1 CLI contract (identical shape to Linux)

```
uia-tool.exe harvest                       → {"apps": ["Word", "한글", ...]}        # top-level windows
uia-tool.exe harvest <app>                 → {"app", "window", "actions", "texts"}
uia-tool.exe harvest <app> --pretty
uia-tool.exe harvest <app> --scope ribbon  → optional: harvest a normally-skipped subtree
```

"App" matching: normalized substring match (lowercase, strip spaces/hyphens/underscores —
same `_normalize` as Linux) against **top-level window titles** *and* **process names**
(`Element.Properties.ProcessId` → `Process.GetProcessById().ProcessName`). Windows has no
AT-SPI "application" node; the desktop's children are top-level windows, so `harvest`
with no args lists `"<process>: <window title>"` pairs and matching accepts either part.

### 2.2 Role mapping: AT-SPI role → UIA ControlType

The output JSON keeps the **Linux role vocabulary** (`"push button"`, `"check box"`,
`"text"`, …) so `formatHarvest`, prompts, and any agent few-shots port unchanged. The
mapping below is applied at emit time (UIA → AT-SPI-style string).

**INTERACTIVE set** (harvested into `actions`):

| UIA ControlType (+ discriminator) | emitted role |
|---|---|
| `Button` (no TogglePattern) | `push button` |
| `Button` with TogglePattern | `toggle button` |
| `CheckBox` | `check box` |
| `RadioButton` | `radio button` |
| `ComboBox` | `combo box` |
| `MenuItem` (plain) | `menu item` |
| `MenuItem` with TogglePattern | `check menu item` |
| `Edit` (`IsPassword=false`) | `text` |
| `Edit` (`IsPassword=true`) | `password text` |
| `Spinner` | `spin button` |
| `Slider` | `slider` |
| `ScrollBar` | `scroll bar` |
| `Hyperlink` | `link` |
| `TabItem` | `page tab` |
| `ListItem` | `list item` |
| `TreeItem` | `tree item` |
| `SplitButton` | `push button` |
| `Custom`/`Pane` with LegacyIAccessible role PUSHBUTTON etc. | mapped from MSAA role (Hancom path, §7) |

**TEXT set** (harvested into `texts`):

| UIA ControlType | emitted role |
|---|---|
| `Text` | `label` |
| `Header` / `HeaderItem` | `heading` |
| `StatusBar` (its `Text`/`Edit` children's names concatenated) | `status bar` |
| `ToolTip` | `description` |
| `Document` | *not walked* — surfaced via TextPattern, see §2.6 |

**NOISE set** (never emitted, but **descended through**): `Pane`, `Group`, `Window`,
`TitleBar`, `Separator`, `Image`, `Thumb`, `ToolBar`, `MenuBar`, `Menu`, `Tree`, `Table`,
`List`, `DataGrid`, `DataItem` (containers only), `AppBar`, `SemanticZoom`, `Custom`
(unless MSAA fallback promotes it).

Like the Linux script, classification is a **post-filter**: containers are traversed but
not emitted; only leaf-classified nodes produce entries.

### 2.3 Label resolution (ordered)

1. `Name` property (cached). UIA `Name` is far more reliably populated than AT-SPI names —
   Office sets it on essentially everything.
2. `LabeledBy` property → that element's cached `Name`. (Direct analogue of AT-SPI
   `LABELLED_BY`.)
3. For `Edit`/`ComboBox` with empty name: nearest **preceding sibling** of ControlType
   `Text` within 2 siblings (same heuristic as Linux `find_label_for`), then parent `Name`.
4. `HelpText` property (often holds the tooltip; ribbon buttons carry rich names already,
   but Win32 dialogs sometimes only have HelpText).
5. `AutomationId` as a last resort, rendered as `#<AutomationId>` (e.g. `#btnSubmit`) —
   stable, developer-assigned, better than `(unlabeled)`.
6. `(unlabeled)`.

Truncate to 80 chars, exactly like Linux.

### 2.4 State flags mapping

| Linux JSON flag | UIA source |
|---|---|
| `checked: true` | `TogglePattern.ToggleState == On` (also `SelectionItemPattern.IsSelected` for radio buttons / tab items — emit as `checked` for radios to match AT-SPI behavior) |
| `editable: true` | ValuePattern available **and** `Value.IsReadOnly == false`; or ControlType `Edit`/`Document` with `IsEnabled` |
| `focused: true` | `HasKeyboardFocus` |
| `expanded: true` | `ExpandCollapsePattern.ExpandCollapseState == Expanded` |
| `disabled: true` | `IsEnabled == false` (AT-SPI `sensitive` absent ⇒ disabled — same semantics) |
| `value` | `Value.Value` (string) for text-bearing controls; `RangeValue.Value` (number) for sliders/spinners — matches Linux where `Value` interface overwrites text value |
| visibility filter | `IsOffscreen == true` ⇒ skip entirely (analogue of requiring extents > 0). Combined with zero-area `BoundingRectangle` check |

### 2.5 Coordinates, dedup, caps

- `x`,`y` = center of `BoundingRectangle` (physical screen px — see DPI note in §8);
  `w`,`h` kept internally for scroll targeting.
- Dedup key `(role, label, x, y)` — identical to Linux.
- Caps identical: `actions[:150]`, `texts[:100]`. Priority when clipping: focused element
  first, then document-order. (The Linux version clips in document order; keep that, but
  never clip out the focused element — cheap insurance for huge apps.)

### 2.6 Output JSON schema (byte-compatible with Linux)

```json
{
  "app": "Word — report.docx",
  "window": { "title": "report.docx - Word", "x": 0, "y": 0, "w": 1920, "h": 1040 },
  "actions": [
    { "id": 1, "role": "push button", "label": "Bold", "x": 212, "y": 118,
      "checked": true, "focused": false }
  ],
  "texts": [
    { "id": 1, "role": "label", "text": "Page 1 of 3" }
  ]
}
```

Errors: `{"error": "...", "available": [...]}` — same as Linux. `formatHarvest` in
desktop-apps.ts works **unmodified**.

One Windows-specific addition that does not break the schema: each action entry also gets
an internal RuntimeId recorded in the **element cache** (§4), keyed by `id`. The JSON the
model sees is unchanged.

### 2.7 Performance strategy for enormous trees (Word, Hancom)

Naive full recursion (the Linux approach) is unusable on Word: the ribbon alone is
thousands of elements and every property read is a cross-process COM call. Design:

1. **Scope to one top-level window.** Never walk from the desktop. Resolve the target
   window first (`FindAllChildren` on the desktop with a `Window` condition, one call),
   then all work happens under it.
2. **Condition-based `FindAll` instead of recursion.** One call:
   `window.FindAll(TreeScope.Descendants, orConditionOfInterestingControlTypes)` where the
   condition is `(ControlType==Button) OR (ControlType==Edit) OR ... OR (ControlType==Text)`
   AND `IsOffscreen==false`. UIA evaluates the condition provider-side; we never
   materialize noise nodes at all. This replaces `collect_all` + `classify` in one shot.
3. **CacheRequest.** Activate a `CacheRequest` around the `FindAll` that prefetches:
   `Name, ControlType, BoundingRectangle, IsEnabled, HasKeyboardFocus, IsOffscreen,
   AutomationId, HelpText, RuntimeId, IsKeyboardFocusable, LabeledBy` plus pattern
   availability (`IsTogglePatternAvailable`, `IsValuePatternAvailable`,
   `IsExpandCollapsePatternAvailable`, `IsRangeValuePatternAvailable`) and the
   `Toggle`/`ExpandCollapse`/`Value`/`RangeValue`/`SelectionItem` patterns themselves.
   Result: the entire harvest is ~2–4 cross-process round trips.
4. **Skip the document content subtree.** Before the big `FindAll`, locate
   `ControlType.Document` children (Word: class `_WwG`; the Document control) with one
   scoped find. Exclude their descendants by running the main `FindAll` with
   `TreeScope.Descendants` from each *sibling region* — or simpler and robust: run the
   full `FindAll`, then drop any hit whose cached `RuntimeId` chain falls inside a
   document's `BoundingRectangle` **and** whose ControlType is `Text` (ribbon/statusbar
   text lives outside the document rect). The document body instead contributes to
   `texts` via **TextPattern**: `doc.Patterns.Text.Pattern.DocumentRange.GetText(4000)`,
   split into ≤300-char chunks emitted as `{"role": "paragraph", ...}` entries. This is
   both faster and higher-fidelity than tree text nodes.
5. **Depth limit for the fallback walker.** When an app's provider mishandles `FindAll`
   with complex conditions (rare; some MSAA-bridged apps), fall back to a manual BFS
   using `TreeWalker.ControlViewWalker` with cached children
   (`CacheRequest.TreeScope = Element | Children`), max depth **12**, max nodes **5000**,
   wall-clock budget **3 s** — emit what we have with `"truncated": true` beyond that.
6. **Ribbon policy (Word).** The ribbon (class `NetUIHWND`) is harvested, but only the
   *active tab's* controls are on-screen; `IsOffscreen==false` in the condition already
   prunes inactive tabs. Additionally cap ribbon-descendant actions at 60 so a maximized
   Home tab can't starve the rest of the window. `--scope ribbon` lifts the cap when the
   agent explicitly wants to enumerate ribbon commands.

Expected budget: Word main window harvest ≤ 400 ms warm; Hancom (MSAA-bridged, slower
provider) ≤ 1.5 s.

---

## 3. The action port: `uia-tool.exe action`

CLI mirrors `a11y-action.py`:

```
uia-tool.exe action <app> click     <label|#id>
uia-tool.exe action <app> set_text  <label|#id> <value>
uia-tool.exe action <app> get_text  <label|#id>
uia-tool.exe action <app> select    <combo_label|#id> <item>
uia-tool.exe action <app> focus     <label|#id>
uia-tool.exe action <app> key       <keyname>            # xdotool-style: ctrl+s, Return, alt+F4
uia-tool.exe action <app> type      <text>
uia-tool.exe action <app> scroll    <direction> [amount]
uia-tool.exe action <app> set_value <label|#id> <number>
uia-tool.exe action <app> action    <label|#id> <name>   # named LegacyIAccessible/pattern action
```

Output: `{"ok": true, ...}` / `{"error": "..."}`, same fields as Linux
(`label`, `role`, `action`, `value_set`, `readback`, …).

### 3.1 Per-command fallback chains

**click** — resolve element (§4), then:
1. `InvokePattern.Invoke()` — buttons, menu items, links.
2. `TogglePattern.Toggle()` — checkboxes, toggle buttons (when Invoke unavailable).
3. `SelectionItemPattern.Select()` — radio buttons, list items, tab items.
4. `ExpandCollapsePattern.Expand()` — parent menu items / split-button dropdowns.
5. `LegacyIAccessiblePattern.DoDefaultAction()` — MSAA bridge; **primary path for
   Hancom** (§7).
6. Hard fallback: `SetForegroundWindow(topLevelHwnd)` + `SendInput` mouse move/click at
   `BoundingRectangle` center. Report `"method": "mouse"` in the JSON so the agent knows
   the semantic path failed. Re-read `BoundingRectangle` *live* (not cached) immediately
   before clicking — layout may have shifted since harvest.

**set_text** —
1. `ValuePattern.SetValue(value)`; readback via `Value.Value` or `TextPattern` and include
   in JSON (matches Linux `readback`).
2. Fallback: `element.Focus()` → `Ctrl+A` → clipboard paste (§3.2) — for rich edits that
   expose TextPattern but a read-only/absent ValuePattern (Word document body, many
   Electron apps).
3. Error if neither pattern nor focus is achievable.

**get_text** — `Value.Value` → `TextPattern.DocumentRange.GetText(-1)` → cached `Name`.

**select** —
1. `combo.Patterns.ExpandCollapse.Pattern.Expand()`; wait ≤500 ms for the popup;
   find descendant `ListItem` whose name matches item (exact then substring — popup may
   parent to the desktop for Win32 combos, so search both the combo subtree and a
   desktop-level `Window/Pane` that appeared after Expand); `SelectionItemPattern.Select()`;
   `Collapse()`.
2. Fallback: `combo.Patterns.Value.Pattern.SetValue(item)` — many Win32/WinForms combos
   accept this directly.
3. Fallback: focus + type first letters + `Enter`.

**focus** — `element.Focus()` (UIA `SetFocus`). If it throws (common on ribbon items),
fall back to a mouse click at center.

**set_value** — `RangeValuePattern.SetValue(n)`; clamp to `Minimum/Maximum` and report the
clamped readback. Fallback for sliders without RangeValue: LegacyIAccessible `accValue`.

**key** — parse the **xdotool syntax the Linux tools already use** (`ctrl+shift+s`,
`Return`, `Escape`, `alt+F4`) into virtual-key chords; `SetForegroundWindow` on the app's
top-level window, then `SendInput` with `KEYEVENTF_SCANCODE` down/up pairs
(scan codes via `MapVirtualKey`, so apps reading hardware scan codes behave). Modifier
name map: `ctrl→VK_CONTROL, alt→VK_MENU, shift→VK_SHIFT, super→VK_LWIN`; key name map for
the xdotool names the tool descriptions advertise (`Return, Tab, Escape, BackSpace,
Delete, Up/Down/Left/Right, F1–F12, Home, End, Prior→PgUp, Next→PgDn`).

**type** — see §3.2.

**scroll** —
1. Find the best scroll target: prefer the element with `ScrollPattern` whose
   `Vertically/HorizontallyScrollable` is true and largest on-screen area (analogue of the
   Linux "largest scroll pane" heuristic); call
   `ScrollPattern.Scroll(ScrollAmount.SmallDecrement/SmallIncrement …)` `amount` times
   (or `LargeIncrement` for amount ≥ 5).
2. Fallback: `SendMessage(hwnd, WM_MOUSEWHEEL, MAKEWPARAM(0, ±120*amount), MAKELPARAM(x,y))`
   to the HWND under the target center (`WindowFromPoint`), no cursor movement needed.
3. Last resort: `SendInput` `MOUSEEVENTF_WHEEL` after moving the cursor to center.

### 3.2 Text entry and the Korean IME caveat

The writing assistant targets Korean users; text will routinely contain Hangul.

**Do not port `xdotool type` as per-character key events.** Synthetic per-char key
injection interacts with the active IME: with the Korean IME on, injected jamo-level
keystrokes enter Hangul *composition* state, and timing-sensitive composition commits
produce dropped or reordered syllables. Even `KEYEVENTF_UNICODE` (which injects
`WM_CHAR`-style input, bypassing layout) is unreliable in apps that run their own
composition handling (Word, HWP) and for surrogate pairs.

Text entry priority for both `set_text` fallback and `type`:

1. **`ValuePattern.SetValue`** — no keyboard, no IME involvement. Always first choice.
2. **Clipboard paste**: save clipboard (`OleGetClipboard`/`GetClipboardData` for
   `CF_UNICODETEXT` + enumerate formats we can round-trip), set `CF_UNICODETEXT` to the
   payload, send `Ctrl+V` via scan-code `SendInput`, wait 150 ms, **restore the previous
   clipboard**. Deterministic for arbitrary Unicode including Hangul, emoji, newlines.
   Report `"method": "paste"`.
3. **`KEYEVENTF_UNICODE` SendInput** only for short ASCII strings (≤ 32 chars, no CJK) —
   e.g. typing into a filter box that intercepts paste.
4. Never scan-code-type text content. Scan codes are reserved for `key` chords only.

The `app_type` tool description on Windows should be updated to say: "uses clipboard
paste under the hood; safe for Korean text."

---

## 4. Element identification and staleness

Linux resolves purely by live label match at action time — acceptable there because
AT-SPI walks are cheap and apps are small. On Windows, re-walking Word per action is
expensive, and label collisions are common (three "Bold" buttons: ribbon, mini toolbar,
context menu). Design: **hybrid — ids + RuntimeId cache, label fallback.**

1. **Harvest assigns ids** (`[1]`, `[2]`, … as today) and writes a sidecar cache:
   `%LOCALAPPDATA%\open-computer\uia-cache\<norm-app>.json`:

   ```json
   { "ts": 1720000000, "hwnd": 132848,
     "elements": { "1": { "runtimeId": [42, 132848, 4, 17], "automationId": "Bold",
                            "label": "Bold", "role": "toggle button", "x": 212, "y": 118 } } }
   ```

2. **Action resolution order** for a `<label|#id>` argument:
   a. If the argument is `#<n>` (or a bare int matching a cached id) and the cache is
      fresher than **TTL = 45 s** and the top-level `hwnd` still exists:
      re-find by **RuntimeId** — one `FindFirst(TreeScope.Descendants,
      PropertyCondition(RuntimeIdProperty, cachedRuntimeId))` scoped to the window.
      (RuntimeId property conditions are supported by UIA; if a given provider rejects
      it, fall back to a cached-`AutomationId` + ControlType condition.)
   b. If (a) misses (element rebuilt, RuntimeIds are not stable across re-creation):
      re-find by cached `AutomationId` (+ControlType), then by cached exact label.
   c. Label path (compatible with Linux behavior and with agents that pass labels):
      scoped `FindAll` with role-filter condition, exact-match first, then substring —
      exactly the Linux `find_element` semantics. Ambiguity (>1 exact match): prefer the
      one nearest the cached coordinates if a cache entry exists, else the first in
      document order, and include `"matches": n` in the JSON so the agent can disambiguate.
3. **Verify before acting**: after resolution, check live `IsOffscreen == false` and
   `IsEnabled` for the requested verb; return
   `{"error": "Element 'Bold' is stale/offscreen — re-run app_read_state"}` rather than
   clicking a ghost.
4. The tool layer keeps passing labels (no desktop-apps.ts changes needed), but the
   formatted harvest already shows `[id]`, so prompt guidance can nudge the model to use
   `#id` for precision. Both address forms hit the same resolver.

Future option (not v1): a persistent sidecar daemon holding live `IUIAutomationElement`
references and speaking newline-delimited JSON over stdio, eliminating cache serialization
entirely. The CLI contract below is designed so the daemon can slot in behind the same
verbs.

---

## 5. Process / tool surface

### 5.1 Spawn contract

Everything in desktop-apps.ts ports with **only the spawn command and env changed**:

```ts
const UIA_TOOL = join(app.getAppPath(), "resources", "uia-tool.exe");

const UIA_OPTS = {
  encoding: "utf8" as const,
  timeout: 15000,
  maxBuffer: 1024 * 1024,
  windowsHide: true,           // no flashing console window
  env: process.env,            // no DISPLAY/DBUS/GTK_MODULES needed on Windows
};

// Linux:  execWithStderr("python3", [A11Y_HARVEST, app], PYTHON_OPTS)
// Windows: execWithStderr(UIA_TOOL, ["harvest", app], UIA_OPTS)
// Linux:  execWithStderr("python3", [A11Y_ACTION, app, "click", label], PYTHON_OPTS)
// Windows: execWithStderr(UIA_TOOL, ["action", app, "click", label], UIA_OPTS)
```

stdout is a single JSON document, UTF-8 (`Console.OutputEncoding = Encoding.UTF8` — do
not let the default OEM codepage mangle Hangul). Exit code 0 for `{"ok":...}`/harvests,
1 for `{"error":...}` — same as the Python scripts.

### 5.2 Tool set (TypeBox schemas)

Identical names and shapes to desktop-apps.ts; deltas noted.

| Tool | Parameters | Windows implementation notes |
|---|---|---|
| `app_list` | `{ query?: string }` | Installed apps from Start Menu `.lnk`/registry `App Paths` + UWP via `Get-StartApps`; "currently running" from `harvest` (top-level windows) |
| `app_open` | `{ exec: string, app_name?: string }` | `spawn` the exe / `explorer.exe shell:AppsFolder\<AUMID>` for UWP; poll `harvest` for the new window (same before/after diff loop) |
| `app_read_state` | `{ app: string, scope?: "default"\|"ribbon"\|"menus" }` | `harvest <app>`; `scope` is the only new (optional) param |
| `app_click` | `{ app, label }` | `action <app> click <label>` (label may be `#id`) |
| `app_fill` | `{ app, label, value }` | `action <app> set_text ...` |
| `app_select` | `{ app, combo_label, item }` | `action <app> select ...` |
| `app_key` | `{ app, key }` | xdotool-style names preserved (`ctrl+s`, `Return`) |
| `app_type` | `{ app, text }` | clipboard-paste path (§3.2); description updated for Korean safety |
| `app_scroll` | `{ app, direction: "up"\|"down"\|"left"\|"right", amount?: number }` | ScrollPattern → WM_MOUSEWHEEL |
| `app_set_value` | `{ app, label, value: number }` | RangeValuePattern |
| `app_do_action` | `{ app, label, action }` | named LegacyIAccessible default action / pattern verb (`expand`, `collapse`, `toggle`, `invoke`, `select`, `dodefaultaction`); unknown name returns the available verbs, same as Linux |
| `app_screenshot` | `{ window?: string, save_path?: string }` | `uia-tool.exe screenshot [title]` using `PrintWindow`/`Graphics.CopyFromScreen`; base64 PNG on stdout JSON |

Example TypeBox (unchanged from Linux except description text):

```ts
parameters: Type.Object({
  app:   Type.String({ description: "App name or window-title substring (e.g. 'word', '한글')" }),
  label: Type.String({ description: "Element label or #id from app_read_state" }),
})
```

---

## 6. Code skeletons (C# / FlaUI, UIA3)

### 6.1 Harvest walker with CacheRequest

```csharp
using FlaUI.Core; using FlaUI.Core.AutomationElements; using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions; using FlaUI.UIA3;

static object HarvestApp(string appQuery)
{
    using var automation = new UIA3Automation();
    var desktop = automation.GetDesktop();

    // 1. Resolve target top-level window (one FindAll, names only).
    var win = FindTopLevelWindow(desktop, appQuery);          // normalized substring match
    if (win == null) return new { error = $"No app matching '{appQuery}'", available = ListWindows(desktop) };

    // 2. Build the cache request: everything we need, prefetched in bulk.
    var cr = new CacheRequest {
        TreeScope = TreeScope.Element,
        AutomationElementMode = AutomationElementMode.Full,   // keep refs for RuntimeId ops
    };
    var pl = automation.PropertyLibrary;
    foreach (var p in new[] { pl.Element.Name, pl.Element.ControlType, pl.Element.BoundingRectangle,
                              pl.Element.IsEnabled, pl.Element.HasKeyboardFocus, pl.Element.IsOffscreen,
                              pl.Element.AutomationId, pl.Element.HelpText, pl.Element.RuntimeId,
                              pl.Element.LabeledBy })
        cr.Add(p);
    cr.Add(automation.PatternLibrary.TogglePattern);
    cr.Add(automation.PatternLibrary.ValuePattern);
    cr.Add(automation.PatternLibrary.ExpandCollapsePattern);
    cr.Add(automation.PatternLibrary.RangeValuePattern);
    cr.Add(automation.PatternLibrary.SelectionItemPattern);

    // 3. One provider-side filtered FindAll for interactive + text control types.
    var cf = automation.ConditionFactory;
    var interesting = new OrCondition(
        cf.ByControlType(ControlType.Button),   cf.ByControlType(ControlType.CheckBox),
        cf.ByControlType(ControlType.RadioButton), cf.ByControlType(ControlType.ComboBox),
        cf.ByControlType(ControlType.MenuItem), cf.ByControlType(ControlType.Edit),
        cf.ByControlType(ControlType.Spinner),  cf.ByControlType(ControlType.Slider),
        cf.ByControlType(ControlType.Hyperlink), cf.ByControlType(ControlType.TabItem),
        cf.ByControlType(ControlType.ListItem), cf.ByControlType(ControlType.TreeItem),
        cf.ByControlType(ControlType.SplitButton), cf.ByControlType(ControlType.Text));
    var cond = new AndCondition(interesting,
        new PropertyCondition(pl.Element.IsOffscreen, false));

    AutomationElement[] found;
    Rectangle[] docRects;
    string docText = "";
    using (cr.Activate())
    {
        // 3a. Document controls first — read via TextPattern, exclude from tree text.
        var docs = win.FindAll(TreeScope.Descendants, cf.ByControlType(ControlType.Document));
        docRects = docs.Select(d => d.Cached.BoundingRectangle).ToArray();
        var tp = docs.FirstOrDefault()?.Patterns.Text.PatternOrDefault;
        if (tp != null) docText = tp.DocumentRange.GetText(4000);

        found = win.FindAll(TreeScope.Descendants, cond);     // the bulk fetch
    }

    var actions = new List<object>(); var texts = new List<object>();
    var seen = new HashSet<(string, string, int, int)>();
    var cache = new Dictionary<int, object>();                // id → RuntimeId sidecar
    int aid = 1, tid = 1;

    foreach (var el in found)
    {
        var c = el.Cached;                                    // no COM round trips below
        var rect = c.BoundingRectangle;
        if (rect.Width <= 0 || rect.Height <= 0) continue;
        int cx = rect.X + rect.Width / 2, cy = rect.Y + rect.Height / 2;
        string role = MapRole(c.ControlType, el);             // §2.2 table (Toggle → toggle button …)

        if (role == "label")                                  // TEXT set
        {
            if (InsideAny(docRects, rect)) continue;          // doc body comes from TextPattern
            var t = (c.Name ?? "").Trim();
            if (t.Length > 1 && texts.Count < 100)
                texts.Add(new { id = tid++, role, text = Truncate(t, 300) });
            continue;
        }

        string label = ResolveLabel(el, c);                   // Name→LabeledBy→sibling→HelpText→#AutomationId
        if (!seen.Add((role, label, cx, cy))) continue;

        var entry = new Dictionary<string, object>
            { ["id"] = aid, ["role"] = role, ["label"] = label, ["x"] = cx, ["y"] = cy };
        var tog = el.Patterns.Toggle.PatternOrDefault;
        if (tog?.ToggleState.ValueOrDefault == ToggleState.On) entry["checked"] = true;
        var val = el.Patterns.Value.PatternOrDefault;
        if (val != null) { if (!val.IsReadOnly.ValueOrDefault) entry["editable"] = true;
                           var v = val.Value.ValueOrDefault; if (!string.IsNullOrEmpty(v)) entry["value"] = Truncate(v, 80); }
        var rv = el.Patterns.RangeValue.PatternOrDefault;
        if (rv != null) entry["value"] = rv.Value.ValueOrDefault;
        var ec = el.Patterns.ExpandCollapse.PatternOrDefault;
        if (ec?.ExpandCollapseState.ValueOrDefault == ExpandCollapseState.Expanded) entry["expanded"] = true;
        if (c.HasKeyboardFocus) entry["focused"] = true;
        if (!c.IsEnabled) entry["disabled"] = true;

        actions.Add(entry);
        cache[aid] = new { runtimeId = c.RuntimeId, automationId = c.AutomationId,
                           label, role, x = cx, y = cy };
        if (++aid > 150) break;
    }

    foreach (var chunk in ChunkText(docText, 300))            // document body as texts
        if (tid <= 100) texts.Add(new { id = tid++, role = "paragraph", text = chunk });

    WriteSidecarCache(appQuery, win.Properties.NativeWindowHandle.Value, cache);
    return new { app = win.Name, window = WindowGeom(win), actions, texts };
}
```

### 6.2 Action dispatcher

```csharp
static object DoClick(UIA3Automation automation, Window win, string labelOrId)
{
    var el = ResolveElement(automation, win, labelOrId);       // §4: RuntimeId → AutomationId → label
    if (el == null) return new { error = $"Element '{labelOrId}' not found" };
    if (el.Properties.IsOffscreen.ValueOrDefault)
        return new { error = $"Element '{labelOrId}' is offscreen — re-run app_read_state" };

    string role = MapRole(el.ControlType, el);
    var inv = el.Patterns.Invoke.PatternOrDefault;
    if (inv != null)  { inv.Invoke();  return Ok("invoke", el, role); }
    var tog = el.Patterns.Toggle.PatternOrDefault;
    if (tog != null)  { tog.Toggle();  return Ok("toggle", el, role); }
    var sel = el.Patterns.SelectionItem.PatternOrDefault;
    if (sel != null)  { sel.Select();  return Ok("select", el, role); }
    var ec  = el.Patterns.ExpandCollapse.PatternOrDefault;
    if (ec != null)   { ec.Expand();   return Ok("expand", el, role); }
    var acc = el.Patterns.LegacyIAccessible.PatternOrDefault;  // MSAA bridge — Hancom path
    if (acc != null)  { acc.DoDefaultAction(); return Ok("dodefaultaction", el, role); }

    var r = el.BoundingRectangle;                              // live, not cached
    Win32.SetForegroundWindow(win.Properties.NativeWindowHandle.Value);
    FlaUI.Core.Input.Mouse.Click(new Point(r.X + r.Width / 2, r.Y + r.Height / 2));
    return new { ok = true, action = "mouse", label = el.Name, role };
}

static object DoSetText(UIA3Automation automation, Window win, string labelOrId, string value)
{
    var el = ResolveElement(automation, win, labelOrId);
    if (el == null) return new { error = $"Text field '{labelOrId}' not found" };

    var vp = el.Patterns.Value.PatternOrDefault;
    if (vp != null && !vp.IsReadOnly.ValueOrDefault)
    {
        vp.SetValue(value);
        var back = vp.Value.ValueOrDefault ?? ReadViaTextPattern(el);
        return new { ok = true, label = el.Name, value_set = value, readback = back, method = "value" };
    }
    if (el.Patterns.Text.IsSupported)                          // rich edit: focus + select-all + paste
    {
        el.Focus();
        Win32.SendChord(VK.CONTROL, VK.KEY_A);
        ClipboardPaste(value);                                 // save → set CF_UNICODETEXT → Ctrl+V → restore
        Thread.Sleep(150);
        return new { ok = true, label = el.Name, value_set = value,
                     readback = ReadViaTextPattern(el), method = "paste" };
    }
    return new { error = $"Element '{labelOrId}' is not editable (no Value/Text pattern)" };
}

static AutomationElement? ResolveElement(UIA3Automation a, Window win, string labelOrId)
{
    var cache = LoadSidecarCache(win);                         // null if missing/expired (TTL 45 s)
    if (TryParseId(labelOrId, out int id) && cache?.Get(id) is { } c)
    {
        var byRid = win.FindFirst(TreeScope.Descendants, new PropertyCondition(
            a.PropertyLibrary.Element.RuntimeId, c.RuntimeId));
        if (byRid != null) return byRid;
        if (!string.IsNullOrEmpty(c.AutomationId))
        {
            var byAid = win.FindFirst(TreeScope.Descendants,
                a.ConditionFactory.ByAutomationId(c.AutomationId));
            if (byAid != null) return byAid;
        }
        labelOrId = c.Label;                                   // fall through to label match
    }
    foreach (bool exact in new[] { true, false })              // exact then substring (Linux semantics)
    {
        var hits = win.FindAll(TreeScope.Descendants,
                a.ConditionFactory.ByName(labelOrId, PropertyConditionFlags.MatchSubstring | PropertyConditionFlags.IgnoreCase))
            .Where(e => exact ? Norm(e.Name) == Norm(labelOrId) : true).ToArray();
        if (hits.Length > 0) return PickNearestToCache(hits, cache);
    }
    return null;
}
```

(`Win32.SendChord`, `ClipboardPaste`, `SetForegroundWindow` are thin P/Invoke wrappers;
`PropertyConditionFlags.MatchSubstring` requires Win10 1809+ — on older builds do the
substring filter client-side over a cached Name `FindAll`.)

---

## 7. App-specific notes

### Microsoft Word

- **Tree quality: good.** Office has first-class UIA providers; names, control types,
  patterns are reliable. Main window class `OpusApp`, document canvas class `_WwG`,
  ribbon hosted in `NetUIHWND` panes.
- **Ribbon is huge**: rely on the `IsOffscreen==false` condition (inactive tabs prune
  themselves) plus the 60-action ribbon cap (§2.7.6). Ribbon tab headers are `TabItem`s —
  clicking a tab then re-harvesting is the natural agent flow for reaching other tabs.
- **Document body: TextPattern only.** Never harvest the document subtree as tree text
  nodes — Word exposes every line/run and it explodes the budget. Use
  `Document → TextPattern.DocumentRange.GetText(N)` for reading (§2.6). For *writing*,
  prefer the COM path (`Word.Application` object model, see `com-toolset-design.md`);
  UIA `set_text` on the body degrades to select-all+paste which destroys formatting.
- Dialogs (Find/Replace, Save As) are ordinary Win32/UIA — the generic harvest works.

### Hancom HWP (한글)

- **UIA support is weak/partial.** Much of HWP's chrome is custom-drawn; the native UIA
  provider is thin, and most controls surface only through the **MSAA→UIA bridge**, i.e.
  as `Pane`/`Custom` elements whose only useful pattern is **`LegacyIAccessiblePattern`**
  (`Role`, `Name`, `DoDefaultAction`, `accValue`). Names may be Korean-only or missing.
- Harvest consequences: the MSAA fallback promotion in §2.2 matters here — when
  ControlType is `Pane`/`Custom` but `LegacyIAccessible.Role` is
  `ROLE_SYSTEM_PUSHBUTTON/CHECKBUTTON/COMBOBOX/TEXT/...`, classify by the MSAA role and
  emit the corresponding AT-SPI-style role string. Expect the fallback BFS walker
  (§2.7.5) to trigger more often; budget accordingly.
- Action consequences: the click chain will usually land on
  `LegacyIAccessible.DoDefaultAction()` or the mouse fallback. Verify effects by
  re-harvest, not by pattern state.
- **Content editing must go through COM**, not UIA: HWP exposes the
  `HWPFrame.HwpObject` COM automation object (HwpAutomation / HwpCtrl API) with proper
  text insertion, field access, and document manipulation. The UIA layer is for chrome
  (menus, dialogs, toolbars) only. See `com-toolset-design.md` for the HWP COM toolset;
  the tool router should send `app_fill` targeting the HWP editing canvas to the COM
  path automatically.
- HWP's security module dialogs (파일 접근 보안) pop modally; the harvest of top-level
  windows must include owned dialog windows of the process so the agent can see and
  dismiss them.

### LibreOffice (Writer)

- **Decent bridge.** LibreOffice on Windows implements IAccessible2; UIA clients see it
  through the MSAA/IA2→UIA proxy. Names and roles come through well; patterns are
  spottier than Office (Invoke often absent → `LegacyIAccessible.DoDefaultAction` is the
  workhorse; ValuePattern usually present on toolbars' combo fields).
- Document body appears as a large `Document`/`Pane` with per-paragraph children — apply
  the same skip-and-TextPattern rule; if TextPattern is unavailable through the bridge,
  fall back to `LegacyIAccessible.Value` per paragraph capped at 100 texts, or prefer the
  UNO/COM path for content.
- Requires LO's accessibility support active; a UIA client connecting is sufficient to
  activate the bridge (no env vars needed, unlike the GTK/AT-SPI side on Linux).

---

## 8. Risks and open questions

**Risks**

1. **Foreground/focus restrictions.** `SendInput` and `SetForegroundWindow` are subject to
   foreground-lock rules; if the sidecar's caller isn't the foreground process, focus
   stealing can silently fail and keys land in the wrong window. Mitigation: verify
   `GetForegroundWindow()` after the attempt and error out loudly; prefer pattern-based
   verbs which don't need focus.
2. **UIPI / elevation.** A non-elevated sidecar cannot automate elevated apps (or the
   secure desktop for UAC prompts). Mitigation: detect `IsProcessElevated` mismatch and
   return a specific error string the agent can relay.
3. **DPI scaling.** `BoundingRectangle` is in physical pixels only if the sidecar is
   Per-Monitor-V2 DPI aware; otherwise coordinates are virtualized and mouse fallbacks
   miss. Mitigation: DPI-awareness manifest on `uia-tool.exe`; all coordinate math in
   physical pixels.
4. **RuntimeId instability.** RuntimeIds change when a control is rebuilt (virtualized
   lists, ribbon relayout). The label fallback covers this, but ambiguous labels + rebuilt
   controls can mis-target. Mitigation: nearest-to-cached-coordinates tie-break and the
   `"matches": n` disclosure.
5. **Clipboard races.** Save/restore around paste can collide with clipboard managers
   (Win+V history) or user copy actions. Mitigation: keep the window small (<200 ms),
   retry `OpenClipboard` with backoff, and document that clipboard history will record
   the payload (privacy note for the writing assistant).
6. **Hancom version drift.** MSAA exposure and the HwpObject COM surface differ across
   HWP 2018/2020/2022/2024. Needs a per-version smoke-test matrix.
7. **Per-action process spawn cost.** CLI mode pays COM init + window resolve (~50–150 ms)
   per call, and the sidecar cache mitigates re-walks but not init. If tool-call latency
   matters, promote to the persistent stdio daemon (§4) — the verb contract already
   allows it.
8. **AV/signing.** An unsigned exe that injects input and reads other apps' UI trees
   looks like a RAT to EDR products. Code-sign `uia-tool.exe` with the app's certificate
   and ship it inside the signed installer.

**Open questions**

1. Does the writing assistant need to drive **UWP/WinUI3** targets (e.g. new Outlook,
   Notepad 11)? They work over UIA but window enumeration and `app_open` (AUMID launch)
   need the UWP paths in §5.2 finished.
2. **Daemon vs CLI at v1** — start CLI-only (drop-in for desktop-apps.ts) and measure; if
   Word action latency exceeds ~400 ms p50, ship the daemon.
3. Should `texts` for Word/HWP include the **selection** (`TextPattern.GetSelection()`)
   as a marked entry? Cheap and very useful for a writing assistant; proposed as
   `{"role": "selection", "text": ...}` — additive, doesn't break the schema.
4. Threshold for **routing `app_fill` to COM** automatically for Word/HWP document
   bodies vs making the agent choose an explicit `doc_*` COM toolset — to be settled
   jointly with `com-toolset-design.md`.
5. Localized UIA `Name`s: on Korean Windows, Word/HWP labels are Korean; agents prompted
   in English may guess English labels. Consider emitting `AutomationId` alongside the
   label (as `#id` suffix) for Office ribbon controls, whose AutomationIds are stable
   English tokens (`Bold`, `FontSizeEditor`).
