import { useEffect, useRef, useState } from "react";
import ReactMarkdown from "react-markdown";

type Item =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string; streaming?: boolean }
  | { kind: "tool"; tool: string; summary: string; done: boolean; preview?: string }
  | { kind: "ask"; requestId: string; method: string; question: string; answered?: string }
  | { kind: "error"; text: string };

interface UsageInfo {
  contextTokens: number;
  contextWindow: number;
  percent: number;
  cost: number | null; // null = model pricing unknown
}

function fmtTokens(n: number): string {
  if (n >= 1_000_000) return (n / 1_000_000).toFixed(1) + "M";
  if (n >= 1000) return (n / 1000).toFixed(n >= 100_000 ? 0 : 1) + "k";
  return String(n);
}

function fmtCost(c: number | null): string {
  if (c === null) return "$—";
  if (c === 0) return "$0.00";
  return c >= 0.1 ? `$${c.toFixed(2)}` : `$${c.toFixed(4)}`;
}

interface SessionMeta {
  id: string;
  title: string;
  updated: number;
}

const TOOL_LABELS: Record<string, string> = {
  doc_new: "Creating document",
  doc_open: "Opening document",
  doc_list_open: "Checking open documents",
  doc_read: "Reading document",
  doc_outline: "Reading outline",
  doc_selection: "Reading your selection",
  doc_insert: "Writing",
  doc_replace: "Replacing text",
  doc_edit_range: "Rewriting passage",
  doc_apply_style: "Styling",
  doc_tables: "Working on table",
  doc_save: "Saving",
  doc_save_as: "Saving as",
  doc_close: "Closing document",
};

function replayToItems(events: any[]): Item[] {
  const items: Item[] = [];
  for (const e of events) {
    if (e.type === "user_message") items.push({ kind: "user", text: e.text });
    else if (e.type === "assistant_message") items.push({ kind: "assistant", text: e.text });
    else if (e.type === "tool_start")
      items.push({ kind: "tool", tool: e.tool, summary: e.summary, done: false });
    else if (e.type === "tool_end") {
      const open = [...items].reverse().find(
        (i) => i.kind === "tool" && i.tool === e.tool && !i.done,
      ) as Extract<Item, { kind: "tool" }> | undefined;
      if (open) {
        open.done = true;
        open.preview = e.preview;
      }
    } else if (e.type === "ask_user")
      items.push({ kind: "ask", requestId: e.requestId, method: e.method, question: e.question });
    else if (e.type === "fatal" || (e.type === "task_done" && e.error))
      items.push({ kind: "error", text: e.text || e.error });
  }
  return items;
}

const SUGGESTIONS = [
  "Write a one-page report on Korean culture in a new Word document",
  "Convert my selected text into a table",
  "Proofread this document and fix grammar",
];

function PinIcon({ off }: { off?: boolean }) {
  return (
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 17v5" />
      <path d="M9 10.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24V16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V7a1 1 0 0 1 1-1 2 2 0 0 0 0-4H8a2 2 0 0 0 0 4 1 1 0 0 1 1 1z" />
      {off && <line x1="3" y1="3" x2="21" y2="21" />}
    </svg>
  );
}

export default function Chat() {
  const api = (window as any).assistant;
  const [items, setItems] = useState<Item[]>([]);
  const [sessions, setSessions] = useState<SessionMeta[]>([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [input, setInput] = useState("");
  const [pinned, setPinned] = useState(true); // window starts alwaysOnTop
  const [usage, setUsage] = useState<UsageInfo | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);

  const togglePin = () => {
    const next = !pinned;
    setPinned(next);
    api.win.setPin(next);
  };

  const autoGrow = () => {
    const el = inputRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 120) + "px";
    // Only show a scrollbar once the input is taller than its max height.
    el.style.overflowY = el.scrollHeight > 120 ? "auto" : "hidden";
  };

  useEffect(() => {
    api.log("info", "chat view mounted");
    window.addEventListener("error", (e) =>
      api.log("error", `window error: ${e.message} @ ${e.filename}:${e.lineno}`),
    );
    window.addEventListener("unhandledrejection", (e: PromiseRejectionEvent) =>
      api.log("error", `unhandled rejection: ${e.reason}`),
    );
    api.sessions.list().then(setSessions);
    const off = api.onEvent((e: any) => {
      api.log("debug", `event: ${e.type}${e.tool ? ` (${e.tool})` : ""}`);
      setItems((prev) => {
        const next = [...prev];
        switch (e.type) {
          case "user_message":
            next.push({ kind: "user", text: e.text });
            break;
          case "assistant_delta": {
            const last = next[next.length - 1];
            if (last?.kind === "assistant" && last.streaming) {
              next[next.length - 1] = { ...last, text: last.text + e.text };
            } else {
              next.push({ kind: "assistant", text: e.text, streaming: true });
            }
            break;
          }
          case "assistant_message": {
            const last = next[next.length - 1];
            if (last?.kind === "assistant" && last.streaming) {
              next[next.length - 1] = { kind: "assistant", text: e.text };
            } else {
              next.push({ kind: "assistant", text: e.text });
            }
            break;
          }
          case "tool_start":
            next.push({ kind: "tool", tool: e.tool, summary: e.summary, done: false });
            break;
          case "tool_end": {
            for (let i = next.length - 1; i >= 0; i--) {
              const item = next[i];
              if (item.kind === "tool" && item.tool === e.tool && !item.done) {
                next[i] = { ...item, done: true, preview: e.preview };
                break;
              }
            }
            break;
          }
          case "ask_user":
            next.push({
              kind: "ask",
              requestId: e.requestId,
              method: e.method,
              question: e.question,
            });
            break;
          case "fatal":
            next.push({ kind: "error", text: e.text });
            break;
          case "task_done":
            if (e.error) next.push({ kind: "error", text: e.error });
            break;
        }
        return next;
      });
      if (e.type === "task_done" || e.type === "fatal") {
        setBusy(false);
        setStatus("");
      }
      if (e.type === "user_message") setBusy(true);
      if (e.type === "status") setStatus(e.text);
      if (e.type === "assistant_delta" || e.type === "assistant_message") setStatus("");
      if (e.type === "usage")
        setUsage({
          contextTokens: e.contextTokens,
          contextWindow: e.contextWindow,
          percent: e.percent,
          cost: e.cost,
        });
    });
    return off;
  }, []);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" });
  }, [items]);

  const send = () => {
    const text = input.trim();
    if (!text || busy) return;
    api.log("info", `sending prompt (${text.length} chars)`);
    setInput("");
    requestAnimationFrame(autoGrow);
    api.send(text);
  };

  const answer = (item: Extract<Item, { kind: "ask" }>, value: string) => {
    api.respond(item.requestId, value);
    setItems((prev) =>
      prev.map((i) =>
        i.kind === "ask" && i.requestId === item.requestId ? { ...i, answered: value } : i,
      ),
    );
  };

  const openSession = async (id: string) => {
    const events = await api.sessions.load(id);
    setItems(replayToItems(events));
    const lastUsage = [...events].reverse().find((e: any) => e.type === "usage");
    setUsage(lastUsage || null);
    setSidebarOpen(false);
    setBusy(false);
  };

  const newChat = async () => {
    await api.sessions.create();
    setItems([]);
    setUsage(null);
    setSessions(await api.sessions.list());
    setSidebarOpen(false);
    setBusy(false);
  };

  return (
    <div className="shell">
      <header className="titlebar">
        <button className="icon-btn no-drag" onClick={() => setSidebarOpen(!sidebarOpen)}>☰</button>
        <span className="title">Writing Assistant</span>
        <div className="win-controls no-drag">
          <button
            className={`icon-btn pin ${pinned ? "active" : ""}`}
            title={pinned ? "Unpin (stop floating above Word)" : "Pin above other windows"}
            onClick={togglePin}
          >
            <PinIcon off={!pinned} />
          </button>
          <button className="icon-btn" title="Minimize to orb" onClick={() => api.win.minimize()}>─</button>
          <button className="icon-btn close" title="Quit" onClick={() => api.win.close()}>✕</button>
        </div>
      </header>

      {sidebarOpen && (
        <aside className="sidebar glass">
          <button className="new-chat" onClick={newChat}>＋ New chat</button>
          <div className="session-list">
            {sessions.map((s) => (
              <button key={s.id} className="session" onClick={() => openSession(s.id)}>
                {s.title}
              </button>
            ))}
            {sessions.length === 0 && <div className="hint">No previous chats yet</div>}
          </div>
        </aside>
      )}

      <div className="messages" ref={scrollRef}>
        {items.length === 0 && (
          <div className="empty">
            <div className="empty-mark">✦</div>
            <h2>What should we write?</h2>
            <p>I work directly in your Word window — every edit lands live, and Ctrl+Z undoes it.</p>
            <div className="suggestions">
              {SUGGESTIONS.map((s) => (
                <button
                  key={s}
                  className="suggestion"
                  onClick={() => {
                    setInput(s);
                    inputRef.current?.focus();
                    requestAnimationFrame(autoGrow);
                  }}
                >
                  {s}
                </button>
              ))}
            </div>
          </div>
        )}
        {items.map((item, i) => {
          switch (item.kind) {
            case "user":
              return (
                <div key={i} className="bubble user glass">{item.text}</div>
              );
            case "assistant":
              return (
                <div key={i} className="bubble assistant glass">
                  <ReactMarkdown>{item.text}</ReactMarkdown>
                </div>
              );
            case "tool":
              return (
                <div key={i} className={`tool-chip glass ${item.done ? "done" : "running"}`}>
                  <span className="spinner">{item.done ? "✓" : "◌"}</span>
                  {TOOL_LABELS[item.tool] || item.tool}
                  {!item.done && "…"}
                </div>
              );
            case "ask":
              return (
                <div key={i} className="ask-card glass">
                  <div className="ask-q">{item.question}</div>
                  {item.answered !== undefined ? (
                    <div className="ask-answered">You: {item.answered || "(dismissed)"}</div>
                  ) : item.method === "confirm" || /\(yes\/no\)/i.test(item.question) ? (
                    <div className="ask-actions">
                      <button onClick={() => answer(item, "yes")}>Yes</button>
                      <button className="secondary" onClick={() => answer(item, "no")}>No</button>
                    </div>
                  ) : (
                    <AskInput onSubmit={(v) => answer(item, v)} />
                  )}
                </div>
              );
            case "error":
              return <div key={i} className="bubble error glass">⚠ {item.text}</div>;
          }
        })}
        {busy && <div className="working">{status || "working…"}</div>}
      </div>

      {usage && (
        <div className="statusbar">
          <div
            className="ctx"
            title={`Context: ${usage.contextTokens.toLocaleString()} of ${usage.contextWindow.toLocaleString()} tokens`}
          >
            <div className="ctx-track">
              <div
                className={`ctx-fill ${usage.percent > 90 ? "danger" : usage.percent > 75 ? "warn" : ""}`}
                style={{ width: `${Math.min(100, usage.percent)}%` }}
              />
            </div>
            <span>
              {fmtTokens(usage.contextTokens)} / {fmtTokens(usage.contextWindow)} ·{" "}
              {usage.percent.toFixed(0)}%
            </span>
          </div>
          <span
            className="cost"
            title={
              usage.cost === null
                ? "Model pricing unknown — set MODEL_COST_INPUT / MODEL_COST_OUTPUT ($ per million tokens) in agent/.env"
                : "Session cost so far"
            }
          >
            {fmtCost(usage.cost)}
          </span>
        </div>
      )}

      <footer className="composer glass">
        <textarea
          ref={inputRef}
          value={input}
          placeholder="Ask anything… (Enter to send, Shift+Enter for newline)"
          rows={1}
          onChange={(e) => {
            setInput(e.target.value);
            autoGrow();
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
        />
        {busy ? (
          <button className="send stop" title="Stop generating" onClick={() => api.stop()}>■</button>
        ) : (
          <button className="send" title="Send" onClick={send} disabled={!input.trim()}>➤</button>
        )}
      </footer>
    </div>
  );
}

function AskInput({ onSubmit }: { onSubmit: (v: string) => void }) {
  const [value, setValue] = useState("");
  return (
    <div className="ask-actions">
      <input
        value={value}
        placeholder="Type your answer…"
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && onSubmit(value)}
      />
      <button onClick={() => onSubmit(value)}>Send</button>
    </div>
  );
}
