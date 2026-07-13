import { useEffect, useRef, useState } from "react";
import ReactMarkdown from "react-markdown";

type Item =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string; streaming?: boolean }
  | { kind: "tool"; tool: string; summary: string; done: boolean; preview?: string }
  | { kind: "ask"; requestId: string; method: string; question: string; answered?: string }
  | { kind: "error"; text: string };

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

export default function Chat() {
  const api = (window as any).assistant;
  const [items, setItems] = useState<Item[]>([]);
  const [sessions, setSessions] = useState<SessionMeta[]>([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [input, setInput] = useState("");
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    api.sessions.list().then(setSessions);
    const off = api.onEvent((e: any) => {
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
      if (e.type === "task_done" || e.type === "fatal") setBusy(false);
      if (e.type === "user_message") setBusy(true);
    });
    return off;
  }, []);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" });
  }, [items]);

  const send = () => {
    const text = input.trim();
    if (!text || busy) return;
    setInput("");
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
    setSidebarOpen(false);
    setBusy(false);
  };

  const newChat = async () => {
    await api.sessions.create();
    setItems([]);
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
          <div className="empty glass">
            <h2>What should we write?</h2>
            <p>
              Try: <em>“Write a one-page report on Korean culture in a new Word document”</em>
              {" "}or select text in Word and ask{" "}
              <em>“convert my selected text into a table”.</em>
            </p>
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
        {busy && <div className="working">working…</div>}
      </div>

      <footer className="composer glass">
        <textarea
          value={input}
          placeholder="Ask anything… (Enter to send, Shift+Enter for newline)"
          rows={1}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
        />
        <button className="send" onClick={send} disabled={!input.trim() || busy}>➤</button>
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
