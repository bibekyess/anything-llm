/**
 * doc-tools.ts — pi extension exposing the docd sidecar's document tools.
 *
 * Slice 1: Word backend (doc_list_open/open/read/outline/insert/replace/
 * edit_range/apply_style/save/save_as/close). Same registerTool pattern as
 * open-computer's desktop-apps.ts, but instead of one-shot scripts we keep a
 * persistent JSON-RPC child process (COM state must live in one apartment —
 * see docs/windows-writing-assistant/com-toolset-design.md §1).
 */
import { spawn, type ChildProcess } from "child_process";
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { Type } from "typebox";

const PYTHON = process.env.DOCD_PYTHON || "python";
const DOCD_ARGS = ["-m", "docd"];
const DOCD_CWD = process.env.DOCD_CWD || __dirname + "/..";
const CALL_TIMEOUT_MS = 30_000;

interface RpcError {
  code: string;
  message: string;
  data?: Record<string, unknown>;
}

class SidecarError extends Error {
  constructor(public rpc: RpcError) {
    super(rpc.message);
  }
}

/** Persistent line-delimited JSON-RPC client over the sidecar's stdio. */
class Sidecar {
  private proc: ChildProcess | null = null;
  private pending = new Map<
    string,
    { resolve: (v: any) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }
  >();
  private nextId = 1;
  private buffer = "";

  private ensure(): ChildProcess {
    if (this.proc && this.proc.exitCode === null) return this.proc;
    this.proc = spawn(PYTHON, DOCD_ARGS, {
      cwd: DOCD_CWD,
      stdio: ["pipe", "pipe", "pipe"],
      env: {
        ...process.env,
        // The RPC pipe is UTF-8; without these, Windows Python decodes stdin
        // with the ANSI code page and shreds Korean/CJK text into surrogates.
        PYTHONUTF8: "1",
        PYTHONIOENCODING: "utf-8",
      },
    });
    this.proc.stdout!.on("data", (chunk: Buffer) => this.onData(chunk));
    this.proc.stderr!.on("data", () => {}); // sidecar logs; keep the pipe drained
    this.proc.on("exit", () => {
      // All handles are now invalid; fail in-flight calls with a clear message.
      for (const [, p] of this.pending) {
        clearTimeout(p.timer);
        p.reject(
          new SidecarError({
            code: "SIDECAR_RESTARTED",
            message:
              "The document sidecar exited; open-document handles are invalid. Re-open documents with doc_open.",
          })
        );
      }
      this.pending.clear();
      this.proc = null;
    });
    return this.proc;
  }

  private onData(chunk: Buffer) {
    this.buffer += chunk.toString("utf-8");
    let nl: number;
    while ((nl = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, nl).trim();
      this.buffer = this.buffer.slice(nl + 1);
      if (!line) continue;
      let msg: any;
      try {
        msg = JSON.parse(line);
      } catch {
        continue;
      }
      if (msg.event) continue; // unsolicited notifications (dialog watchdog, §8) — slice 2
      const p = msg.id != null ? this.pending.get(String(msg.id)) : undefined;
      if (!p) continue;
      this.pending.delete(String(msg.id));
      clearTimeout(p.timer);
      if (msg.error) p.reject(new SidecarError(msg.error as RpcError));
      else p.resolve(msg.result);
    }
  }

  call(method: string, params: Record<string, unknown>, timeoutMs = CALL_TIMEOUT_MS): Promise<any> {
    const proc = this.ensure();
    const id = String(this.nextId++);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(
          new SidecarError({
            code: "TIMEOUT",
            message: `${method} timed out after ${timeoutMs}ms — the app may be showing a modal dialog.`,
          })
        );
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      proc.stdin!.write(JSON.stringify({ id, method, params }) + "\n");
    });
  }
}

const sidecar = new Sidecar();

function errText(err: unknown): string {
  if (err instanceof SidecarError) {
    let text = `[${err.rpc.code}] ${err.rpc.message}`;
    const ctx = err.rpc.data?.context as
      | Array<{ para: number; hash: string; text: string }>
      | undefined;
    if (ctx?.length) {
      text +=
        "\nCurrent content near the anchor:\n" +
        ctx.map((c) => `  [p${c.para}#${c.hash}] ${c.text}`).join("\n");
    }
    return text;
  }
  return String(err);
}

function affectedText(res: any): string {
  const list = (res.affected ?? []) as Array<[number, string]>;
  return list.length ? "Affected: " + list.map(([i, h]) => `[p${i}#${h}]`).join(" ") : "";
}

type ToolResult = { content: Array<{ type: "text"; text: string }>; details: Record<string, unknown> };

function ok(text: string, details: any = {}): ToolResult {
  return { content: [{ type: "text", text }], details };
}

export default function (pi: ExtensionAPI) {
  /** Approval gate for destructive ops (design doc §8), via the ask-user channel. */
  async function confirm(ctx: any, question: string): Promise<boolean> {
    try {
      const answer = await ctx.ui.input({ message: `${question} (yes/no)` });
      return /^y(es)?$/i.test(String(answer ?? "").trim());
    } catch {
      return false; // no UI channel -> refuse destructive ops
    }
  }

  pi.registerTool({
    name: "doc_list_open",
    label: "List Open Documents",
    description:
      "List documents currently open in supported apps (Word). Returns a handle per document (e.g. w1) used by all other doc_* tools.",
    parameters: Type.Object({}),
    async execute() {
      try {
        const res = await sidecar.call("doc_list_open", {});
        if (!res.docs.length) return ok("No documents open.");
        const lines = res.docs.map(
          (d: any) =>
            `[${d.doc}] ${d.backend}  ${d.path}  (${d.dirty ? "dirty" : "saved"}, ${d.paragraphs} paras${d.read_only ? ", read-only" : ""})`
        );
        return ok("Open documents:\n" + lines.join("\n"), res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_open",
    label: "Open Document",
    description:
      "Open a document (or attach to it if already open). The hosting app is chosen by extension (.docx/.doc → Word). The window stays visible so the user can watch edits.",
    parameters: Type.Object({
      path: Type.String({ description: "Absolute path to the document." }),
      app: Type.Optional(
        Type.Union([Type.Literal("word"), Type.Literal("fake")], {
          description: "Force a hosting app.",
        })
      ),
      read_only: Type.Optional(Type.Boolean({ description: "Open without write access." })),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_open", params);
        return ok(
          `Opened [${res.doc}] ${res.path} — ${res.paragraphs} paragraphs${res.read_only ? " (read-only)" : ""}. Use doc_read to see content with paragraph ids.`,
          res
        );
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_new",
    label: "New Document",
    description:
      "Create a blank document in a visible window (e.g. to write a fresh report into). Returns a handle; the document is unsaved until doc_save_as.",
    parameters: Type.Object({
      app: Type.Optional(
        Type.Union([Type.Literal("word"), Type.Literal("fake")], {
          description: "Hosting app (default: Word when available).",
        })
      ),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_new", params);
        return ok(
          `Created blank document [${res.doc}] (unsaved — use doc_save_as when done). Write into it with doc_insert.`,
          res
        );
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_selection",
    label: "Read User's Selection",
    description:
      "Read the text the user currently has selected in the document window, with its paragraph range and hashes — use this when the user says 'this', 'the selected text', etc. Then edit via doc_edit_range or doc_tables create with replace_range.",
    parameters: Type.Object({ doc: Type.String() }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_selection", params);
        if (res.collapsed) {
          return ok(
            "Nothing is selected (the caret is just placed in the document). Ask the user to select the text, or use doc_read to locate it.",
            res
          );
        }
        return ok(
          `Selected p${res.from_para}..p${res.to_para} (hashes: ${res.hashes.join(", ")}):\n${res.text}`,
          res
        );
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_tables",
    label: "Document Tables",
    description:
      "Work with tables. op=list (all tables), read (cells as pipe rows), write (set a cell, or a 2-D block via `values` starting at `cell`), create (build a real table from `values`; anchor with at=end/before_para/after_para+para, or pass replace_range {from_para,to_para,expect_hashes} to REPLACE those paragraphs with the table — the 'convert this text to a table' flow).",
    parameters: Type.Object({
      doc: Type.String(),
      op: Type.Union([
        Type.Literal("list"),
        Type.Literal("read"),
        Type.Literal("write"),
        Type.Literal("create"),
      ]),
      table: Type.Optional(Type.String({ description: "Table id from list/read, e.g. 't0'." })),
      cell: Type.Optional(
        Type.Object({ row: Type.Number(), col: Type.Number() }, { description: "0-based, for write." })
      ),
      value: Type.Optional(Type.String({ description: "Single-cell write value." })),
      values: Type.Optional(
        Type.Array(Type.Array(Type.String()), {
          description: "2-D rows×cols block: cell contents for create, or block write starting at `cell`.",
        })
      ),
      at: Type.Optional(
        Type.Union([Type.Literal("end"), Type.Literal("before_para"), Type.Literal("after_para")], {
          description: "Anchor for create (default end).",
        })
      ),
      para: Type.Optional(Type.Number()),
      expect_hash: Type.Optional(Type.String()),
      replace_range: Type.Optional(
        Type.Object({
          from_para: Type.Number(),
          to_para: Type.Number(),
          expect_hashes: Type.Array(Type.String()),
        })
      ),
      header_row: Type.Optional(Type.Boolean({ description: "Bold + repeat first row as header (create)." })),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_tables", params);
        if (params.op === "list") {
          const tables = res.tables as Array<any>;
          if (!tables.length) return ok("No tables in this document.", res);
          return ok(
            tables.map((t) => `[${t.table}] ${t.rows}x${t.cols} at p${t.at_para}`).join("\n"),
            res
          );
        }
        if (params.op === "read") return ok(res.text, res);
        if (params.op === "write") return ok(`Wrote ${res.written} cell(s).`, res);
        const replaced = res.deleted_paras
          ? ` (replaced ${res.deleted_paras} source paragraph(s))`
          : "";
        return ok(
          `Created table [${res.table}] ${res.rows}x${res.cols} at p${res.at_para}${replaced}.`,
          res
        );
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_read",
    label: "Read Document",
    description:
      "Read a document as text with stable paragraph ids like [p4#a1b2] (index + content hash). Mutating tools need these hashes. For long documents read doc_outline first, then ranges.",
    parameters: Type.Object({
      doc: Type.String({ description: "Document handle from doc_open/doc_list_open." }),
      from_para: Type.Optional(Type.Number({ description: "First paragraph index (0-based)." })),
      to_para: Type.Optional(Type.Number({ description: "Last paragraph index, inclusive." })),
      max_chars: Type.Optional(Type.Number({ description: "Truncate output (default 20000)." })),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_read", params);
        return ok(res.text, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_outline",
    label: "Document Outline",
    description: "Get the heading tree of a document with paragraph indices.",
    parameters: Type.Object({ doc: Type.String() }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_outline", params);
        return ok(res.text, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_insert",
    label: "Insert Text",
    description:
      "Insert text into a document. '\\n' starts a new paragraph. Markdown becomes REAL formatting: # headings -> Heading styles, **bold**, *italic*, `code`, -/* bullets, 1. numbered lists, [ ]/[x] checkboxes, > quotes — so write normal markdown, never raw markers meant to be visible. Anchor with where=end/cursor, or before_para/after_para + para index (+ expect_hash from doc_read to guard against concurrent edits), or a Word bookmark. The user sees the edit live; it is undoable with Ctrl+Z.",
    parameters: Type.Object({
      doc: Type.String(),
      text: Type.String(),
      where: Type.Union([
        Type.Literal("end"),
        Type.Literal("cursor"),
        Type.Literal("before_para"),
        Type.Literal("after_para"),
        Type.Literal("bookmark"),
      ]),
      para: Type.Optional(Type.Number()),
      expect_hash: Type.Optional(Type.String()),
      bookmark: Type.Optional(Type.String()),
      style_map: Type.Optional(Type.Boolean({ description: "Map '#' prefixes to Heading styles (default true)." })),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_insert", params);
        const moved = res.moved ? " (anchor had moved; re-anchored by content hash)" : "";
        return ok(
          `Inserted ${res.inserted} paragraph(s) at p${res.first_para}${moved}. ${affectedText(res)}`,
          res
        );
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_replace",
    label: "Find & Replace",
    description:
      "Find and replace text in a document through the app's own object model (live, undoable). Prefer this over doc_edit_range for small textual changes. regex uses Python syntax (emulated per paragraph — matches cannot span paragraphs).",
    parameters: Type.Object({
      doc: Type.String(),
      find: Type.String(),
      replace: Type.String(),
      regex: Type.Optional(Type.Boolean()),
      match_case: Type.Optional(Type.Boolean()),
      occurrence: Type.Optional(
        Type.Union([Type.Literal("all"), Type.Literal("first"), Type.Number()], {
          description: "'all' (default), 'first', or the 1-based Nth occurrence.",
        })
      ),
      scope: Type.Optional(
        Type.Object(
          { from_para: Type.Number(), to_para: Type.Number() },
          { description: "Restrict to a paragraph range." }
        )
      ),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_replace", params);
        return ok(`Replaced ${res.replaced} occurrence(s). ${affectedText(res)}`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_edit_range",
    label: "Rewrite Paragraph Range",
    description:
      "Surgically rewrite (or delete, with empty new_text) a paragraph range. Requires expect_hashes for EVERY paragraph in the range from the latest doc_read — refused as STALE_RANGE if the user changed them meanwhile.",
    parameters: Type.Object({
      doc: Type.String(),
      from_para: Type.Number(),
      to_para: Type.Number(),
      expect_hashes: Type.Array(Type.String()),
      new_text: Type.String({ description: "'\\n' separates paragraphs; empty deletes the range. Markdown formatting (headings, **bold**, *italic*, lists, checkboxes) is rendered as real formatting." }),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_edit_range", params);
        const what = res.deleted
          ? `Deleted ${res.deleted} paragraph(s).`
          : `Rewrote range into ${res.replaced} paragraph(s).`;
        return ok(`${what} ${affectedText(res)}`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_apply_style",
    label: "Apply Paragraph Style",
    description:
      "Apply a named paragraph style ('Heading 1'…'Heading 9', 'Normal', 'Title', 'Quote', 'List Bullet', 'List Number') to a paragraph range.",
    parameters: Type.Object({
      doc: Type.String(),
      from_para: Type.Number(),
      to_para: Type.Optional(Type.Number()),
      style: Type.String(),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_apply_style", params);
        return ok(`Styled ${res.styled} paragraph(s). ${affectedText(res)}`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_save",
    label: "Save Document",
    description: "Save the document, overwriting its file. Asks the user for confirmation first.",
    parameters: Type.Object({ doc: Type.String() }),
    async execute(_id, params, _signal, _onUpdate, ctx) {
      if (!(await confirm(ctx, `Save ${params.doc} and overwrite the file on disk?`))) {
        return ok("Save cancelled by user.");
      }
      try {
        const res = await sidecar.call("doc_save", params);
        return ok(`Saved ${res.path}.`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_save_as",
    label: "Save Document As",
    description: "Save a copy of the document to a new path/format (docx, doc, pdf, odt, txt).",
    parameters: Type.Object({
      doc: Type.String(),
      path: Type.String(),
      format: Type.Union([
        Type.Literal("docx"),
        Type.Literal("doc"),
        Type.Literal("pdf"),
        Type.Literal("odt"),
        Type.Literal("txt"),
      ]),
    }),
    async execute(_id, params) {
      try {
        const res = await sidecar.call("doc_save_as", params);
        return ok(`Saved as ${res.path} (${res.format}).`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });

  pi.registerTool({
    name: "doc_close",
    label: "Close Document",
    description:
      "Close a document. Unsaved changes are saved unless discard_changes=true (which asks the user for confirmation).",
    parameters: Type.Object({
      doc: Type.String(),
      discard_changes: Type.Optional(Type.Boolean()),
    }),
    async execute(_id, params, _signal, _onUpdate, ctx) {
      if (
        params.discard_changes &&
        !(await confirm(ctx, `Close ${params.doc} and DISCARD unsaved changes?`))
      ) {
        return ok("Close cancelled by user.");
      }
      try {
        const res = await sidecar.call("doc_close", params);
        return ok(`Closed${res.was_dirty ? " (had unsaved changes)" : ""}.`, res);
      } catch (e) {
        return ok(errText(e));
      }
    },
  });
}
