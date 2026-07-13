/**
 * PiSession — spawns `pi --mode rpc` and adapts its JSON-lines events into
 * app events for the renderer. This is the ONLY module that knows pi's RPC
 * schema (mirrored from open-computer's session/hypervisor.js); if a pi
 * upgrade changes event shapes, fix them here.
 *
 * Raw events are also appended to <userData>/pi-raw.log for debugging.
 */
import { ChildProcess, spawn } from "child_process";
import * as fs from "fs";
import * as path from "path";
import {
  ASSISTANT_ROOT,
  PROVIDER_NAME,
  SYSTEM_PROMPT,
  loadLlmConfig,
  writeModelsJson,
} from "./config";

export type AppEvent =
  | { type: "user_message"; text: string }
  | { type: "assistant_delta"; text: string }
  | { type: "assistant_message"; text: string }
  | { type: "tool_start"; tool: string; summary: string }
  | { type: "tool_end"; tool: string; preview: string }
  | { type: "ask_user"; requestId: string; method: string; question: string; title: string }
  | { type: "task_done"; error?: string }
  | { type: "status"; text: string }
  | { type: "fatal"; text: string };

export class PiSession {
  private proc: ChildProcess | null = null;
  private buffer = "";
  private rawLog: fs.WriteStream | null = null;

  constructor(
    public sessionId: string,
    private userDataDir: string,
    private onEvent: (e: AppEvent) => void,
  ) {}

  get running(): boolean {
    return !!this.proc && this.proc.exitCode === null;
  }

  sendPrompt(text: string): void {
    this.ensureProcess();
    if (!this.proc) return;
    this.onEvent({ type: "user_message", text });
    this.proc.stdin!.write(
      JSON.stringify({
        id: `prompt-${Date.now()}`,
        type: "prompt",
        message: text,
        streamingBehavior: "followUp",
      }) + "\n",
    );
  }

  /** Answer an ask_user / confirm request raised by an extension. */
  respond(requestId: string, value: string): void {
    if (!this.proc) return;
    this.proc.stdin!.write(
      JSON.stringify({ type: "extension_ui_response", id: requestId, value }) + "\n",
    );
  }

  stop(): void {
    if (this.proc) {
      try {
        this.proc.kill();
      } catch {}
      this.proc = null;
    }
    this.rawLog?.end();
    this.rawLog = null;
  }

  // ── internals ────────────────────────────────────────────────────────
  private ensureProcess(): void {
    if (this.running) return;
    const cfg = loadLlmConfig();
    if ("error" in cfg) {
      this.onEvent({ type: "fatal", text: cfg.error });
      return;
    }
    writeModelsJson(cfg);

    const args = [
      "--mode", "rpc",
      "--provider", PROVIDER_NAME,
      "--model", cfg.model,
      "--session-id", this.sessionId,
      "--approve",
      "--extension", path.join(ASSISTANT_ROOT, "extension", "doc-tools.ts"),
      "--system-prompt", SYSTEM_PROMPT,
    ];

    this.rawLog = fs.createWriteStream(path.join(this.userDataDir, "pi-raw.log"), { flags: "a" });
    this.proc = spawn("pi", args, {
      cwd: ASSISTANT_ROOT,
      shell: process.platform === "win32", // resolve pi.cmd from npm on PATH
      stdio: ["pipe", "pipe", "pipe"],
      env: {
        ...process.env,
        DOCD_PYTHON: process.env.DOCD_PYTHON || "python",
        DOCD_CWD: ASSISTANT_ROOT,
        PYTHONUTF8: "1",
        PYTHONIOENCODING: "utf-8",
        ...(cfg.apiKey ? { OPENAI_API_KEY: cfg.apiKey } : {}),
      },
    });

    this.proc.on("error", (err) => {
      this.onEvent({
        type: "fatal",
        text:
          `Could not launch the 'pi' agent (${err.message}). Install it with: ` +
          `npm install -g --ignore-scripts @earendil-works/pi-coding-agent`,
      });
    });
    this.proc.on("exit", (code) => {
      if (code !== 0 && code !== null) {
        this.onEvent({ type: "fatal", text: `Agent process exited (code ${code}).` });
      }
      this.proc = null;
    });
    this.proc.stderr!.on("data", (chunk: Buffer) => {
      this.rawLog?.write(`[stderr] ${chunk.toString("utf-8")}`);
    });
    this.proc.stdout!.on("data", (chunk: Buffer) => {
      this.buffer += chunk.toString("utf-8");
      let nl: number;
      while ((nl = this.buffer.indexOf("\n")) >= 0) {
        const line = this.buffer.slice(0, nl).trim();
        this.buffer = this.buffer.slice(nl + 1);
        if (line) this.handleLine(line);
      }
    });
  }

  private textBuf = "";

  private flushTextBuf(): void {
    if (this.textBuf.trim()) {
      this.onEvent({ type: "assistant_delta", text: this.textBuf });
      this.textBuf = "";
    }
  }

  /** Event field names mirror session/hypervisor.js:handleRpcEvent. */
  private handleLine(line: string): void {
    this.rawLog?.write(line + "\n");
    let event: any;
    try {
      event = JSON.parse(line);
    } catch {
      return; // non-JSON chatter on stdout
    }
    const etype = event.type || event.event;
    switch (etype) {
      case "message_start": {
        const msg = event.message;
        if (msg?.role === "assistant" && Array.isArray(msg.content)) {
          for (const part of msg.content) {
            if (part.type === "text" && part.text) {
              this.onEvent({ type: "assistant_message", text: part.text });
            }
          }
        }
        break;
      }
      case "message_update": {
        const ae = event.assistantMessageEvent;
        if (ae?.type === "text_delta" && ae.delta) {
          this.textBuf += ae.delta;
          this.onEvent({ type: "assistant_delta", text: ae.delta });
        }
        break;
      }
      case "message_end":
        this.textBuf = "";
        break;
      case "tool_execution_start": {
        this.textBuf = "";
        if (event.toolName) {
          const argsJson = event.args ? JSON.stringify(event.args) : "";
          this.onEvent({
            type: "tool_start",
            tool: event.toolName,
            summary: argsJson.slice(0, 200),
          });
        }
        break;
      }
      case "tool_execution_end": {
        if (event.toolName) {
          let text = "";
          for (const part of event.result?.content || []) {
            if (part.type === "text" && part.text) text += part.text;
          }
          this.onEvent({
            type: "tool_end",
            tool: event.toolName,
            preview: text.slice(0, 400),
          });
        }
        break;
      }
      case "extension_ui_request": {
        if (["input", "confirm", "select"].includes(event.method)) {
          this.onEvent({
            type: "ask_user",
            requestId: event.id,
            method: event.method,
            question:
              event.placeholder || event.message || event.prompt ||
              "The agent has a question.",
            title: event.title || "Agent question",
          });
        }
        break;
      }
      case "response": {
        this.onEvent({
          type: "task_done",
          error: event.success === false && event.error ? String(event.error) : undefined,
        });
        break;
      }
      case "streaming_start":
        this.onEvent({ type: "status", text: "thinking" });
        break;
      default:
        break;
    }
  }
}
