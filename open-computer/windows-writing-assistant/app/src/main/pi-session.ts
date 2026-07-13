/**
 * PiSession — spawns `pi --mode rpc` and adapts its JSON-lines events into
 * app events for the renderer. This is the ONLY module that knows pi's RPC
 * schema (mirrored from open-computer's session/hypervisor.js); if a pi
 * upgrade changes event shapes, fix them here.
 *
 * Raw events are also appended to <userData>/pi-raw.log for debugging.
 */
import { ChildProcess, spawn, spawnSync } from "child_process";
import * as fs from "fs";
import * as path from "path";
import {
  ASSISTANT_ROOT,
  PROVIDER_NAME,
  SYSTEM_PROMPT,
  loadLlmConfig,
  resolveModelCost,
  writeModelsJson,
} from "./config";
import log from "./logger";

/**
 * Resolve how to launch pi WITHOUT a shell. On Windows the global `pi` is a
 * .cmd shim (not directly spawnable, and shell:true mangles multi-line args
 * like --system-prompt), so we locate the package's real JS entry next to
 * the shim and run it with node directly.
 */
function resolvePiCommand(): { cmd: string; argv0: string[] } | { error: string } {
  if (process.platform !== "win32") {
    return { cmd: "pi", argv0: [] };
  }
  const where = spawnSync("where", ["pi"], { encoding: "utf-8" });
  const shim = (where.stdout || "")
    .split(/\r?\n/)
    .find((line) => line.trim().toLowerCase().endsWith(".cmd"));
  if (!shim) {
    return {
      error:
        "The 'pi' agent CLI was not found on PATH. Install it with: " +
        "npm install -g --ignore-scripts @earendil-works/pi-coding-agent " +
        "(then restart the app so PATH refreshes).",
    };
  }
  const pkgDir = path.join(
    path.dirname(shim.trim()),
    "node_modules",
    "@earendil-works",
    "pi-coding-agent",
  );
  try {
    const pkg = JSON.parse(fs.readFileSync(path.join(pkgDir, "package.json"), "utf-8"));
    const bin =
      typeof pkg.bin === "string" ? pkg.bin : pkg.bin?.pi || Object.values(pkg.bin || {})[0];
    if (!bin) return { error: `No bin entry in ${pkgDir}/package.json` };
    const entry = path.join(pkgDir, bin as string);
    if (!fs.existsSync(entry)) return { error: `pi entry script not found: ${entry}` };
    log.info(`[pi] resolved shim ${shim.trim()} -> node ${entry}`);
    return { cmd: "node", argv0: [entry] };
  } catch (err: any) {
    return { error: `Could not resolve pi's entry script from ${pkgDir}: ${err.message}` };
  }
}

export type AppEvent =
  | { type: "user_message"; text: string }
  | { type: "assistant_delta"; text: string }
  | { type: "assistant_message"; text: string }
  | { type: "tool_start"; tool: string; summary: string }
  | { type: "tool_end"; tool: string; preview: string }
  | { type: "ask_user"; requestId: string; method: string; question: string; title: string }
  | { type: "task_done"; error?: string }
  | { type: "status"; text: string }
  | {
      type: "usage";
      contextTokens: number;
      contextWindow: number;
      percent: number;
      /** Accumulated session $; null when the model's pricing is unknown. */
      cost: number | null;
    }
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

  async sendPrompt(text: string): Promise<void> {
    log.info(`[session ${this.sessionId}] prompt received (${text.length} chars)`);
    await this.ensureProcess();
    if (!this.proc) {
      log.error(`[session ${this.sessionId}] no pi process after ensureProcess — prompt dropped`);
      return;
    }
    this.resetPromptState();
    this.onEvent({ type: "user_message", text });
    const payload = JSON.stringify({
      id: `prompt-${Date.now()}`,
      type: "prompt",
      message: text,
      streamingBehavior: "followUp",
    });
    log.debug(`[session ${this.sessionId}] -> pi stdin: ${payload.slice(0, 200)}`);
    this.proc.stdin!.write(payload + "\n");
  }

  /** Interrupt the current task; pi settles the turn and goes idle. */
  abort(): void {
    if (!this.running) {
      log.warn(`[session ${this.sessionId}] abort() with no running pi process`);
      return;
    }
    log.info(`[session ${this.sessionId}] user abort`);
    this.userAborted = true;
    this.proc!.stdin!.write(JSON.stringify({ type: "abort" }) + "\n");
    this.onEvent({ type: "status", text: "stopping…" });
  }

  /** Restore accumulated cost when resuming a session from history. */
  seedCost(cost: number): void {
    this.sessionCost = cost;
  }

  /** Answer an ask_user / confirm request raised by an extension. */
  respond(requestId: string, value: string): void {
    log.info(`[session ${this.sessionId}] ui response for ${requestId}: ${value.slice(0, 80)}`);
    if (!this.proc) {
      log.warn(`[session ${this.sessionId}] respond() with no pi process`);
      return;
    }
    this.proc.stdin!.write(
      JSON.stringify({ type: "extension_ui_response", id: requestId, value }) + "\n",
    );
  }

  stop(): void {
    if (this.proc) {
      log.info(`[session ${this.sessionId}] stopping pi (pid ${this.proc.pid})`);
      try {
        this.proc.kill();
      } catch {}
      this.proc = null;
    }
    this.rawLog?.end();
    this.rawLog = null;
  }

  // ── internals ────────────────────────────────────────────────────────
  private starting: Promise<void> | null = null;

  /** Serialize startup so two quick prompts can't spawn two pi processes. */
  private ensureProcess(): Promise<void> {
    if (this.running) {
      log.debug(`[session ${this.sessionId}] pi already running (pid ${this.proc!.pid})`);
      return Promise.resolve();
    }
    if (!this.starting) {
      this.starting = this.startProcess().finally(() => {
        this.starting = null;
      });
    }
    return this.starting;
  }

  private async startProcess(): Promise<void> {
    const cfg = loadLlmConfig();
    if ("error" in cfg) {
      log.error(`[session ${this.sessionId}] config error: ${cfg.error}`);
      this.onEvent({ type: "fatal", text: cfg.error });
      return;
    }
    log.info(
      `[session ${this.sessionId}] config: model=${cfg.model} baseUrl=${cfg.baseUrl} ` +
      `context=${cfg.contextWindow} keySet=${!!cfg.apiKey}`,
    );
    this.contextWindow = cfg.contextWindow;
    const cost = await resolveModelCost(cfg);
    this.costKnown = !!cost;
    if (cost) {
      log.info(
        `[session ${this.sessionId}] pricing: $${cost.input}/M in, $${cost.output}/M out` +
        (cfg.cost ? " (from .env)" : " (from OpenRouter catalog)"),
      );
    } else {
      log.warn(
        `[session ${this.sessionId}] no pricing for ${cfg.model} — cost will show as unknown. ` +
        `Set MODEL_COST_INPUT / MODEL_COST_OUTPUT ($ per million tokens) in agent/.env.`,
      );
    }
    writeModelsJson(cfg, cost);

    const resolved = resolvePiCommand();
    if ("error" in resolved) {
      log.error(`[session ${this.sessionId}] ${resolved.error}`);
      this.onEvent({ type: "fatal", text: resolved.error });
      return;
    }

    const args = [
      ...resolved.argv0,
      "--mode", "rpc",
      "--provider", PROVIDER_NAME,
      "--model", cfg.model,
      "--session-id", this.sessionId,
      "--approve",
      "--extension", path.join(ASSISTANT_ROOT, "extension", "doc-tools.ts"),
      "--system-prompt", SYSTEM_PROMPT,
    ];

    this.rawLog = fs.createWriteStream(path.join(this.userDataDir, "pi-raw.log"), { flags: "a" });
    const logArgs = args.map((a) => (a === SYSTEM_PROMPT ? "<system-prompt>" : a));
    log.info(
      `[session ${this.sessionId}] spawning: ${resolved.cmd} ${logArgs.join(" ")} (cwd=${ASSISTANT_ROOT})`,
    );
    // NOTE: no shell — shell:true does not escape args on Windows and mangles
    // the multi-line system prompt through cmd.exe.
    this.proc = spawn(resolved.cmd, args, {
      cwd: ASSISTANT_ROOT,
      shell: false,
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
    log.info(`[session ${this.sessionId}] pi spawned, pid=${this.proc.pid}`);

    this.proc.on("error", (err) => {
      log.error(`[session ${this.sessionId}] spawn error: ${err.message}`);
      this.onEvent({
        type: "fatal",
        text:
          `Could not launch the 'pi' agent (${err.message}). Install it with: ` +
          `npm install -g --ignore-scripts @earendil-works/pi-coding-agent`,
      });
    });
    this.proc.on("exit", (code, signal) => {
      log.warn(`[session ${this.sessionId}] pi exited code=${code} signal=${signal}`);
      if (code !== 0 && code !== null) {
        this.onEvent({
          type: "fatal",
          text: `Agent process exited (code ${code}). See logs/main.log and pi-raw.log.`,
        });
      }
      this.proc = null;
    });
    this.proc.stderr!.on("data", (chunk: Buffer) => {
      const text = chunk.toString("utf-8");
      this.rawLog?.write(`[stderr] ${text}`);
      log.warn(`[pi stderr] ${text.trim().slice(0, 500)}`);
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

  // ── per-prompt streaming state ───────────────────────────────────────
  private streamedChars = 0;        // deltas emitted for the current message
  private startText = "";           // text seen in message_start (fallback)
  private emittedThisPrompt = false; // any assistant text since last prompt
  private lastError = "";           // last error string seen from pi
  private retries = 0;
  private userAborted = false;      // user pressed stop for the current prompt
  private contextWindow = 0;        // from LLM config, for usage percent
  private sessionCost = 0;          // accumulated $ across the session
  private costKnown = false;        // pricing resolved for the model

  private resetPromptState(): void {
    this.streamedChars = 0;
    this.startText = "";
    this.emittedThisPrompt = false;
    this.lastError = "";
    this.retries = 0;
    this.userAborted = false;
  }

  /**
   * Emit context/cost usage from an assistant message, mirroring pi's own
   * math: context = usage.totalTokens (or the sum of parts); aborted/error
   * messages carry no valid usage.
   */
  private emitUsage(msg: any): void {
    const usage = msg?.usage;
    if (!usage || msg.stopReason === "aborted" || msg.stopReason === "error") return;
    const contextTokens =
      usage.totalTokens ||
      (usage.input || 0) + (usage.output || 0) + (usage.cacheRead || 0) + (usage.cacheWrite || 0);
    this.sessionCost += usage.cost?.total || 0;
    if (contextTokens <= 0) return;
    this.onEvent({
      type: "usage",
      contextTokens,
      contextWindow: this.contextWindow,
      percent: this.contextWindow > 0 ? (contextTokens / this.contextWindow) * 100 : 0,
      cost: this.costKnown ? this.sessionCost : null,
    });
  }

  /** Pull all text parts out of a pi message object. */
  private static extractText(msg: any): string {
    if (!msg || msg.role !== "assistant" || !Array.isArray(msg.content)) return "";
    return msg.content
      .filter((p: any) => p?.type === "text" && p.text)
      .map((p: any) => p.text)
      .join("");
  }

  /** Remember any error-ish field so we can report it at settle time. */
  private captureError(event: any): void {
    const err =
      event.error || event.errorMessage || event.message?.errorMessage ||
      (event.message?.stopReason === "error" ? event.message?.stopReasonMessage : "");
    if (err) {
      this.lastError = String(typeof err === "object" ? JSON.stringify(err) : err);
      log.warn(`[session ${this.sessionId}] pi reported error: ${this.lastError.slice(0, 400)}`);
    }
  }

  /**
   * Event adapter. Supports both the schema open-computer's hypervisor.js
   * documents (message_update deltas, response = done) and the newer agent
   * lifecycle vocabulary observed in the field (agent_start/turn_end,
   * auto_retry_start/end, agent_settled; response = prompt ack).
   */
  private handleLine(line: string): void {
    this.rawLog?.write(line + "\n");
    let event: any;
    try {
      event = JSON.parse(line);
    } catch {
      log.debug(`[pi stdout, non-JSON] ${line.slice(0, 300)}`);
      return; // non-JSON chatter on stdout
    }
    const etype = event.type || event.event;
    log.debug(
      `[session ${this.sessionId}] <- pi event: ${etype}` +
      (event.toolName ? ` tool=${event.toolName}` : ""),
    );
    // Payload-level visibility for the events that matter when diagnosing.
    if (["message_end", "turn_end", "agent_end", "response", "auto_retry_start",
         "agent_settled", "error"].includes(etype)) {
      log.debug(`[session ${this.sessionId}] raw ${etype}: ${line.slice(0, 600)}`);
    }
    this.captureError(event);

    switch (etype) {
      case "agent_start":
      case "turn_start":
        this.onEvent({ type: "status", text: "thinking…" });
        break;

      case "message_start": {
        this.streamedChars = 0;
        this.startText = PiSession.extractText(event.message);
        break;
      }
      case "message_update": {
        const ae = event.assistantMessageEvent;
        if (ae?.type === "text_delta" && ae.delta) {
          this.streamedChars += ae.delta.length;
          this.emittedThisPrompt = true;
          this.onEvent({ type: "assistant_delta", text: ae.delta });
        }
        break;
      }
      case "message_end": {
        // Some pi versions deliver the complete message only here (or only in
        // message_start) with no deltas in between — emit it exactly once.
        const full = PiSession.extractText(event.message) || this.startText;
        if (full && this.streamedChars === 0) {
          this.emittedThisPrompt = true;
          this.onEvent({ type: "assistant_message", text: full });
        }
        this.emitUsage(event.message);
        this.streamedChars = 0;
        this.startText = "";
        break;
      }

      case "tool_execution_start": {
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

      case "auto_retry_start":
        this.retries++;
        this.onEvent({
          type: "status",
          text: `LLM call failed — retrying (attempt ${this.retries})…`,
        });
        break;

      case "agent_settled": {
        // Definitive end-of-prompt in the lifecycle schema.
        let error: string | undefined;
        if (!this.emittedThisPrompt && !this.userAborted) {
          error =
            this.lastError ||
            "The model returned no response. Check that your LLM endpoint is " +
            "reachable (is Ollama running? is the model pulled?) — see logs/main.log.";
        }
        this.onEvent({ type: "task_done", error });
        this.resetPromptState();
        break;
      }

      case "response": {
        // Older schema: response = prompt finished. Newer schema: response is
        // just the RPC ack for the prompt command — only surface failures.
        if (event.success === false && event.error) {
          this.onEvent({ type: "task_done", error: String(event.error) });
          this.resetPromptState();
        }
        break;
      }

      case "streaming_start":
        this.onEvent({ type: "status", text: "thinking…" });
        break;
      default:
        break;
    }
  }
}
