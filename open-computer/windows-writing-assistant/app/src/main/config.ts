/**
 * LLM endpoint config: reads ../agent/.env (shared with run_agent.py) and
 * writes the provider entry into ~/.pi/agent/models.json — the same schema
 * open-computer generates in interface-service/pi/process.js.
 */
import * as fs from "fs";
import * as os from "os";
import * as path from "path";

export const PROVIDER_NAME = "writing-assistant";

/** windows-writing-assistant/ — app/ lives one level below it. */
export const ASSISTANT_ROOT =
  process.env.DOCD_ROOT || path.resolve(__dirname, "..", "..", "..");

export interface LlmConfig {
  baseUrl: string;
  apiKey: string;
  model: string;
  contextWindow: number;
}

function loadDotenv(file: string): Record<string, string> {
  const out: Record<string, string> = {};
  if (!fs.existsSync(file)) return out;
  for (const line of fs.readFileSync(file, "utf-8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#") || !trimmed.includes("=")) continue;
    const idx = trimmed.indexOf("=");
    const key = trimmed.slice(0, idx).trim();
    const value = trimmed.slice(idx + 1).trim().replace(/^['"]|['"]$/g, "");
    if (key) out[key] = value;
  }
  return out;
}

export function loadLlmConfig(): LlmConfig | { error: string } {
  const envFile = path.join(ASSISTANT_ROOT, "agent", ".env");
  const dotenv = loadDotenv(envFile);
  const get = (k: string) => process.env[k] || dotenv[k] || "";
  const baseUrl = get("OPENAI_BASE_URL");
  const model = get("OPENAI_MODEL");
  if (!baseUrl || !model) {
    return {
      error:
        `Missing LLM config. Set OPENAI_BASE_URL and OPENAI_MODEL in ${envFile} ` +
        `(see agent/.env.example for OpenRouter and Ollama examples).`,
    };
  }
  return {
    baseUrl,
    apiKey: get("OPENAI_API_KEY"),
    model,
    contextWindow: parseInt(get("CONTEXT_WINDOW") || "128000", 10),
  };
}

export function writeModelsJson(cfg: LlmConfig): void {
  const piDir = path.join(os.homedir(), ".pi", "agent");
  fs.mkdirSync(piDir, { recursive: true });
  const modelsPath = path.join(piDir, "models.json");
  let config: any = {};
  try {
    config = JSON.parse(fs.readFileSync(modelsPath, "utf-8"));
  } catch {
    config = {};
  }
  config.providers = config.providers || {};
  config.providers[PROVIDER_NAME] = {
    baseUrl: cfg.baseUrl,
    api: "openai-completions",
    apiKey: cfg.apiKey || "sk-no-key-required",
    compat: {
      supportsDeveloperRole: false,
      supportsReasoningEffort: false,
      supportsUsageInStreaming: false,
      supportsStrictMode: false,
    },
    models: [{ id: cfg.model, contextWindow: cfg.contextWindow }],
  };
  fs.writeFileSync(modelsPath, JSON.stringify(config, null, 2), "utf-8");
}

export const SYSTEM_PROMPT = `You are an AI writing assistant controlling Microsoft Word on the user's own Windows PC. The user watches every edit land live in their real Word window, and can undo anything with Ctrl+Z.

Use the doc_* tools for ALL document work; their descriptions are the source of truth for arguments. Core workflow:
- Creating content: doc_new (blank document), then doc_insert — write normal markdown and it becomes real Word formatting: # headings -> Heading styles, **bold**, *italic*, \`code\`, bullet/numbered lists, [ ]/[x] checkboxes, > quotes. Write the full text yourself; never insert placeholders.
- Editing content: ALWAYS doc_read (or doc_selection, when the user refers to what they selected) first, then edit citing the paragraph hashes you saw (expect_hash / expect_hashes). Prefer doc_replace for small textual changes, doc_edit_range for rewriting passages, doc_tables create for building tables (use replace_range to convert selected text into a table).
- If a tool returns STALE_RANGE, the user changed that text meanwhile: re-read and re-apply, never force.
- Saving: doc_save_as for new files; doc_save overwrites and asks the user first. Do not save unless the user asked for it.

Rules:
- These are the user's real documents. Make exactly the changes asked for — nothing else.
- Never fabricate facts, quotes, or citations in documents; state uncertainty in the chat instead.
- Ask the user (via the chat) when the request is ambiguous about scope, location, or formatting.`;
