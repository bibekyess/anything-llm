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

/** USD per million tokens — the units pi's calculateCost expects. */
export interface ModelCost {
  input: number;
  output: number;
  cacheRead: number;
  cacheWrite: number;
}

export interface LlmConfig {
  baseUrl: string;
  apiKey: string;
  model: string;
  contextWindow: number;
  /** Explicit pricing from .env; when absent we try OpenRouter's catalog. */
  cost?: ModelCost;
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
  const costInput = parseFloat(get("MODEL_COST_INPUT"));
  const costOutput = parseFloat(get("MODEL_COST_OUTPUT"));
  return {
    baseUrl,
    apiKey: get("OPENAI_API_KEY"),
    model,
    contextWindow: parseInt(get("CONTEXT_WINDOW") || "128000", 10),
    ...(Number.isFinite(costInput) && Number.isFinite(costOutput)
      ? {
          cost: {
            input: costInput,
            output: costOutput,
            cacheRead: parseFloat(get("MODEL_COST_CACHE_READ")) || 0,
            cacheWrite: parseFloat(get("MODEL_COST_CACHE_WRITE")) || 0,
          },
        }
      : {}),
  };
}

/** model id -> cost (null = looked up, not priced); avoids refetching per session. */
const openRouterCostCache = new Map<string, ModelCost | null>();

/**
 * Resolve $/Mtok pricing for the configured model: explicit .env values win;
 * otherwise, for OpenRouter endpoints, look the model up in their public
 * catalog (per-token prices as decimal strings). Returns null when unknown.
 */
export async function resolveModelCost(cfg: LlmConfig): Promise<ModelCost | null> {
  if (cfg.cost) return cfg.cost;
  if (!/openrouter\.ai/i.test(cfg.baseUrl)) return null;
  if (openRouterCostCache.has(cfg.model)) return openRouterCostCache.get(cfg.model)!;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 5000);
  try {
    const res = await fetch("https://openrouter.ai/api/v1/models", {
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const body: any = await res.json();
    const entry = (body.data || []).find((m: any) => m.id === cfg.model);
    const p = entry?.pricing;
    const cost: ModelCost | null = p
      ? {
          input: (parseFloat(p.prompt) || 0) * 1e6,
          output: (parseFloat(p.completion) || 0) * 1e6,
          cacheRead: (parseFloat(p.input_cache_read) || 0) * 1e6,
          cacheWrite: (parseFloat(p.input_cache_write) || 0) * 1e6,
        }
      : null;
    openRouterCostCache.set(cfg.model, cost);
    return cost;
  } catch {
    return null; // offline or slow catalog — cost stays unknown, don't block the prompt
  } finally {
    clearTimeout(timer);
  }
}

export function writeModelsJson(cfg: LlmConfig, cost: ModelCost | null): void {
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
    models: [
      {
        id: cfg.model,
        contextWindow: cfg.contextWindow,
        // pi prices each message from this block; without it cost is always 0.
        ...(cost ? { cost } : {}),
      },
    ],
  };
  fs.writeFileSync(modelsPath, JSON.stringify(config, null, 2), "utf-8");
}

export const SYSTEM_PROMPT = `You are an AI writing assistant controlling Microsoft Word, PowerPoint, and Hancom Office on the user's own Windows PC. The user watches every edit land live in the real app window, and can undo anything with Ctrl+Z.

Use the doc_*/slide_* tools for ALL document work; their descriptions are the source of truth for arguments.

Text documents (Word .docx, Hancom .hwp/.hwpx):
- Creating content: doc_new (app=word or hwp), then doc_insert — write normal markdown and it becomes real formatting in Word: # headings -> Heading styles, **bold**, *italic*, \`code\`, bullet/numbered lists, [ ]/[x] checkboxes, > quotes. Write the full text yourself; never insert placeholders.
- Editing content: ALWAYS doc_read (or doc_selection, when the user refers to what they selected) first, then edit citing the paragraph hashes you saw (expect_hash / expect_hashes). Prefer doc_replace for small textual changes, doc_edit_range for rewriting passages, doc_tables create for building tables (use replace_range to convert selected text into a table).
- HWP specifics: markdown styling is not applied (plain text); prefer doc_replace (find-anchored) and doc_insert at end or at bookmark (bookmark = named HWP field; list them with hwp_fields). Index-based edits are refused on HWP.
- If a tool returns STALE_RANGE, the user changed that text meanwhile: re-read and re-apply, never force.

Presentations (.pptx — handle from doc_open or doc_new app=powerpoint):
- slide_list first (outline), slide_read for one slide's shapes, then slide_add / slide_edit_text (by shape id + expect_hash) / slide_notes_edit / slide_reorder / slide_delete.
- Slides are visual: after layout-affecting edits, verify with slide_thumbnail.
- Save with pres_save_as (pptx/pdf/png).

Saving: doc_save_as / pres_save_as for new files; doc_save overwrites and asks the user first. Do not save unless the user asked for it.

Rules:
- These are the user's real documents. Make exactly the changes asked for — nothing else.
- Never fabricate facts, quotes, or citations in documents; state uncertainty in the chat instead.
- Ask the user (via the chat) when the request is ambiguous about scope, location, or formatting.`;
