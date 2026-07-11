"""Launch the AI writing assistant: the pi agent + doc-tools extension,
against ANY OpenAI-compatible endpoint (OpenRouter, Ollama, LM Studio, ...).

    cd open-computer\\windows-writing-assistant
    pip install pywin32
    npm install -g --ignore-scripts @earendil-works/pi-coding-agent
    copy agent\\.env.example agent\\.env   (then edit: endpoint + key + model)
    python agent\\run_agent.py

This mirrors how open-computer wires pi to a provider
(services/interface-service/pi/process.js): a provider entry in
~/.pi/agent/models.json with api "openai-completions", then
`pi --provider <name> --model <id> --extension doc-tools.ts`.
pi hosts the agent loop; the LLM generates text/decisions as doc_* tool
calls; doc-tools.ts forwards them to the docd sidecar; docd drives Word.

Use --dry-run to print the resolved config and command without launching.
"""

import argparse
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)  # windows-writing-assistant/
EXTENSION = os.path.join(ROOT, "extension", "doc-tools.ts")
PROVIDER_NAME = "writing-assistant"

SYSTEM_PROMPT = """You are an AI writing assistant controlling Microsoft Word on the user's own Windows PC. The user watches every edit land live in their real Word window, and can undo anything with Ctrl+Z.

Use the doc_* tools for ALL document work; their descriptions are the source of truth for arguments. Core workflow:
- Creating content: doc_new (blank document), then doc_insert — write normal markdown and it becomes real Word formatting: # headings -> Heading styles, **bold**, *italic*, `code`, bullet/numbered lists, [ ]/[x] checkboxes, > quotes. Write the full text yourself; never insert placeholders.
- Editing content: ALWAYS doc_read (or doc_selection, when the user refers to what they selected) first, then edit citing the paragraph hashes you saw (expect_hash / expect_hashes). Prefer doc_replace for small textual changes, doc_edit_range for rewriting passages, doc_tables create for building tables (use replace_range to convert selected text into a table).
- If a tool returns STALE_RANGE, the user changed that text meanwhile: re-read and re-apply, never force.
- Saving: doc_save_as for new files; doc_save overwrites and asks the user first. Do not save unless the user asked for it.

Rules:
- These are the user's real documents. Make exactly the changes asked for — nothing else.
- Never fabricate facts, quotes, or citations in documents; state uncertainty in the chat instead.
- Ask the user (via the chat) when the request is ambiguous about scope, location, or formatting."""


def load_dotenv(path):
    """Tiny .env loader: KEY=VALUE lines, no expansion; env wins over file."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip().strip("'\"")
            if key and key not in os.environ:
                os.environ[key] = value


def write_models_json(base_url, api_key, model, context_window):
    """Merge our provider into ~/.pi/agent/models.json (same schema as
    open-computer's writeModelsJson — other providers are left untouched)."""
    pi_dir = os.path.join(os.path.expanduser("~"), ".pi", "agent")
    os.makedirs(pi_dir, exist_ok=True)
    models_path = os.path.join(pi_dir, "models.json")
    config = {}
    if os.path.exists(models_path):
        try:
            with open(models_path, encoding="utf-8") as f:
                config = json.load(f)
        except (json.JSONDecodeError, OSError):
            config = {}
    config.setdefault("providers", {})[PROVIDER_NAME] = {
        "baseUrl": base_url,
        "api": "openai-completions",
        "apiKey": api_key or "sk-no-key-required",
        "compat": {
            "supportsDeveloperRole": False,
            "supportsReasoningEffort": False,
            "supportsUsageInStreaming": False,
            "supportsStrictMode": False,
        },
        "models": [{"id": model, "contextWindow": context_window}],
    }
    with open(models_path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)
    return models_path


def main():
    parser = argparse.ArgumentParser(prog="run_agent")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print resolved config and command; don't launch.")
    parser.add_argument("prompt", nargs="*",
                        help="Optional one-shot prompt (default: interactive chat).")
    args = parser.parse_args()

    load_dotenv(os.path.join(HERE, ".env"))
    base_url = os.environ.get("OPENAI_BASE_URL", "")
    api_key = os.environ.get("OPENAI_API_KEY", "")
    model = os.environ.get("OPENAI_MODEL", "")
    context_window = int(os.environ.get("CONTEXT_WINDOW", "128000"))

    if not base_url or not model:
        sys.exit(
            "Missing config. Set OPENAI_BASE_URL and OPENAI_MODEL "
            "(plus OPENAI_API_KEY for cloud providers) in agent/.env — "
            "see agent/.env.example for OpenRouter and Ollama examples."
        )

    pi_cmd = shutil.which("pi")
    if not pi_cmd and not args.dry_run:
        sys.exit(
            "The 'pi' agent CLI is not on PATH. Install it with:\n"
            "  npm install -g --ignore-scripts @earendil-works/pi-coding-agent"
        )

    models_path = write_models_json(base_url, api_key, model, context_window)

    cmd = [
        pi_cmd or "pi",
        "--provider", PROVIDER_NAME,
        "--model", model,
        "--extension", EXTENSION,
        "--system-prompt", SYSTEM_PROMPT,
    ]
    if args.prompt:
        cmd += ["-p", " ".join(args.prompt)]

    env = {
        **os.environ,
        # doc-tools.ts spawns the sidecar with these:
        "DOCD_PYTHON": sys.executable,
        "DOCD_CWD": ROOT,
        # Keep the whole pipeline UTF-8 on Windows (Korean/CJK safety).
        "PYTHONUTF8": "1",
        "PYTHONIOENCODING": "utf-8",
    }
    if api_key:
        env["OPENAI_API_KEY"] = api_key

    print(f"[run_agent] provider={PROVIDER_NAME}  model={model}")
    print(f"[run_agent] endpoint={base_url}")
    print(f"[run_agent] models.json={models_path}")
    print(f"[run_agent] extension={EXTENSION}")
    if args.dry_run:
        print(f"[run_agent] would run: {subprocess.list2cmdline(cmd[:7])} --system-prompt <...>")
        return
    print("[run_agent] Starting pi — chat with your documents. Ctrl+C to exit.\n")
    sys.exit(subprocess.call(cmd, env=env, cwd=ROOT))


if __name__ == "__main__":
    main()
