/**
 * Session history: one JSONL file of AppEvents per session under
 * <userData>/sessions/, plus an index.json with titles for the sidebar.
 */
import * as fs from "fs";
import * as path from "path";
import type { AppEvent } from "./pi-session";

export interface SessionMeta {
  id: string;
  title: string;
  updated: number;
}

export class History {
  private dir: string;
  private indexPath: string;

  constructor(userDataDir: string) {
    this.dir = path.join(userDataDir, "sessions");
    fs.mkdirSync(this.dir, { recursive: true });
    this.indexPath = path.join(this.dir, "index.json");
  }

  list(): SessionMeta[] {
    try {
      const items: SessionMeta[] = JSON.parse(fs.readFileSync(this.indexPath, "utf-8"));
      return items.sort((a, b) => b.updated - a.updated);
    } catch {
      return [];
    }
  }

  load(id: string): AppEvent[] {
    const file = path.join(this.dir, `${id}.jsonl`);
    if (!fs.existsSync(file)) return [];
    return fs
      .readFileSync(file, "utf-8")
      .split("\n")
      .filter(Boolean)
      .map((line) => {
        try {
          return JSON.parse(line) as AppEvent;
        } catch {
          return null;
        }
      })
      .filter(Boolean) as AppEvent[];
  }

  append(id: string, event: AppEvent): void {
    // Deltas are too granular to replay well; persist consolidated events only.
    if (event.type === "assistant_delta" || event.type === "status") return;
    fs.appendFileSync(path.join(this.dir, `${id}.jsonl`), JSON.stringify(event) + "\n");
    const items = this.list();
    let meta = items.find((s) => s.id === id);
    if (!meta) {
      meta = { id, title: "New chat", updated: Date.now() };
      items.push(meta);
    }
    if (event.type === "user_message" && meta.title === "New chat") {
      meta.title = event.text.slice(0, 60);
    }
    meta.updated = Date.now();
    fs.writeFileSync(this.indexPath, JSON.stringify(items, null, 2));
  }
}
