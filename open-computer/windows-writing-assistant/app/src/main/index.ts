import { app, BrowserWindow, globalShortcut, ipcMain } from "electron";
import { randomUUID } from "crypto";
import { History } from "./history";
import log from "./logger";
import { AppEvent, PiSession } from "./pi-session";
import { createChatWindow, createOrbWindow, supportsAcrylic } from "./windows";

let chatWin: BrowserWindow | null = null;
let orbWin: BrowserWindow | null = null;
let session: PiSession | null = null;
let history: History;

function broadcast(event: AppEvent): void {
  try {
    history.append(session!.sessionId, event);
  } catch (err: any) {
    log.error(`history.append failed: ${err.message}`);
  }
  if (chatWin && !chatWin.isDestroyed()) {
    chatWin.webContents.send("assistant:event", event);
  } else {
    log.warn(`event dropped (no chat window): ${event.type}`);
  }
}

function startSession(id?: string): void {
  session?.stop();
  const sid = id || randomUUID().slice(0, 8);
  log.info(`starting session ${sid}${id ? " (resumed)" : ""}`);
  session = new PiSession(sid, app.getPath("userData"), broadcast);
}

function showChat(): void {
  orbWin?.hide();
  if (!chatWin || chatWin.isDestroyed()) chatWin = wireChatWindow();
  chatWin.show();
  chatWin.focus();
}

function hideToOrb(): void {
  chatWin?.hide();
  if (!orbWin || orbWin.isDestroyed()) orbWin = createOrbWindow();
  orbWin.show();
}

function wireChatWindow(): BrowserWindow {
  const win = createChatWindow();
  // Minimize -> floating orb instead of taskbar.
  win.on("minimize" as any, (e: Electron.Event) => {
    e.preventDefault();
    hideToOrb();
  });
  win.on("close", () => app.quit());
  return win;
}

app.whenReady().then(() => {
  log.info(
    `app ready — version=${app.getVersion()} electron=${process.versions.electron} ` +
    `platform=${process.platform} acrylic=${supportsAcrylic()} userData=${app.getPath("userData")}`,
  );
  history = new History(app.getPath("userData"));
  startSession();
  chatWin = wireChatWindow();

  const registered = globalShortcut.register("Alt+Space", () => {
    if (chatWin && chatWin.isVisible() && chatWin.isFocused()) hideToOrb();
    else showChat();
  });
  log.info(`global shortcut Alt+Space registered=${registered}`);

  // ── IPC ──────────────────────────────────────────────────────────────
  ipcMain.on("chat:send", (_e, text: string) => {
    log.info(`ipc chat:send (${String(text).length} chars)`);
    session?.sendPrompt(text);
  });
  ipcMain.on("chat:respond", (_e, requestId: string, value: string) => {
    log.info(`ipc chat:respond id=${requestId}`);
    session?.respond(requestId, value);
  });
  ipcMain.on("chat:stop", () => {
    log.info("ipc chat:stop");
    session?.abort();
  });
  ipcMain.handle("sessions:list", () => history.list());
  ipcMain.handle("sessions:load", (_e, id: string) => {
    log.info(`ipc sessions:load ${id}`);
    startSession(id); // same id -> pi resumes its own session context
    const events = history.load(id);
    // Carry the session's accumulated cost across the process restart.
    const lastUsage = [...events].reverse().find((e) => e.type === "usage");
    if (lastUsage && lastUsage.type === "usage" && typeof lastUsage.cost === "number") {
      session!.seedCost(lastUsage.cost);
    }
    return events;
  });
  ipcMain.handle("sessions:new", () => {
    startSession();
    return session!.sessionId;
  });
  ipcMain.on("win:minimize", () => hideToOrb());
  ipcMain.on("win:pin", (_e, pinned: boolean) => {
    log.info(`ipc win:pin ${pinned}`);
    chatWin?.setAlwaysOnTop(pinned);
  });
  ipcMain.on("win:close", () => {
    log.info("ipc win:close — quitting");
    app.quit();
  });
  ipcMain.on("orb:restore", () => showChat());
  ipcMain.on("renderer:log", (_e, level: string, message: string) => {
    const fn = (log as any)[level] || log.info;
    fn(`[renderer] ${message}`);
  });
});

process.on("uncaughtException", (err) => {
  log.error(`uncaughtException: ${err.stack || err.message}`);
});
process.on("unhandledRejection", (reason: any) => {
  log.error(`unhandledRejection: ${reason?.stack || reason}`);
});

app.on("will-quit", () => {
  globalShortcut.unregisterAll();
  session?.stop();
});

app.on("window-all-closed", () => {
  // Orb-only state keeps the app alive; explicit close quits via win:close.
  if (process.platform !== "darwin") app.quit();
});
