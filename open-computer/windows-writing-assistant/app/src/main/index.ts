import { app, BrowserWindow, globalShortcut, ipcMain } from "electron";
import { randomUUID } from "crypto";
import { History } from "./history";
import { AppEvent, PiSession } from "./pi-session";
import { createChatWindow, createOrbWindow } from "./windows";

let chatWin: BrowserWindow | null = null;
let orbWin: BrowserWindow | null = null;
let session: PiSession | null = null;
let history: History;

function broadcast(event: AppEvent): void {
  history.append(session!.sessionId, event);
  if (chatWin && !chatWin.isDestroyed()) {
    chatWin.webContents.send("assistant:event", event);
  }
}

function startSession(id?: string): void {
  session?.stop();
  session = new PiSession(id || randomUUID().slice(0, 8), app.getPath("userData"), broadcast);
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
  history = new History(app.getPath("userData"));
  startSession();
  chatWin = wireChatWindow();

  globalShortcut.register("Alt+Space", () => {
    if (chatWin && chatWin.isVisible() && chatWin.isFocused()) hideToOrb();
    else showChat();
  });

  // ── IPC ──────────────────────────────────────────────────────────────
  ipcMain.on("chat:send", (_e, text: string) => session?.sendPrompt(text));
  ipcMain.on("chat:respond", (_e, requestId: string, value: string) =>
    session?.respond(requestId, value),
  );
  ipcMain.handle("sessions:list", () => history.list());
  ipcMain.handle("sessions:load", (_e, id: string) => {
    startSession(id); // same id -> pi resumes its own session context
    return history.load(id);
  });
  ipcMain.handle("sessions:new", () => {
    startSession();
    return session!.sessionId;
  });
  ipcMain.on("win:minimize", () => hideToOrb());
  ipcMain.on("win:close", () => app.quit());
  ipcMain.on("orb:restore", () => showChat());
});

app.on("will-quit", () => {
  globalShortcut.unregisterAll();
  session?.stop();
});

app.on("window-all-closed", () => {
  // Orb-only state keeps the app alive; explicit close quits via win:close.
  if (process.platform !== "darwin") app.quit();
});
