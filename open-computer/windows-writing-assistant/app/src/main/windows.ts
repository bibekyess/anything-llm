/**
 * Window management: the glass chat window (Win11 acrylic) and the floating
 * orb shown while the chat is hidden.
 */
import { BrowserWindow, screen, shell } from "electron";
import * as os from "os";
import * as path from "path";

const isDev = !!process.env["ELECTRON_RENDERER_URL"];

/** Windows 11 = NT 10.0 build 22000+; acrylic backgroundMaterial needs it. */
export function supportsAcrylic(): boolean {
  if (process.platform !== "win32") return false;
  const build = parseInt(os.release().split(".")[2] || "0", 10);
  return build >= 22000;
}

function loadView(win: BrowserWindow, view: string): void {
  if (isDev) {
    win.loadURL(`${process.env["ELECTRON_RENDERER_URL"]}?view=${view}`);
  } else {
    win.loadFile(path.join(__dirname, "../renderer/index.html"), {
      query: { view },
    });
  }
}

export function createChatWindow(): BrowserWindow {
  // Dock as a right-side panel next to the document the user is working on.
  const { workArea } = screen.getPrimaryDisplay();
  const width = 460;
  const height = Math.min(920, workArea.height - 40);
  const win = new BrowserWindow({
    width,
    height,
    x: workArea.x + workArea.width - width - 16,
    y: workArea.y + Math.round((workArea.height - height) / 2),
    minWidth: 380,
    minHeight: 480,
    frame: false,
    show: false,
    autoHideMenuBar: true,
    // Float over Word so the assistant stays visible while the user types
    // there; the titlebar pin button toggles this off.
    alwaysOnTop: true,
    // Truly transparent window: Word stays visible through the panel at all
    // times. (Acrylic was tried first, but Windows suspends its blur whenever
    // the window is unfocused — i.e. almost always for a side-assistant — and
    // it fell back to flat grey.) The smoked-glass look is drawn in CSS.
    transparent: true,
    backgroundColor: "#00000000",
    hasShadow: false, // shadow is drawn in CSS (native shadow breaks on transparent windows)
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.once("ready-to-show", () => win.show());
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });
  loadView(win, "chat");
  return win;
}

export function createOrbWindow(): BrowserWindow {
  const { workArea } = screen.getPrimaryDisplay();
  const size = 64;
  const win = new BrowserWindow({
    width: size,
    height: size,
    x: workArea.x + workArea.width - size - 24,
    y: workArea.y + workArea.height - size - 24,
    frame: false,
    transparent: true,
    resizable: false,
    alwaysOnTop: true,
    skipTaskbar: true,
    show: false,
    hasShadow: false,
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  loadView(win, "orb");
  return win;
}
