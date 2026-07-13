import { contextBridge, ipcRenderer } from "electron";

export interface AssistantApi {
  view: string;
  send(text: string): void;
  respond(requestId: string, value: string): void;
  onEvent(cb: (event: any) => void): () => void;
  sessions: {
    list(): Promise<Array<{ id: string; title: string; updated: number }>>;
    load(id: string): Promise<any[]>;
    create(): Promise<string>;
  };
  win: { minimize(): void; close(): void; restore(): void };
}

const view = new URLSearchParams(location.search).get("view") || "chat";

const api: AssistantApi = {
  view,
  send: (text) => ipcRenderer.send("chat:send", text),
  respond: (requestId, value) => ipcRenderer.send("chat:respond", requestId, value),
  onEvent: (cb) => {
    const listener = (_e: unknown, event: any) => cb(event);
    ipcRenderer.on("assistant:event", listener);
    return () => ipcRenderer.removeListener("assistant:event", listener);
  },
  sessions: {
    list: () => ipcRenderer.invoke("sessions:list"),
    load: (id) => ipcRenderer.invoke("sessions:load", id),
    create: () => ipcRenderer.invoke("sessions:new"),
  },
  win: {
    minimize: () => ipcRenderer.send("win:minimize"),
    close: () => ipcRenderer.send("win:close"),
    restore: () => ipcRenderer.send("orb:restore"),
  },
};

contextBridge.exposeInMainWorld("assistant", api);
