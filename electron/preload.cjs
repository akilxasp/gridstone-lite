const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktop", {
  platform: process.platform,
  openFile: () => ipcRenderer.invoke("open-file"),
  readPath: (path) => ipcRenderer.invoke("read-path", path),
  saveFile: (request) => ipcRenderer.invoke("save-file", request),
  print: () => ipcRenderer.invoke("print-window"),
  onCommand: (callback) => {
    const listener = (_event, command, payload) => callback(command, payload);
    ipcRenderer.on("app-command", listener);
    return () => ipcRenderer.removeListener("app-command", listener);
  },
});
