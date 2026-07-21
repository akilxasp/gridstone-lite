const { app, BrowserWindow, dialog, ipcMain, Menu, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs/promises");
const path = require("node:path");

let mainWindow;
const isMac = process.platform === "darwin";
let pendingFilePath = process.argv.find((argument) => /\.(xlsx?|csv|tsv)$/i.test(argument));
let updatesWired = false;

function sendCommand(command, payload) {
  mainWindow?.webContents.send("app-command", command, payload);
}

function sendUpdateStatus(status) {
  sendCommand("update-status", status);
}

// Registers the autoUpdater event listeners once and kicks off the first check.
// Only runs for packaged builds — in dev there is nothing to update.
function setupAutoUpdates() {
  if (updatesWired || !app.isPackaged) return;
  updatesWired = true;

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on("checking-for-update", () => sendUpdateStatus({ state: "checking" }));
  autoUpdater.on("update-available", (info) => sendUpdateStatus({ state: "available", version: info?.version }));
  autoUpdater.on("update-not-available", () => sendUpdateStatus({ state: "none" }));
  autoUpdater.on("download-progress", (progress) => sendUpdateStatus({ state: "downloading", percent: Math.round(progress?.percent || 0) }));
  autoUpdater.on("update-downloaded", (info) => sendUpdateStatus({ state: "downloaded", version: info?.version }));
  autoUpdater.on("error", (error) => sendUpdateStatus({ state: "error", message: String(error?.message || error) }));

  autoUpdater.checkForUpdates().catch((error) => sendUpdateStatus({ state: "error", message: String(error?.message || error) }));
}

async function checkForUpdatesManually() {
  if (!app.isPackaged) { sendUpdateStatus({ state: "dev" }); return; }
  setupAutoUpdates();
  try { await autoUpdater.checkForUpdates(); }
  catch (error) { sendUpdateStatus({ state: "error", message: String(error?.message || error) }); }
}

function createMenu() {
  const template = [
    ...(isMac ? [{ label: app.name, submenu: [{ role: "about" }, { label: "Check for Updates…", click: () => void checkForUpdatesManually() }, { type: "separator" }, { role: "services" }, { type: "separator" }, { role: "hide" }, { role: "hideOthers" }, { role: "unhide" }, { type: "separator" }, { role: "quit" }] }] : []),
    { label: "File", submenu: [
      { label: "New Workbook", accelerator: "CmdOrCtrl+N", click: () => sendCommand("new") },
      { label: "Open…", accelerator: "CmdOrCtrl+O", click: () => sendCommand("open") },
      { type: "separator" },
      { label: "Save", accelerator: "CmdOrCtrl+S", click: () => sendCommand("save") },
      { label: "Save As…", accelerator: "CmdOrCtrl+Shift+S", click: () => sendCommand("save-as") },
      { type: "separator" },
      { label: "Print…", accelerator: "CmdOrCtrl+P", click: () => sendCommand("print") },
      ...(isMac ? [] : [{ type: "separator" }, { role: "quit" }]),
    ]},
    { label: "Edit", submenu: [{ role: "undo" }, { role: "redo" }, { type: "separator" }, { role: "cut" }, { role: "copy" }, { role: "paste" }, { role: "selectAll" }, { type: "separator" }, { label: "Find", accelerator: "CmdOrCtrl+F", click: () => sendCommand("find") }] },
    { label: "View", submenu: [{ role: "reload" }, { role: "toggleDevTools" }, { type: "separator" }, { role: "resetZoom" }, { role: "zoomIn" }, { role: "zoomOut" }, { type: "separator" }, { role: "togglefullscreen" }] },
    { label: "Window", submenu: [{ role: "minimize" }, { role: "zoom" }, ...(isMac ? [{ type: "separator" }, { role: "front" }] : [{ role: "close" }])] },
    { label: "Help", submenu: [{ label: "Gridstone documentation", click: () => shell.openExternal("https://github.com/") }, ...(isMac ? [] : [{ label: "Check for Updates…", click: () => void checkForUpdatesManually() }])] },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 920,
    minWidth: 980,
    minHeight: 640,
    titleBarStyle: isMac ? "hiddenInset" : "default",
    backgroundColor: "#f4f5f0",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  if (!app.isPackaged) mainWindow.loadURL("http://127.0.0.1:5173");
  else mainWindow.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  mainWindow.webContents.once("did-finish-load", () => {
    if (pendingFilePath) { sendCommand("open-path", pendingFilePath); pendingFilePath = undefined; }
    setupAutoUpdates();
  });
}

app.on("open-file", (event, filePath) => {
  event.preventDefault();
  if (mainWindow?.webContents && !mainWindow.webContents.isLoading()) sendCommand("open-path", filePath);
  else pendingFilePath = filePath;
});

app.whenReady().then(() => {
  createMenu();
  createWindow();
  app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on("window-all-closed", () => { if (!isMac) app.quit(); });

ipcMain.handle("open-file", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [
      { name: "Spreadsheets", extensions: ["xlsx", "xls", "csv", "tsv"] },
      { name: "Excel workbooks", extensions: ["xlsx", "xls"] },
      { name: "Delimited text", extensions: ["csv", "tsv"] },
    ],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const filePath = result.filePaths[0];
  const data = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), extension: path.extname(filePath).slice(1).toLowerCase(), data };
});

ipcMain.handle("read-path", async (_event, filePath) => {
  const data = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), extension: path.extname(filePath).slice(1).toLowerCase(), data };
});

ipcMain.handle("save-file", async (_event, request) => {
  let filePath = request.path;
  if (!filePath || request.saveAs) {
    const extension = request.extension || "xlsx";
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: request.suggestedName || `Untitled.${extension}`,
      filters: extension === "csv"
        ? [{ name: "CSV file", extensions: ["csv"] }]
        : [{ name: "Excel workbook", extensions: ["xlsx"] }, { name: "CSV file", extensions: ["csv"] }],
    });
    if (result.canceled || !result.filePath) return null;
    filePath = result.filePath;
  }
  await fs.writeFile(filePath, Buffer.from(request.data));
  return { path: filePath, name: path.basename(filePath), extension: path.extname(filePath).slice(1).toLowerCase() };
});

ipcMain.handle("print-window", async () => {
  return new Promise((resolve) => mainWindow.webContents.print({ printBackground: true }, (success, failureReason) => resolve({ success, failureReason })));
});

ipcMain.handle("check-for-updates", async () => { await checkForUpdatesManually(); });

// Quits and relaunches into the freshly downloaded version. Only valid once an
// "update-downloaded" status has been emitted.
ipcMain.handle("install-update", () => { if (app.isPackaged) autoUpdater.quitAndInstall(); });
