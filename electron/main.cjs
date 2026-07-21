const { app, BrowserWindow, dialog, ipcMain, Menu, net, shell } = require("electron");
const { autoUpdater } = require("electron-updater");
const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const path = require("node:path");

const UPDATE_REPO = "akilxasp/gridstone-lite";

let mainWindow;
const isMac = process.platform === "darwin";
let pendingFilePath = process.argv.find((argument) => /\.(xlsx?|csv|tsv)$/i.test(argument));
let updatesWired = false;
// On macOS we can't run Squirrel auto-update without an Apple Developer ID
// signature, so mac uses a manual flow: fetch the latest GitHub release, and if
// it's newer, download the .dmg and open it for the user to drag into place.
// This holds the asset chosen by the most recent mac check.
let macPendingAsset = null;

function sendCommand(command, payload) {
  mainWindow?.webContents.send("app-command", command, payload);
}

function sendUpdateStatus(status) {
  sendCommand("update-status", status);
}

// --- Windows / signed builds: electron-updater (Squirrel) ---

// Registers the autoUpdater event listeners once and kicks off the first check.
function setupSquirrelUpdates() {
  if (updatesWired) return;
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

// --- macOS unsigned: manual download-and-open ---

// Returns true when semver string `remote` is strictly higher than `current`.
function isNewerVersion(remote, current) {
  const a = String(remote).split(".").map((part) => parseInt(part, 10) || 0);
  const b = String(current).split(".").map((part) => parseInt(part, 10) || 0);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    const left = a[i] || 0, right = b[i] || 0;
    if (left !== right) return left > right;
  }
  return false;
}

function httpGetJson(url) {
  return new Promise((resolve, reject) => {
    const request = net.request({ method: "GET", url });
    request.setHeader("User-Agent", "Gridstone-Updater");
    request.setHeader("Accept", "application/vnd.github+json");
    request.on("response", (response) => {
      if (response.statusCode !== 200) { reject(new Error(`GitHub responded ${response.statusCode}`)); return; }
      let body = "";
      response.on("data", (chunk) => { body += chunk; });
      response.on("end", () => { try { resolve(JSON.parse(body)); } catch (error) { reject(error); } });
      response.on("error", reject);
    });
    request.on("error", reject);
    request.end();
  });
}

function downloadFile(url, destination, onProgress) {
  return new Promise((resolve, reject) => {
    const request = net.request({ method: "GET", url });
    request.setHeader("User-Agent", "Gridstone-Updater");
    request.on("response", (response) => {
      if ((response.statusCode || 0) >= 400) { reject(new Error(`Download failed (${response.statusCode})`)); return; }
      const header = response.headers["content-length"];
      const total = Number(Array.isArray(header) ? header[0] : header) || 0;
      let received = 0;
      const file = fsSync.createWriteStream(destination);
      response.on("data", (chunk) => { received += chunk.length; file.write(chunk); if (total) onProgress(Math.round((received / total) * 100)); });
      response.on("end", () => file.end(() => resolve()));
      response.on("error", (error) => { file.destroy(); reject(error); });
    });
    request.on("error", reject);
    request.end();
  });
}

async function checkMacUpdate() {
  sendUpdateStatus({ state: "checking" });
  try {
    const release = await httpGetJson(`https://api.github.com/repos/${UPDATE_REPO}/releases/latest`);
    const version = String(release.tag_name || release.name || "").replace(/^v/i, "");
    if (!version || !isNewerVersion(version, app.getVersion())) { sendUpdateStatus({ state: "none" }); return; }
    const dmgs = (release.assets || []).filter((asset) => /\.dmg$/i.test(asset.name));
    const asset = dmgs.find((item) => item.name.includes(process.arch)) || dmgs.find((item) => /universal/i.test(item.name)) || dmgs[0];
    if (!asset) { sendUpdateStatus({ state: "error", message: "Latest release has no .dmg download" }); return; }
    macPendingAsset = { url: asset.browser_download_url, name: asset.name };
    sendUpdateStatus({ state: "available", version, manual: true });
  } catch (error) {
    sendUpdateStatus({ state: "error", message: String(error?.message || error) });
  }
}

async function downloadMacUpdate() {
  if (!macPendingAsset) return;
  const destination = path.join(app.getPath("downloads"), macPendingAsset.name);
  try {
    sendUpdateStatus({ state: "downloading", percent: 0 });
    await downloadFile(macPendingAsset.url, destination, (percent) => sendUpdateStatus({ state: "downloading", percent }));
    sendUpdateStatus({ state: "ready-manual", path: destination });
    await shell.openPath(destination);
  } catch (error) {
    sendUpdateStatus({ state: "error", message: String(error?.message || error) });
  }
}

// --- shared entry points ---

// Kicks off the automatic check on launch for packaged builds.
function setupAutoUpdates() {
  if (!app.isPackaged) return;
  if (isMac) void checkMacUpdate();
  else setupSquirrelUpdates();
}

async function checkForUpdatesManually() {
  if (!app.isPackaged) { sendUpdateStatus({ state: "dev" }); return; }
  if (isMac) { await checkMacUpdate(); return; }
  setupSquirrelUpdates();
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

// macOS manual flow: download the pending .dmg and open it.
ipcMain.handle("download-update", async () => { await downloadMacUpdate(); });

// Windows/Squirrel: quit and relaunch into the freshly downloaded version.
// Only valid once an "update-downloaded" status has been emitted.
ipcMain.handle("install-update", () => { if (app.isPackaged && !isMac) autoUpdater.quitAndInstall(); });
