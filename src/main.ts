import { app, BrowserWindow, dialog, ipcMain } from "electron";
import path from "node:path";
import { processWorkbook } from "./core/processor";
import type { ProcessRequest } from "./core/types";

function createWindow(): void {
  const window = new BrowserWindow({
    width: 760,
    height: 560,
    backgroundColor: "#f5f8fc",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  window.loadFile(path.join(__dirname, "renderer", "index.html"));
}

ipcMain.handle("dialog:open-input", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel files", extensions: ["xlsx", "xlsm", "xls"] }]
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle("dialog:save-output", async (_event, defaultPath?: string) => {
  const result = await dialog.showSaveDialog({
    defaultPath,
    filters: [{ name: "Excel files", extensions: ["xlsx"] }]
  });
  return result.canceled ? null : result.filePath;
});

ipcMain.handle("workbook:list-sheets", (_event, inputPath: string) => {
  const XLSX = require("xlsx") as typeof import("xlsx");
  const wb = XLSX.readFile(inputPath, { bookSheets: true });
  return wb.SheetNames;
});

ipcMain.handle("workbook:process", (_event, request: ProcessRequest) => {
  return processWorkbook(request);
});

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
