import { app, BrowserWindow, dialog, ipcMain } from "electron";
import path from "node:path";
import { getWorkbookSheetInfo, processWorkbook } from "./core/processor";
import type { ProcessRequest } from "./core/types";

function createWindow(): void {
  const window = new BrowserWindow({
    width: 980,
    height: 720,
    minWidth: 820,
    minHeight: 620,
    backgroundColor: "#f2f4f1",
    title: "Excel Súradnice Script",
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
    title: "Vyber Excel súbor",
    properties: ["openFile"],
    filters: [{ name: "Excel files", extensions: ["xlsx", "xlsm", "xls"] }]
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle("dialog:save-output", async (_event, defaultPath?: string) => {
  const result = await dialog.showSaveDialog({
    title: "Uložiť spracovaný Excel",
    defaultPath,
    filters: [{ name: "Excel files", extensions: ["xlsx", "xlsm", "xls"] }]
  });
  return result.canceled ? null : result.filePath;
});

ipcMain.handle("workbook:list-sheets", (_event, inputPath: string) => {
  return getWorkbookSheetInfo(inputPath);
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
