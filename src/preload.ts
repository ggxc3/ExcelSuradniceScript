import { contextBridge, ipcRenderer } from "electron";
import type { ProcessRequest, SheetInfo } from "./core/types";

contextBridge.exposeInMainWorld("api", {
  openInput: () => ipcRenderer.invoke("dialog:open-input") as Promise<string | null>,
  saveOutput: (defaultPath?: string) =>
    ipcRenderer.invoke("dialog:save-output", defaultPath) as Promise<string | null>,
  listSheets: (inputPath: string) =>
    ipcRenderer.invoke("workbook:list-sheets", inputPath) as Promise<SheetInfo[]>,
  processWorkbook: (request: ProcessRequest) =>
    ipcRenderer.invoke("workbook:process", request) as Promise<string>
});
