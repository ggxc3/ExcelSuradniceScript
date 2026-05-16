type ProcessRequest = {
  inputPath: string;
  outputPath: string;
  sheet: string;
  columns: string;
  startRow: number;
  endRow: number;
};

type SheetInfo = {
  name: string;
  rowCount: number;
};

type Api = {
  openInput: () => Promise<string | null>;
  saveOutput: (defaultPath?: string) => Promise<string | null>;
  listSheets: (inputPath: string) => Promise<SheetInfo[]>;
  processWorkbook: (request: ProcessRequest) => Promise<string>;
};

interface Window {
  api?: Api;
}

const inputEl = document.querySelector<HTMLInputElement>("#inputPath")!;
const outputEl = document.querySelector<HTMLInputElement>("#outputPath")!;
const sheetEl = document.querySelector<HTMLSelectElement>("#sheet")!;
const columnsEl = document.querySelector<HTMLInputElement>("#columns")!;
const startEl = document.querySelector<HTMLInputElement>("#startRow")!;
const endEl = document.querySelector<HTMLInputElement>("#endRow")!;
const statusEl = document.querySelector<HTMLParagraphElement>("#status")!;
const inputButton = document.querySelector<HTMLButtonElement>("#btnInput");
const outputButton = document.querySelector<HTMLButtonElement>("#btnOutput");
const processButton = document.querySelector<HTMLButtonElement>("#btnProcess");
const cancelButton = document.querySelector<HTMLButtonElement>("#btnCancel");
const formEl = document.querySelector<HTMLFormElement>("#processForm")!;
const sheetMetaEl = document.querySelector<HTMLSpanElement>("#sheetMeta")!;
const statusPanel = document.querySelector<HTMLDivElement>("#statusPanel")!;

const api = window.api;

if (!api) {
  setStatus("Interná chyba: nepodarilo sa inicializovať prepojenie medzi aplikáciou a UI (preload).", "error");
  inputButton?.setAttribute("disabled", "true");
  outputButton?.setAttribute("disabled", "true");
  processButton?.setAttribute("disabled", "true");
}

function setStatus(message: string, tone: "idle" | "busy" | "success" | "error" = "idle"): void {
  statusEl.textContent = message;
  statusPanel.dataset.tone = tone;
}

function toUserMessage(err: unknown): string {
  const message = err instanceof Error ? err.message : String(err);
  return message.replace(/^Error invoking remote method '[^']+': Error: /, "");
}

function defaultOutputPath(sourcePath: string): string {
  const match = sourcePath.match(/^(.*?)(\.(xlsx|xlsm|xls))$/i);
  return match ? `${match[1]}_formatted.xlsx` : `${sourcePath}_formatted.xlsx`;
}

async function refreshSheets(inputPath: string): Promise<void> {
  if (!api) return;

  sheetEl.innerHTML = "";
  sheetMetaEl.textContent = "";
  try {
    const sheets = await api.listSheets(inputPath);
    for (const info of sheets) {
      const opt = document.createElement("option");
      opt.value = info.name;
      opt.textContent = info.name;
      opt.dataset.rows = String(info.rowCount);
      sheetEl.appendChild(opt);
    }
    updateSheetMeta();
    if (sheets.length > 0 && (!endEl.value || endEl.value === "1")) {
      endEl.value = String(Math.max(1, sheets[0].rowCount));
    }
    setStatus(sheets.length > 0 ? "Hárky sú načítané. Skontroluj stĺpce a rozsah riadkov." : "Súbor neobsahuje žiadne hárky.");
  } catch (err) {
    setStatus(`Nepodarilo sa načítať hárky: ${toUserMessage(err)}`, "error");
  }
}

function updateSheetMeta(): void {
  const selected = sheetEl.selectedOptions[0];
  if (!selected?.dataset.rows) {
    sheetMetaEl.textContent = "";
    return;
  }

  const rows = Number.parseInt(selected.dataset.rows, 10);
  sheetMetaEl.textContent = Number.isFinite(rows) ? `${rows} riadkov` : "";
}

function parsePositiveRow(input: HTMLInputElement, label: string): number {
  const parsed = Number.parseInt(input.value, 10);
  if (!Number.isSafeInteger(parsed) || parsed <= 0 || String(parsed) !== input.value.trim()) {
    throw new Error(`${label} musí byť kladné celé číslo`);
  }
  return parsed;
}

function setBusy(isBusy: boolean): void {
  inputButton?.toggleAttribute("disabled", isBusy);
  outputButton?.toggleAttribute("disabled", isBusy);
  processButton?.toggleAttribute("disabled", isBusy);
  cancelButton?.toggleAttribute("disabled", isBusy);
  formEl.classList.toggle("is-busy", isBusy);
}

inputButton?.addEventListener("click", async () => {
  if (!api) return;

  const selected = await api.openInput();
  if (!selected) return;

  inputEl.value = selected;
  if (!outputEl.value || outputEl.value.toLowerCase().endsWith("_formatted.xlsx")) {
    outputEl.value = defaultOutputPath(selected);
  }
  await refreshSheets(selected);
});

outputButton?.addEventListener("click", async () => {
  if (!api) return;

  const selected = await api.saveOutput(outputEl.value);
  if (selected) outputEl.value = selected;
});

sheetEl.addEventListener("change", updateSheetMeta);

formEl.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!api) return;

  try {
    const startRow = parsePositiveRow(startEl, "Od riadku");
    const endRow = parsePositiveRow(endEl, "Do riadku");
    const request: ProcessRequest = {
      inputPath: inputEl.value.trim(),
      outputPath: outputEl.value.trim(),
      sheet: sheetEl.value,
      columns: columnsEl.value.trim(),
      startRow,
      endRow
    };

    setBusy(true);
    setStatus("Spracovanie prebieha. Pri väčších súboroch to môže chvíľu trvať...", "busy");
    const result = await api.processWorkbook(request);
    setStatus(result, "success");
  } catch (err) {
    setStatus(toUserMessage(err), "error");
  } finally {
    setBusy(false);
  }
});

cancelButton?.addEventListener("click", () => window.close());
