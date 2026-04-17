type ProcessRequest = {
  inputPath: string;
  outputPath: string;
  sheet: string;
  columns: string;
  startRow: number;
  endRow: number;
};

type Api = {
  openInput: () => Promise<string | null>;
  saveOutput: (defaultPath?: string) => Promise<string | null>;
  listSheets: (inputPath: string) => Promise<string[]>;
  processWorkbook: (request: ProcessRequest) => Promise<string>;
};

declare global {
  interface Window {
    api: Api;
  }
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

if (!window.api) {
  statusEl.textContent =
    "Interná chyba: nepodarilo sa inicializovať prepojenie medzi aplikáciou a UI (preload).";
  inputButton?.setAttribute("disabled", "true");
  outputButton?.setAttribute("disabled", "true");
  processButton?.setAttribute("disabled", "true");
}

function defaultOutputPath(sourcePath: string): string {
  if (sourcePath.toLowerCase().endsWith(".xlsx")) {
    return `${sourcePath.slice(0, -5)}_formatted.xlsx`;
  }
  return `${sourcePath}_formatted.xlsx`;
}

async function refreshSheets(inputPath: string): Promise<void> {
  sheetEl.innerHTML = "";
  try {
    const sheets = await window.api.listSheets(inputPath);
    for (const name of sheets) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetEl.appendChild(opt);
    }
    statusEl.textContent = sheets.length > 0 ? "Načítané hárky zo vstupného súboru." : "Súbor neobsahuje žiadne hárky.";
  } catch (err) {
    statusEl.textContent = `Nepodarilo sa načítať hárky: ${(err as Error).message}`;
  }
}

inputButton?.addEventListener("click", async () => {
  const selected = await window.api.openInput();
  if (!selected) return;

  inputEl.value = selected;
  if (!outputEl.value || outputEl.value.toLowerCase().endsWith("_formatted.xlsx")) {
    outputEl.value = defaultOutputPath(selected);
  }
  await refreshSheets(selected);
});

outputButton?.addEventListener("click", async () => {
  const selected = await window.api.saveOutput(outputEl.value);
  if (selected) outputEl.value = selected;
});

processButton?.addEventListener("click", async () => {
  const request: ProcessRequest = {
    inputPath: inputEl.value,
    outputPath: outputEl.value,
    sheet: sheetEl.value,
    columns: columnsEl.value,
    startRow: Number.parseInt(startEl.value, 10),
    endRow: Number.parseInt(endEl.value, 10)
  };

  statusEl.textContent = "Spracovanie prebieha...";
  try {
    const result = await window.api.processWorkbook(request);
    statusEl.textContent = result;
  } catch (err) {
    statusEl.textContent = (err as Error).message;
  }
});

cancelButton?.addEventListener("click", () => window.close());


export {};
