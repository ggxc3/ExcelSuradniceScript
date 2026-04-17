import * as XLSX from "xlsx";
import { ProcessRequest } from "./types";

export type ColumnSpec = [column: string, coordinateType: "N" | "E" | "V"];

export function parseColumnsSpec(spec: string): ColumnSpec[] {
  const normalized = spec.toUpperCase().replaceAll(" ", "");
  if (!normalized) {
    throw new Error("Pole Stĺpce je povinné (napr. A-N,B-E,C-V)");
  }

  return normalized.split(",").map((part) => {
    const split = part.split("-");
    if (split.length !== 2 || !split[0] || !split[1]) {
      throw new Error(`Neplatný formát stĺpcov: ${part}`);
    }
    if (split[1] !== "N" && split[1] !== "E" && split[1] !== "V") {
      throw new Error(`Neplatný typ ${split[1]} v ${part} (povolené: N, E, V)`);
    }
    return [split[0], split[1] as "N" | "E" | "V"];
  });
}

export function parsePositiveInt(value: string, fieldName: string): number {
  const parsed = Number.parseInt(value.trim(), 10);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error(`${fieldName} musí byť kladné celé číslo`);
  }
  return parsed;
}

function isDigit(char: string): boolean {
  return /^\d$/.test(char);
}

function startsWithTwoDigits(value: string): boolean {
  return value.length >= 2 && isDigit(value[0]) && isDigit(value[1]);
}

export function formatCoordinate(value: string, coordinateType: "N" | "E"): string | null {
  if (!value || value.length < 5 || value[1] === "." || value[1] === ",") {
    return null;
  }

  let text = value.replaceAll(" ", "").replaceAll(",", ".");
  if (text.startsWith("0")) {
    text = text.slice(1);
  }
  const first = text[0];
  const second = text[1];

  if (!(startsWithTwoDigits(text) || ((first === "E" || first === "N") && isDigit(second)))) {
    return null;
  }

  if (coordinateType === "E") {
    text = text.replaceAll("N", "");
    if (!text.includes("E")) {
      text = `${text.slice(0, 2)}E${text.slice(2)}`;
    }
  } else {
    text = text.replaceAll("E", "");
    if (!text.includes("N")) {
      text = `${text.slice(0, 2)}N${text.slice(2)}`;
    }
  }

  return text;
}

export function formatHeight(value: string): string {
  if (!value) return "";

  const normalized = value.replaceAll(" ", "").replaceAll(".", ",");
  const parseTarget = normalized.replaceAll(",", ".");

  const floatValue = Number.parseFloat(parseTarget);
  if (!Number.isNaN(floatValue) && /^\d+([.,]\d+)?$/.test(normalized)) {
    return Math.round(floatValue).toString();
  }

  let digits = "";
  for (const char of normalized) {
    if (isDigit(char)) {
      digits += char;
    } else {
      break;
    }
  }
  return digits;
}

export function columnToIndex(column: string): number {
  return column
    .toUpperCase()
    .split("")
    .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0);
}

export function processRows(rows: string[][], columns: ColumnSpec[], startRow: number, endRow: number): string[][] {
  let effectiveEnd = Math.min(endRow, rows.length);

  for (const [column, type] of columns) {
    const colIndex = columnToIndex(column) - 1;

    if (type === "V") {
      for (let row = startRow; row <= effectiveEnd; row += 1) {
        if (!rows[row - 1]) continue;
        rows[row - 1][colIndex] = formatHeight(rows[row - 1][colIndex] ?? "");
      }
      continue;
    }

    const deleteRows: number[] = [];
    for (let row = startRow; row <= effectiveEnd; row += 1) {
      const source = rows[row - 1]?.[colIndex] ?? "";
      const formatted = formatCoordinate(source, type);
      if (formatted === null) {
        deleteRows.push(row);
      } else {
        rows[row - 1][colIndex] = formatted;
      }
    }

    for (let i = deleteRows.length - 1; i >= 0; i -= 1) {
      rows.splice(deleteRows[i] - 1, 1);
    }
    effectiveEnd = Math.max(startRow - 1, effectiveEnd - deleteRows.length);
  }

  return rows;
}

export function processWorkbook(request: ProcessRequest): string {
  if (!request.inputPath.trim()) throw new Error("Vyber vstupný Excel súbor");
  if (!request.outputPath.trim()) throw new Error("Zvoľ cieľový súbor pre uloženie");
  if (!request.sheet.trim()) throw new Error("Vyber hárok");
  if (request.startRow <= 0 || request.endRow <= 0) {
    throw new Error("Riadky musia byť kladné celé čísla");
  }
  if (request.startRow > request.endRow) {
    throw new Error("Počiatočný riadok nesmie byť väčší ako koncový");
  }

  const columns = parseColumnsSpec(request.columns);
  const workbook = XLSX.readFile(request.inputPath, { cellDates: false });
  const sheet = workbook.Sheets[request.sheet];
  if (!sheet) {
    throw new Error(`Hárok ${request.sheet} sa v súbore nenašiel`);
  }

  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, {
    header: 1,
    raw: false,
    defval: ""
  });

  const processedRows = processRows(rows, columns, request.startRow, request.endRow);
  const outputSheet = XLSX.utils.aoa_to_sheet(processedRows);

  const outputWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWb, outputSheet, request.sheet);
  XLSX.writeFile(outputWb, request.outputPath);

  return `Hotovo. Súbor bol úspešne uložený: ${request.outputPath}`;
}
