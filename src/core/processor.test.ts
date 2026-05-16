import { describe, it } from "node:test";
import assert from "node:assert/strict";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import * as XLSX from "@e965/xlsx";
import {
  formatCoordinate,
  formatHeight,
  getWorkbookSheetInfo,
  parseColumnsSpec,
  parsePositiveInt,
  processWorkbook,
  processRows
} from "./processor";

describe("parseColumnsSpec", () => {
  it("parses valid specs", () => {
    assert.deepEqual(parseColumnsSpec("a-n, b-e, c-v"), [
      ["A", "N"],
      ["B", "E"],
      ["C", "V"]
    ]);
  });

  it("throws on invalid type", () => {
    assert.throws(() => parseColumnsSpec("A-X"), /Neplatný typ/);
  });
});

describe("numeric parsing", () => {
  it("parses positive ints", () => {
    assert.equal(parsePositiveInt("42", "Od riadku"), 42);
    assert.throws(() => parsePositiveInt("0", "Od riadku"), /kladné celé číslo/);
  });
});

describe("formatting", () => {
  it("formats coordinates", () => {
    assert.equal(formatCoordinate("48123", "N"), "48N123");
    assert.equal(formatCoordinate("48N123", "E"), "48E123");
    assert.equal(formatCoordinate("4,123", "N"), null);
  });

  it("formats heights", () => {
    assert.equal(formatHeight("123.6"), "124");
    assert.equal(formatHeight("123m"), "123");
    assert.equal(formatHeight("abc"), "");
  });
});

describe("processRows", () => {
  it("deletes invalid coordinate rows and formats height", () => {
    const rows = [
      ["headerA", "headerB", "headerC"],
      ["48123", "17123", "120.4"],
      ["bad", "19123", "130"],
      ["50123", "20123", "131m"]
    ];

    const processed = processRows(rows, [["A", "N"], ["B", "E"], ["C", "V"]], 2, 4);

    assert.deepEqual(processed, [
      ["headerA", "headerB", "headerC"],
      ["48N123", "17E123", "120"],
      ["50N123", "20E123", "131"]
    ]);
  });
});

describe("processWorkbook", () => {
  it("rejects invalid row numbers", () => {
    assert.throws(
      () =>
      processWorkbook({
        inputPath: "input.xlsx",
        outputPath: "output.xlsx",
        sheet: "Sheet1",
        columns: "A-N",
        startRow: Number.NaN,
        endRow: 10
      }),
      /kladné celé čísla/
    );
  });

  it("updates the selected sheet and preserves the other sheets", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-suradnice-"));
    const inputPath = path.join(dir, "input.xlsx");
    const outputPath = path.join(dir, "output.xlsx");

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["coord"],
        ["48123"],
        ["bad"]
      ]),
      "Data"
    );
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["keep"], ["unchanged"]]), "Notes");
    XLSX.writeFile(workbook, inputPath);

    assert.deepEqual(getWorkbookSheetInfo(inputPath), [
      { name: "Data", rowCount: 3 },
      { name: "Notes", rowCount: 2 }
    ]);

    processWorkbook({
      inputPath,
      outputPath,
      sheet: "Data",
      columns: "A-N",
      startRow: 2,
      endRow: 3
    });

    const output = XLSX.readFile(outputPath, { cellDates: false });
    assert.deepEqual(output.SheetNames, ["Data", "Notes"]);
    assert.deepEqual(XLSX.utils.sheet_to_json(output.Sheets.Data, { header: 1 }), [["coord"], ["48N123"]]);
    assert.deepEqual(XLSX.utils.sheet_to_json(output.Sheets.Notes, { header: 1 }), [["keep"], ["unchanged"]]);

    fs.rmSync(dir, { recursive: true, force: true });
  });
});
