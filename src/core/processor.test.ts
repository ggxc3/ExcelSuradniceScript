import { describe, expect, it } from "vitest";
import {
  formatCoordinate,
  formatHeight,
  parseColumnsSpec,
  parsePositiveInt,
  processRows
} from "./processor";

describe("parseColumnsSpec", () => {
  it("parses valid specs", () => {
    expect(parseColumnsSpec("a-n, b-e, c-v")).toEqual([
      ["A", "N"],
      ["B", "E"],
      ["C", "V"]
    ]);
  });

  it("throws on invalid type", () => {
    expect(() => parseColumnsSpec("A-X")).toThrow(/Neplatný typ/);
  });
});

describe("numeric parsing", () => {
  it("parses positive ints", () => {
    expect(parsePositiveInt("42", "Od riadku")).toBe(42);
    expect(() => parsePositiveInt("0", "Od riadku")).toThrow(/kladné celé číslo/);
  });
});

describe("formatting", () => {
  it("formats coordinates", () => {
    expect(formatCoordinate("48123", "N")).toBe("48N123");
    expect(formatCoordinate("48N123", "E")).toBe("48E123");
    expect(formatCoordinate("4,123", "N")).toBeNull();
  });

  it("formats heights", () => {
    expect(formatHeight("123.6")).toBe("124");
    expect(formatHeight("123m")).toBe("123");
    expect(formatHeight("abc")).toBe("");
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

    expect(processed).toEqual([
      ["headerA", "headerB", "headerC"],
      ["48N123", "17E123", "120"],
      ["50N123", "20E123", "131"]
    ]);
  });
});
