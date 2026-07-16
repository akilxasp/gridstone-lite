import { describe, expect, it } from "vitest";
import { addressOf, columnName, createWorkbook, evaluateCell, exportWorkbook, importWorkbook, pointFromAddress, setCellInput } from "./workbook";

describe("cell addressing", () => {
  it("supports columns beyond Z", () => {
    expect(columnName(0)).toBe("A");
    expect(columnName(25)).toBe("Z");
    expect(columnName(26)).toBe("AA");
    expect(addressOf(9, 27)).toBe("AB10");
    expect(pointFromAddress("$AB$10")).toEqual({ row: 9, col: 27 });
  });
});

describe("formula engine", () => {
  it("calculates arithmetic, references and ranges", () => {
    const workbook = createWorkbook();
    let sheet = workbook.sheets[0];
    sheet = setCellInput(sheet, "A1", "10");
    sheet = setCellInput(sheet, "A2", "20");
    sheet = setCellInput(sheet, "A3", "=SUM(A1:A2) * 2");
    workbook.sheets[0] = sheet;
    expect(evaluateCell(workbook, sheet, "A3")).toBe(60);
  });

  it("detects circular references", () => {
    const workbook = createWorkbook();
    let sheet = workbook.sheets[0];
    sheet = setCellInput(sheet, "A1", "=B1");
    sheet = setCellInput(sheet, "B1", "=A1");
    workbook.sheets[0] = sheet;
    expect(evaluateCell(workbook, sheet, "A1")).toBe("#CIRC!");
  });
});

describe("file compatibility", () => {
  it("round trips xlsx workbooks", () => {
    const workbook = createWorkbook("Round trip");
    workbook.sheets[0] = setCellInput(workbook.sheets[0], "C4", "42");
    const bytes = exportWorkbook(workbook, "xlsx");
    const restored = importWorkbook(bytes, "Round trip.xlsx");
    expect(restored.sheets[0].cells.C4.raw).toBe(42);
  });
});
