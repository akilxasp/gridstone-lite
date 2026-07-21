import * as XLSX from "xlsx";

export type CellValue = string | number | boolean | null;

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
  fillColor?: string;
  align?: "left" | "center" | "right";
  format?: "general" | "number" | "currency" | "percent";
  wrap?: boolean;
}

export interface CellData {
  raw: CellValue;
  style?: CellStyle;
  note?: string;
}

export interface SheetData {
  id: string;
  name: string;
  cells: Record<string, CellData>;
  columnWidths: Record<number, number>;
  frozenRows: number;
  frozenColumns: number;
  merges: string[];
}

export interface WorkbookData {
  id: string;
  name: string;
  sheets: SheetData[];
  activeSheetId: string;
  createdAt: string;
  updatedAt: string;
}

export interface CellPoint { row: number; col: number }
export interface Selection { start: CellPoint; end: CellPoint }

export const DEFAULT_ROWS = 100;
export const DEFAULT_COLS = 26;

export function uid(prefix = "id") {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function createSheet(name = "Sheet 1"): SheetData {
  return {
    id: uid("sheet"), name, cells: {}, columnWidths: {},
    frozenRows: 0, frozenColumns: 0, merges: [],
  };
}

export function createWorkbook(name = "Untitled"): WorkbookData {
  const sheet = createSheet();
  const now = new Date().toISOString();
  return { id: uid("book"), name, sheets: [sheet], activeSheetId: sheet.id, createdAt: now, updatedAt: now };
}

export function columnName(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

export function addressOf(row: number, col: number) { return `${columnName(col)}${row + 1}`; }

export function pointFromAddress(address: string): CellPoint | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address.trim());
  if (!match) return null;
  let col = 0;
  for (const letter of match[1].toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, col: col - 1 };
}

export function normalizedSelection(selection: Selection): Selection {
  return {
    start: { row: Math.min(selection.start.row, selection.end.row), col: Math.min(selection.start.col, selection.end.col) },
    end: { row: Math.max(selection.start.row, selection.end.row), col: Math.max(selection.start.col, selection.end.col) },
  };
}

export function selectionAddresses(selection: Selection): string[] {
  const n = normalizedSelection(selection);
  const result: string[] = [];
  for (let row = n.start.row; row <= n.end.row; row++)
    for (let col = n.start.col; col <= n.end.col; col++) result.push(addressOf(row, col));
  return result;
}

export interface MergeRegion { anchor: CellPoint; rows: number; cols: number }

export function mergeRegionOf(range: string): MergeRegion | null {
  const parts = range.split(":");
  const start = pointFromAddress(parts[0]);
  const end = pointFromAddress(parts[1] ?? parts[0]);
  if (!start || !end) return null;
  const anchor = { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) };
  return { anchor, rows: Math.abs(end.row - start.row) + 1, cols: Math.abs(end.col - start.col) + 1 };
}

export function rangeString(selection: Selection): string {
  const n = normalizedSelection(selection);
  return `${addressOf(n.start.row, n.start.col)}:${addressOf(n.end.row, n.end.col)}`;
}

function regionsOverlap(a: MergeRegion, b: MergeRegion): boolean {
  return a.anchor.row <= b.anchor.row + b.rows - 1 && b.anchor.row <= a.anchor.row + a.rows - 1
    && a.anchor.col <= b.anchor.col + b.cols - 1 && b.anchor.col <= a.anchor.col + a.cols - 1;
}

export function mergesOverlapping(merges: string[], selection: Selection): string[] {
  const target = mergeRegionOf(rangeString(selection));
  if (!target) return [];
  return merges.filter((range) => { const region = mergeRegionOf(range); return region ? regionsOverlap(region, target) : false; });
}

function parseLiteral(value: string): CellValue {
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  if (/^-?(?:\d+\.?\d*|\.\d+)$/.test(trimmed)) return Number(trimmed);
  return value;
}

export function setCellInput(sheet: SheetData, address: string, input: string): SheetData {
  const cells = { ...sheet.cells };
  const existing = cells[address];
  const raw = input.startsWith("=") ? input : parseLiteral(input);
  if (raw === null && !existing?.style && !existing?.note) delete cells[address];
  else cells[address] = { ...(existing || {}), raw };
  return { ...sheet, cells };
}

type FormulaValue = string | number | boolean | null | FormulaValue[];

function flatten(values: FormulaValue[]): FormulaValue[] {
  return values.flatMap((value) => Array.isArray(value) ? flatten(value) : [value]);
}

function numberValue(value: FormulaValue): number {
  if (Array.isArray(value)) return value.reduce<number>((sum, item) => sum + numberValue(item), 0);
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function tokenize(source: string): string[] {
  const matches = source.match(/"(?:[^"]|"")*"|'[^']+'![A-Z]+\d+|[A-Za-z_][A-Za-z0-9_.]*![A-Z]+\d+|\$?[A-Z]+\$?\d+|<=|>=|<>|!=|==|[A-Za-z_][A-Za-z0-9_.]*|\d+(?:\.\d+)?|[()+\-*/^%,:<>="&]/g);
  return matches || [];
}

export function evaluateCell(workbook: WorkbookData, sheet: SheetData, address: string, trail = new Set<string>()): FormulaValue {
  const cell = sheet.cells[address];
  if (!cell) return null;
  if (typeof cell.raw !== "string" || !cell.raw.startsWith("=")) return cell.raw;
  const key = `${sheet.id}:${address}`;
  if (trail.has(key)) return "#CIRC!";
  const nextTrail = new Set(trail).add(key);
  try {
    return evaluateFormula(cell.raw.slice(1), workbook, sheet, nextTrail);
  } catch {
    return "#ERROR!";
  }
}

export function evaluateFormula(source: string, workbook: WorkbookData, currentSheet: SheetData, trail = new Set<string>()): FormulaValue {
  const tokens = tokenize(source);
  let index = 0;
  const peek = () => tokens[index];
  const take = () => tokens[index++];
  const resolveAddress = (token: string): FormulaValue => {
    let sheet = currentSheet;
    let address = token;
    if (token.includes("!")) {
      const split = token.lastIndexOf("!");
      const sheetName = token.slice(0, split).replace(/^'|'$/g, "");
      address = token.slice(split + 1);
      sheet = workbook.sheets.find((item) => item.name.toLowerCase() === sheetName.toLowerCase()) || currentSheet;
    }
    return evaluateCell(workbook, sheet, address.replace(/\$/g, "").toUpperCase(), trail);
  };
  const rangeValues = (startToken: string, endToken: string): FormulaValue[] => {
    const start = pointFromAddress(startToken.includes("!") ? startToken.split("!").pop()! : startToken);
    const end = pointFromAddress(endToken.includes("!") ? endToken.split("!").pop()! : endToken);
    if (!start || !end) return [];
    const values: FormulaValue[] = [];
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++)
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) values.push(resolveAddress(addressOf(row, col)));
    return values;
  };
  const callFunction = (name: string, args: FormulaValue[]): FormulaValue => {
    const values = flatten(args);
    const numbers = values.map(numberValue);
    switch (name.toUpperCase()) {
      case "SUM": return numbers.reduce((a, b) => a + b, 0);
      case "AVERAGE": case "AVG": return numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0;
      case "MIN": return numbers.length ? Math.min(...numbers) : 0;
      case "MAX": return numbers.length ? Math.max(...numbers) : 0;
      case "COUNT": return values.filter((value) => typeof value === "number").length;
      case "COUNTA": return values.filter((value) => value !== null && value !== "").length;
      case "IF": return args[0] ? (args[1] ?? true) : (args[2] ?? false);
      case "AND": return values.every(Boolean);
      case "OR": return values.some(Boolean);
      case "NOT": return !args[0];
      case "ABS": return Math.abs(numberValue(args[0]));
      case "ROUND": return Number(numberValue(args[0]).toFixed(numberValue(args[1])));
      case "FLOOR": return Math.floor(numberValue(args[0]));
      case "CEILING": return Math.ceil(numberValue(args[0]));
      case "LEN": return String(args[0] ?? "").length;
      case "LOWER": return String(args[0] ?? "").toLowerCase();
      case "UPPER": return String(args[0] ?? "").toUpperCase();
      case "CONCAT": case "CONCATENATE": return values.map((value) => value ?? "").join("");
      case "TODAY": return new Date().toISOString().slice(0, 10);
      case "NOW": return new Date().toLocaleString();
      default: return `#NAME?`;
    }
  };
  const primary = (): FormulaValue => {
    const token = take();
    if (token === undefined) return null;
    if (token === "(") { const value = comparison(); if (peek() === ")") take(); return value; }
    if (token.startsWith('"')) return token.slice(1, -1).replace(/""/g, '"');
    if (/^\d/.test(token)) return Number(token);
    if (/^(TRUE|FALSE)$/i.test(token)) return token.toUpperCase() === "TRUE";
    if (/^(?:'[^']+'|[A-Za-z_][\w.]*)![A-Z]+\d+$/i.test(token) || /^\$?[A-Z]+\$?\d+$/i.test(token)) {
      if (peek() === ":") { take(); return rangeValues(token, take()); }
      return resolveAddress(token);
    }
    if (peek() === "(") {
      take();
      const args: FormulaValue[] = [];
      while (peek() !== ")" && peek() !== undefined) { args.push(comparison()); if (peek() === ",") take(); else break; }
      if (peek() === ")") take();
      return callFunction(token, args);
    }
    return token;
  };
  const unary = (): FormulaValue => { if (peek() === "-") { take(); return -numberValue(unary()); } if (peek() === "+") { take(); return numberValue(unary()); } return primary(); };
  const power = (): FormulaValue => { let left = unary(); while (peek() === "^") { take(); left = Math.pow(numberValue(left), numberValue(unary())); } return left; };
  const multiply = (): FormulaValue => { let left = power(); while (["*", "/", "%"].includes(peek())) { const op = take(); const right = numberValue(power()); left = op === "*" ? numberValue(left) * right : op === "/" ? (right === 0 ? "#DIV/0!" : numberValue(left) / right) : numberValue(left) % right; } return left; };
  const addition = (): FormulaValue => { let left = multiply(); while (["+", "-", "&"].includes(peek())) { const op = take(); const right = multiply(); left = op === "+" ? numberValue(left) + numberValue(right) : op === "-" ? numberValue(left) - numberValue(right) : `${left ?? ""}${right ?? ""}`; } return left; };
  const comparison = (): FormulaValue => { let left = addition(); while (["=", "==", "!=", "<>", "<", ">", "<=", ">="].includes(peek())) { const op = take(); const right = addition(); if (op === "=" || op === "==") left = left === right; else if (op === "!=" || op === "<>") left = left !== right; else if (op === "<") left = numberValue(left) < numberValue(right); else if (op === ">") left = numberValue(left) > numberValue(right); else if (op === "<=") left = numberValue(left) <= numberValue(right); else left = numberValue(left) >= numberValue(right); } return left; };
  return comparison();
}

export function displayValue(value: FormulaValue, style?: CellStyle): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (style?.format === "currency") return new Intl.NumberFormat(undefined, { style: "currency", currency: "USD" }).format(value);
    if (style?.format === "percent") return new Intl.NumberFormat(undefined, { style: "percent", maximumFractionDigits: 2 }).format(value);
    return new Intl.NumberFormat(undefined, { maximumFractionDigits: 8 }).format(value);
  }
  return String(value);
}

export function importWorkbook(data: Uint8Array, name: string): WorkbookData {
  const source = XLSX.read(data, { type: "array", cellStyles: true, cellDates: true });
  const sheets = source.SheetNames.map((sheetName) => {
    const ws = source.Sheets[sheetName];
    const sheet = createSheet(sheetName);
    for (const address of Object.keys(ws)) {
      if (address.startsWith("!")) continue;
      const sourceCell = ws[address];
      let raw: CellValue = sourceCell.f ? `=${sourceCell.f}` : sourceCell.v;
      if (sourceCell.v instanceof Date) raw = sourceCell.v.toISOString();
      sheet.cells[address] = { raw, style: {
        bold: Boolean(sourceCell.s?.font?.bold), italic: Boolean(sourceCell.s?.font?.italic),
        underline: Boolean(sourceCell.s?.font?.underline), textColor: sourceCell.s?.font?.color?.rgb ? `#${sourceCell.s.font.color.rgb.slice(-6)}` : undefined,
        fillColor: sourceCell.s?.fill?.fgColor?.rgb ? `#${sourceCell.s.fill.fgColor.rgb.slice(-6)}` : undefined,
        align: sourceCell.s?.alignment?.horizontal, wrap: sourceCell.s?.alignment?.wrapText,
      }};
    }
    if (ws["!cols"]) ws["!cols"].forEach((column, index) => { if (column?.wch) sheet.columnWidths[index] = Math.max(64, column.wch * 8); });
    if (ws["!merges"]) sheet.merges = ws["!merges"].map((m) => `${addressOf(m.s.r, m.s.c)}:${addressOf(m.e.r, m.e.c)}`);
    return sheet;
  });
  const first = sheets[0] || createSheet();
  const now = new Date().toISOString();
  return { id: uid("book"), name: name.replace(/\.[^.]+$/, ""), sheets: sheets.length ? sheets : [first], activeSheetId: first.id, createdAt: now, updatedAt: now };
}

export function exportWorkbook(workbook: WorkbookData, extension: string): Uint8Array {
  const output = XLSX.utils.book_new();
  for (const sheet of workbook.sheets) {
    const ws: XLSX.WorkSheet = {};
    let maxRow = 0, maxCol = 0;
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const point = pointFromAddress(address); if (!point) continue;
      maxRow = Math.max(maxRow, point.row); maxCol = Math.max(maxCol, point.col);
      const raw = cell.raw;
      const exported: XLSX.CellObject = typeof raw === "string" && raw.startsWith("=")
        ? { t: "n", f: raw.slice(1), v: Number(evaluateCell(workbook, sheet, address)) || 0 }
        : { t: typeof raw === "number" ? "n" : typeof raw === "boolean" ? "b" : "s", v: raw ?? "" } as XLSX.CellObject;
      if (cell.style?.format === "percent") exported.z = "0.00%";
      if (cell.style?.format === "currency") exported.z = "$#,##0.00";
      exported.s = { font: { bold: cell.style?.bold, italic: cell.style?.italic, underline: cell.style?.underline, color: cell.style?.textColor ? { rgb: cell.style.textColor.slice(1) } : undefined }, fill: cell.style?.fillColor ? { patternType: "solid", fgColor: { rgb: cell.style.fillColor.slice(1) } } : undefined, alignment: { horizontal: cell.style?.align, wrapText: cell.style?.wrap } };
      ws[address] = exported;
    }
    ws["!ref"] = `A1:${addressOf(maxRow, maxCol)}`;
    ws["!cols"] = Array.from({ length: maxCol + 1 }, (_, index) => ({ wch: (sheet.columnWidths[index] || 112) / 8 }));
    if (sheet.merges?.length) ws["!merges"] = sheet.merges.map((range) => { const region = mergeRegionOf(range)!; return { s: { r: region.anchor.row, c: region.anchor.col }, e: { r: region.anchor.row + region.rows - 1, c: region.anchor.col + region.cols - 1 } }; }).filter(Boolean);
    XLSX.utils.book_append_sheet(output, ws, sheet.name.slice(0, 31));
  }
  if (extension === "csv") {
    const active = workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook.sheets[0];
    return new TextEncoder().encode(XLSX.utils.sheet_to_csv(output.Sheets[active.name]));
  }
  return XLSX.write(output, { type: "array", bookType: "xlsx", cellStyles: true });
}

export function sampleWorkbook(): WorkbookData {
  const workbook = createWorkbook("Untitled");
  const sheet = workbook.sheets[0];
  sheet.cells = {
    A1: { raw: "Quarterly sales", style: { bold: true, textColor: "#173d33", fillColor: "#dceee7" } },
    A3: { raw: "Region", style: { bold: true, fillColor: "#edf2ec" } },
    B3: { raw: "Q1", style: { bold: true, fillColor: "#edf2ec", align: "right" } },
    C3: { raw: "Q2", style: { bold: true, fillColor: "#edf2ec", align: "right" } },
    D3: { raw: "Total", style: { bold: true, fillColor: "#edf2ec", align: "right" } },
    A4: { raw: "North" }, B4: { raw: 12800 }, C4: { raw: 15100 }, D4: { raw: "=SUM(B4:C4)", style: { bold: true } },
    A5: { raw: "South" }, B5: { raw: 9900 }, C5: { raw: 13450 }, D5: { raw: "=SUM(B5:C5)", style: { bold: true } },
    A6: { raw: "West" }, B6: { raw: 14300 }, C6: { raw: 16200 }, D6: { raw: "=SUM(B6:C6)", style: { bold: true } },
    A7: { raw: "Total", style: { bold: true } }, B7: { raw: "=SUM(B4:B6)", style: { bold: true } }, C7: { raw: "=SUM(C4:C6)", style: { bold: true } }, D7: { raw: "=SUM(D4:D6)", style: { bold: true, fillColor: "#dceee7" } },
  };
  return workbook;
}
