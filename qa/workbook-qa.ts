import {
  addressOf, createSheet, createWorkbook, displayValue, evaluateCell, exportWorkbook,
  importWorkbook, setCellInput,
} from "../src/workbook";

type Result = { area: string; name: string; status: "PASS" | "FAIL"; expected?: unknown; actual?: unknown; detail?: string };
const results: Result[] = [];

function check(area: string, name: string, actual: unknown, expected: unknown) {
  const pass = JSON.stringify(actual) === JSON.stringify(expected);
  results.push({ area, name, status: pass ? "PASS" : "FAIL", ...(pass ? {} : { expected, actual }) });
}

function checkTruthy(area: string, name: string, actual: unknown, detail?: string) {
  results.push({ area, name, status: actual ? "PASS" : "FAIL", ...(actual ? {} : { actual, detail }) });
}

// Addressing and basic editing model.
check("Cells", "Address beyond Z", addressOf(9, 27), "AB10");
const literals = createWorkbook("Literals");
let literalSheet = literals.sheets[0];
literalSheet = setCellInput(literalSheet, "A1", "42");
literalSheet = setCellInput(literalSheet, "A2", "true");
literalSheet = setCellInput(literalSheet, "A3", "hello");
literals.sheets[0] = literalSheet;
check("Cells", "Number parsing", literalSheet.cells.A1.raw, 42);
check("Cells", "Boolean parsing", literalSheet.cells.A2.raw, true);
check("Cells", "Text parsing", literalSheet.cells.A3.raw, "hello");

// Core formula expectations.
let formulas = createWorkbook("Formulas");
let formulaSheet = formulas.sheets[0];
formulaSheet = setCellInput(formulaSheet, "A1", "10");
formulaSheet = setCellInput(formulaSheet, "A2", "20");
formulaSheet = setCellInput(formulaSheet, "A3", "=SUM(A1:A2) * 2");
formulaSheet = setCellInput(formulaSheet, "B1", "=IF(A1>5,\"yes\",\"no\")");
formulaSheet = setCellInput(formulaSheet, "B2", "=1/0");
formulas.sheets[0] = formulaSheet;
check("Formulas", "Arithmetic and SUM range", evaluateCell(formulas, formulaSheet, "A3"), 60);
check("Formulas", "IF and comparison", evaluateCell(formulas, formulaSheet, "B1"), "yes");
check("Formulas", "Division-by-zero error", evaluateCell(formulas, formulaSheet, "B2"), "#DIV/0!");

let blankMath = createWorkbook("Blank math");
let blankSheet = blankMath.sheets[0];
blankSheet = setCellInput(blankSheet, "A1", "10");
blankSheet = setCellInput(blankSheet, "B1", "=AVERAGE(A1:A2)");
blankSheet = setCellInput(blankSheet, "B2", "=MIN(A1:A2)");
blankMath.sheets[0] = blankSheet;
check("Formulas", "AVERAGE ignores blank cells", evaluateCell(blankMath, blankSheet, "B1"), 10);
check("Formulas", "MIN ignores blank cells", evaluateCell(blankMath, blankSheet, "B2"), 10);

let crossSheet = createWorkbook("Cross sheet");
crossSheet.sheets[0].name = "Summary";
let details = createSheet("Details");
details = setCellInput(details, "A1", "5");
details = setCellInput(details, "A2", "7");
crossSheet.sheets.push(details);
crossSheet.sheets[0] = setCellInput(crossSheet.sheets[0], "A1", "=SUM(Details!A1:A2)");
check("Formulas", "Cross-sheet range", evaluateCell(crossSheet, crossSheet.sheets[0], "A1"), 12);

let invalidFormula = createWorkbook("Invalid formula");
invalidFormula.sheets[0] = setCellInput(invalidFormula.sheets[0], "A1", "=1@2");
check("Formulas", "Invalid token reports error", evaluateCell(invalidFormula, invalidFormula.sheets[0], "A1"), "#ERROR!");

let circular = createWorkbook("Circular");
circular.sheets[0] = setCellInput(circular.sheets[0], "A1", "=B1");
circular.sheets[0] = setCellInput(circular.sheets[0], "B1", "=A1");
check("Formulas", "Circular reference detection", evaluateCell(circular, circular.sheets[0], "A1"), "#CIRC!");

// Display formatting.
check("Formatting", "Currency display", displayValue(1234.5, { format: "currency" }).includes("1,234.50"), true);
check("Formatting", "Percent display", displayValue(0.125, { format: "percent" }), "12.5%");
check("Formatting", "Date display changes numeric serial", displayValue(45292, { format: "date" }) !== displayValue(45292, { format: "general" }), true);

// XLSX round-trip with multiple sheets, formulas, Unicode, and styles.
let roundTrip = createWorkbook("Round trip");
roundTrip.sheets[0].name = "Overview";
roundTrip.sheets[0] = setCellInput(roundTrip.sheets[0], "A1", "Résumé ✓");
roundTrip.sheets[0] = setCellInput(roundTrip.sheets[0], "B1", "1250");
roundTrip.sheets[0] = setCellInput(roundTrip.sheets[0], "C1", "=B1*2");
roundTrip.sheets[0].cells.A1.style = { bold: true, fillColor: "#dceee7" };
const second = createSheet("Second");
second.cells.A1 = { raw: "second sheet" };
roundTrip.sheets.push(second);
const xlsxBytes = exportWorkbook(roundTrip, "xlsx");
const restored = importWorkbook(xlsxBytes, "Round trip.xlsx");
check("XLSX", "Sheet count round-trip", restored.sheets.length, 2);
check("XLSX", "Unicode round-trip", restored.sheets[0].cells.A1.raw, "Résumé ✓");
check("XLSX", "Number round-trip", restored.sheets[0].cells.B1.raw, 1250);
check("XLSX", "Formula round-trip", restored.sheets[0].cells.C1.raw, "=B1*2");
check("XLSX", "Second sheet round-trip", restored.sheets[1].cells.A1.raw, "second sheet");
check("XLSX", "Bold style round-trip", restored.sheets[0].cells.A1.style?.bold, true);
check("XLSX", "Fill style round-trip", restored.sheets[0].cells.A1.style?.fillColor?.toLowerCase(), "#dceee7");

// CSV and TSV compatibility.
let delimited = createWorkbook("Delimited");
delimited.sheets[0] = setCellInput(delimited.sheets[0], "A1", "name");
delimited.sheets[0] = setCellInput(delimited.sheets[0], "B1", "value");
delimited.sheets[0] = setCellInput(delimited.sheets[0], "A2", "comma, quote \" and ✓");
delimited.sheets[0] = setCellInput(delimited.sheets[0], "B2", "21");
delimited.sheets[0] = setCellInput(delimited.sheets[0], "C2", "=B2*2");
const csvBytes = exportWorkbook(delimited, "csv");
const csvRestored = importWorkbook(csvBytes, "Delimited.csv");
check("CSV", "Quoted Unicode text round-trip", csvRestored.sheets[0].cells.A2.raw, "comma, quote \" and ✓");
check("CSV", "Number round-trip", csvRestored.sheets[0].cells.B2.raw, 21);
check("CSV", "Formula exported as computed value", csvRestored.sheets[0].cells.C2.raw, 42);

const tsvBytes = new TextEncoder().encode("Name\tScore\nAda\t98\nLinus\t91");
const tsvRestored = importWorkbook(tsvBytes, "Scores.tsv");
check("TSV", "TSV text import", tsvRestored.sheets[0].cells.A2.raw, "Ada");
check("TSV", "TSV number import", tsvRestored.sheets[0].cells.B3.raw, 91);

// Larger workbook engine boundary.
let large = createWorkbook("Large");
large.sheets[0] = setCellInput(large.sheets[0], "A501", "row 501");
large.sheets[0] = setCellInput(large.sheets[0], "CW1", "column 101");
const largeRestored = importWorkbook(exportWorkbook(large, "xlsx"), "Large.xlsx");
check("Scale", "Engine preserves row 501", largeRestored.sheets[0].cells.A501.raw, "row 501");
check("Scale", "Engine preserves column 101", largeRestored.sheets[0].cells.CW1.raw, "column 101");
checkTruthy("Scale", "XLSX output is non-empty", xlsxBytes.byteLength > 1000, `Only ${xlsxBytes.byteLength} bytes`);

const passed = results.filter((result) => result.status === "PASS").length;
const failed = results.filter((result) => result.status === "FAIL").length;
console.log(JSON.stringify({ summary: { total: results.length, passed, failed }, results }, null, 2));
