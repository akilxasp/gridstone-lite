import fs from "node:fs";

const app = fs.readFileSync(new URL("../src/App.tsx", import.meta.url), "utf8");
const workbook = fs.readFileSync(new URL("../src/workbook.ts", import.meta.url), "utf8");
const electron = fs.readFileSync(new URL("../electron/main.cjs", import.meta.url), "utf8");

const results = [];
const check = (area, name, pass, detail) => results.push({ area, name, status: pass ? "PASS" : "FAIL", ...(pass ? {} : { detail }) });

check("Editing", "Inline editor has local input state", /function CellEditor[\s\S]*useState\(initialValue\)/.test(app), "Cell keystrokes should not update top-level App state.");
check("Editing", "Formula editor has local input state", /function FormulaEditor[\s\S]*useState\(initialValue\)/.test(app), "Formula-bar keystrokes should not update top-level App state.");
check("Editing", "Duplicate edit commit guard", /finished\.current/.test(app), "Blur and Enter can otherwise commit the same edit twice.");
check("Accessibility", "Spreadsheet exposes grid semantics", /role="grid"/.test(app) && /role="gridcell"/.test(app), "Grid and cells need accessibility roles.");
check("Accessibility", "Dynamic status is announced", /aria-live="polite"/.test(app), "Status changes need a live region.");
check("File", "XLSX and CSV export paths exist", /exportWorkbook\(current, extension\)/.test(app) && /saveFile\(true, "csv"\)/.test(app), "Both export paths should exist.");
check("File", "Save As format follows selected dialog format", !/filters: extension === "csv"[\s\S]*Excel workbook[\s\S]*CSV file/.test(electron), "Renderer serializes data before the user chooses XLSX or CSV in the native dialog; choosing CSV from Save As can write XLSX bytes with a .csv name.");
check("File", "New Workbook is blank", !/const next = sampleWorkbook\(\)/.test(app), "New Workbook currently reloads the demonstration sales data.");
check("Formatting", "Font-family selector changes cell style", /aria-label="Font family"[^>]*onChange=/.test(app), "Font-family selector has no change handler.");
check("Formatting", "Font-size selector changes cell style", /aria-label="Font size"[^>]*onChange=/.test(app), "Font-size selector has no change handler.");
check("Formatting", "Date format has a display implementation", /style\?\.format === "date"/.test(workbook), "Date is selectable but displayValue does not format dates.");
check("Charts", "Add chart persists a chart model", /charts\s*[:=]/.test(app) || /charts\s*[:=]/.test(workbook), "Add chart only closes the preview and changes status text.");
check("View", "Frozen rows affect grid rendering", (app.match(/frozenRows/g) || []).length > 2, "Freeze top row changes model state but rendering never reads it.");
check("View", "Zoom changes cell geometry", /gridTemplateColumns[\s\S]*zoom/.test(app), "Zoom changes font size only; row heights and column widths do not scale.");
check("Scale", "All imported rows remain reachable", !/Math\.min\(rows, 500\)/.test(app), "Rows after 500 are preserved by the file engine but not rendered.");
check("Scale", "All imported columns remain reachable", !/Math\.min\(cols, 100\)/.test(app), "Columns after 100 are preserved by the file engine but not rendered.");
check("Notes", "Typing a note is buffered before workbook history", !/textarea[\s\S]{0,400}onChange=[\s\S]{0,300}updateActiveSheet/.test(app), "Every note keystroke writes the workbook and consumes an undo entry.");
check("Collaboration", "Share performs a real share workflow", !/Cloud collaboration can be connected from Settings/.test(app), "Share is currently a placeholder status message.");
check("Desktop", "Single-instance file-open handling", /requestSingleInstanceLock/.test(electron), "Windows can open a second app instance instead of forwarding a double-clicked workbook.");

const passed = results.filter((result) => result.status === "PASS").length;
const failed = results.length - passed;
console.log(JSON.stringify({ summary: { total: results.length, passed, failed }, results }, null, 2));
