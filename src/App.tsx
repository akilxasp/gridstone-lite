import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownWideNarrow, ArrowUpNarrowWide,
  BarChart3, Bold, Check, ChevronDown, Columns, Combine, Copy, DollarSign, Download, FilePlus2,
  FolderOpen, Italic, Moon, PaintBucket, PanelRight, Percent, Plus, Printer,
  Redo2, Save, Scissors, Search, Snowflake, Sun, Trash2, Underline, Undo2, X,
} from "lucide-react";
import {
  CellData, CellStyle, DEFAULT_COLS, DEFAULT_ROWS, MergeRegion, Selection, SheetData, WorkbookData,
  addressOf, columnName, createSheet, displayValue, evaluateCell, exportWorkbook,
  importWorkbook, mergeRegionOf, mergesOverlapping, normalizedSelection, pointFromAddress,
  rangeString, sampleWorkbook, selectionAddresses, setCellInput,
} from "./workbook";

type RibbonTab = "Home" | "Insert" | "Data" | "View";
type FileState = { path?: string; name: string; extension: string };
type EditSession = { address: string; initialValue: string };

const INITIAL_SELECTION: Selection = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

function clone<T>(value: T): T { return structuredClone(value); }

function IconButton({ label, active, disabled, children, onClick }: { label: string; active?: boolean; disabled?: boolean; children: React.ReactNode; onClick?: () => void }) {
  return <button className={`icon-button ${active ? "is-active" : ""}`} type="button" title={label} aria-label={label} aria-pressed={active} disabled={disabled} onClick={onClick}>{children}</button>;
}

function ToolButton({ label, icon, onClick, disabled }: { label: string; icon: React.ReactNode; onClick?: () => void; disabled?: boolean }) {
  return <button className="tool-button" type="button" onClick={onClick} disabled={disabled}>{icon}<span>{label}</span></button>;
}

function ColorControl({ label, value, icon, onChange }: { label: string; value: string; icon: React.ReactNode; onChange: (value: string) => void }) {
  return <label className="color-control" title={label} aria-label={label}>{icon}<span className="color-swatch" style={{ background: value }} /><input type="color" value={value} onChange={(event) => onChange(event.target.value)} /></label>;
}

function CellEditor({ initialValue, onCommit, onCancel }: { initialValue: string; onCommit: (value: string) => void; onCancel: () => void }) {
  const [value, setValue] = useState(initialValue);
  const inputRef = useRef<HTMLInputElement>(null);
  const finished = useRef(false);
  const finish = useCallback(() => {
    if (finished.current) return;
    finished.current = true;
    onCommit(value);
  }, [onCommit, value]);

  useEffect(() => { inputRef.current?.focus(); inputRef.current?.select(); }, []);

  return <input
    ref={inputRef}
    className="cell-editor"
    value={value}
    aria-label="Edit cell value"
    onMouseDown={(event) => event.stopPropagation()}
    onChange={(event) => setValue(event.target.value)}
    onBlur={finish}
    onKeyDown={(event) => {
      event.stopPropagation();
      if (event.key === "Enter" || event.key === "Tab") { event.preventDefault(); finish(); }
      if (event.key === "Escape") { event.preventDefault(); finished.current = true; onCancel(); }
    }}
  />;
}

function FormulaEditor({ address, initialValue, onCommit, onFocusGrid }: { address: string; initialValue: string; onCommit: (address: string, value: string) => void; onFocusGrid: () => void }) {
  const [value, setValue] = useState(initialValue);
  const [active, setActive] = useState(false);
  const finished = useRef(false);

  useEffect(() => { if (!active) setValue(initialValue); }, [initialValue, active]);

  const finish = useCallback(() => {
    if (finished.current) return;
    finished.current = true;
    setActive(false);
    if (value !== initialValue) onCommit(address, value);
  }, [address, initialValue, onCommit, value]);

  const cancel = useCallback(() => {
    finished.current = true;
    setValue(initialValue);
    setActive(false);
    onFocusGrid();
  }, [initialValue, onFocusGrid]);

  return <>
    <input
      aria-label={`Value or formula for ${address}`}
      value={value}
      onFocus={() => { finished.current = false; setActive(true); }}
      onChange={(event) => setValue(event.target.value)}
      onBlur={finish}
      onKeyDown={(event) => {
        if (event.key === "Enter") { event.preventDefault(); finish(); onFocusGrid(); }
        if (event.key === "Escape") { event.preventDefault(); cancel(); }
      }}
    />
    {active && <div className="formula-actions">
      <button aria-label="Cancel edit" onMouseDown={(event) => event.preventDefault()} onClick={cancel}><X size={16} /></button>
      <button aria-label="Accept edit" onMouseDown={(event) => event.preventDefault()} onClick={() => { finish(); onFocusGrid(); }}><Check size={16} /></button>
    </div>}
  </>;
}

export default function App() {
  const [workbook, setWorkbook] = useState<WorkbookData>(() => sampleWorkbook());
  const [selection, setSelection] = useState<Selection>(INITIAL_SELECTION);
  const [file, setFile] = useState<FileState>({ name: "Untitled.xlsx", extension: "xlsx" });
  const [dirty, setDirty] = useState(false);
  const [ribbon, setRibbon] = useState<RibbonTab>("Home");
  const [editing, setEditing] = useState<EditSession | null>(null);
  const [dragging, setDragging] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => (localStorage.getItem("gridstone-theme") as "light" | "dark") || "light");
  const [status, setStatus] = useState("Ready");
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState("");
  const [chartOpen, setChartOpen] = useState(false);
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [showFileMenu, setShowFileMenu] = useState(false);
  const [fileMenuPos, setFileMenuPos] = useState({ top: 0, left: 0 });
  const [zoom, setZoom] = useState(100);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const fileBtnRef = useRef<HTMLButtonElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const dragKind = useRef<"cell" | "row" | "col">("cell");
  const headerAnchor = useRef(0);
  const undoStack = useRef<WorkbookData[]>([]);
  const redoStack = useRef<WorkbookData[]>([]);
  const workbookRef = useRef(workbook);
  const fileRef = useRef(file);

  useEffect(() => { workbookRef.current = workbook; }, [workbook]);
  useEffect(() => { fileRef.current = file; }, [file]);

  const activeSheet = useMemo(() => workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook.sheets[0], [workbook]);
  const selectedAddress = addressOf(selection.end.row, selection.end.col);
  const selectedCell = activeSheet.cells[selectedAddress];

  const dimensions = useMemo(() => {
    let rows = DEFAULT_ROWS, cols = DEFAULT_COLS;
    Object.keys(activeSheet.cells).forEach((address) => { const point = pointFromAddress(address); if (point) { rows = Math.max(rows, point.row + 20); cols = Math.max(cols, point.col + 8); } });
    (activeSheet.merges || []).forEach((range) => { const region = mergeRegionOf(range); if (region) { rows = Math.max(rows, region.anchor.row + region.rows); cols = Math.max(cols, region.anchor.col + region.cols); } });
    return { rows: Math.min(rows, 500), cols: Math.min(cols, 100) };
  }, [activeSheet.cells, activeSheet.merges]);

  const mergeMap = useMemo(() => {
    const anchors = new Map<string, MergeRegion>();
    const covered = new Set<string>();
    (activeSheet.merges || []).forEach((range) => {
      const region = mergeRegionOf(range); if (!region) return;
      anchors.set(`${region.anchor.row},${region.anchor.col}`, region);
      for (let row = region.anchor.row; row < region.anchor.row + region.rows; row++)
        for (let col = region.anchor.col; col < region.anchor.col + region.cols; col++)
          if (row !== region.anchor.row || col !== region.anchor.col) covered.add(`${row},${col}`);
    });
    return { anchors, covered };
  }, [activeSheet.merges]);

  const colOffsets = useMemo(() => {
    const offsets: number[] = [48];
    for (let col = 0; col < dimensions.cols; col++) offsets.push(offsets[col] + (activeSheet.columnWidths[col] || 112));
    return offsets;
  }, [dimensions.cols, activeSheet.columnWidths]);

  const commit = useCallback((updater: (current: WorkbookData) => WorkbookData, message = "Updated") => {
    setWorkbook((current) => {
      undoStack.current.push(clone(current));
      if (undoStack.current.length > 100) undoStack.current.shift();
      redoStack.current = [];
      const next = updater(current);
      return { ...next, updatedAt: new Date().toISOString() };
    });
    setDirty(true);
    setStatus(message);
  }, []);

  const updateActiveSheet = useCallback((updater: (sheet: SheetData) => SheetData, message?: string) => {
    commit((current) => ({ ...current, sheets: current.sheets.map((sheet) => sheet.id === current.activeSheetId ? updater(sheet) : sheet) }), message);
  }, [commit]);

  const commitEdit = useCallback((address: string, value: string) => {
    updateActiveSheet((sheet) => setCellInput(sheet, address, value), `Updated ${address}`);
    setEditing((current) => current?.address === address ? null : current);
    setTimeout(() => gridRef.current?.focus(), 0);
  }, [updateActiveSheet]);

  const beginEdit = useCallback((address: string, seed?: string) => {
    setEditing({ address, initialValue: seed ?? String(activeSheet.cells[address]?.raw ?? "") });
  }, [activeSheet]);

  const undo = useCallback(() => {
    const previous = undoStack.current.pop(); if (!previous) return;
    redoStack.current.push(clone(workbookRef.current)); setWorkbook(previous); setDirty(true); setStatus("Undid last change");
  }, []);
  const redo = useCallback(() => {
    const next = redoStack.current.pop(); if (!next) return;
    undoStack.current.push(clone(workbookRef.current)); setWorkbook(next); setDirty(true); setStatus("Redid change");
  }, []);

  const newWorkbook = useCallback(() => {
    const next = sampleWorkbook();
    undoStack.current = []; redoStack.current = []; setWorkbook(next); setFile({ name: "Untitled.xlsx", extension: "xlsx" });
    localStorage.removeItem("gridstone-recovery");
    setSelection(INITIAL_SELECTION); setDirty(false); setStatus("New workbook created"); setShowFileMenu(false);
  }, []);

  const loadDesktopFile = useCallback((opened: DesktopFile) => {
    try {
      const next = importWorkbook(new Uint8Array(opened.data), opened.name);
      setWorkbook(next); setFile({ path: opened.path, name: opened.name, extension: opened.extension });
      localStorage.removeItem("gridstone-recovery");
      undoStack.current = []; redoStack.current = []; setSelection(INITIAL_SELECTION); setDirty(false);
      const recent = JSON.parse(localStorage.getItem("gridstone-recent") || "[]") as FileState[];
      localStorage.setItem("gridstone-recent", JSON.stringify([{ path: opened.path, name: opened.name, extension: opened.extension }, ...recent.filter((item) => item.path !== opened.path)].slice(0, 8)));
      setStatus(`Opened ${opened.name}`); setShowFileMenu(false);
    } catch (error) { setStatus(`Could not open file: ${error instanceof Error ? error.message : "Unknown format"}`); }
  }, []);

  const openFile = useCallback(async () => {
    if (window.desktop) { const opened = await window.desktop.openFile(); if (opened) loadDesktopFile(opened); }
    else fileInputRef.current?.click();
  }, [loadDesktopFile]);

  const openRecent = useCallback(async (item: FileState) => {
    if (!item.path || !window.desktop) return;
    try { loadDesktopFile(await window.desktop.readPath(item.path)); }
    catch { setStatus(`Could not reopen ${item.name}`); }
  }, [loadDesktopFile]);

  const saveFile = useCallback(async (saveAs = false, forceExtension?: string) => {
    const current = workbookRef.current;
    const currentFile = fileRef.current;
    const extension = forceExtension || currentFile.extension || "xlsx";
    const data = exportWorkbook(current, extension);
    if (window.desktop) {
      const saved = await window.desktop.saveFile({ path: forceExtension ? undefined : currentFile.path, saveAs: saveAs || Boolean(forceExtension), extension, suggestedName: `${current.name || "Untitled"}.${extension}`, data });
      if (saved) { setFile(saved); setDirty(false); localStorage.removeItem("gridstone-recovery"); setStatus(`Saved ${saved.name}`); }
    } else {
      const blobBytes = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
      const blob = new Blob([blobBytes], { type: extension === "csv" ? "text/csv" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = `${current.name || "Untitled"}.${extension}`; anchor.click(); URL.revokeObjectURL(url);
      setDirty(false); localStorage.removeItem("gridstone-recovery"); setStatus(`Exported ${anchor.download}`);
    }
    setShowFileMenu(false);
  }, []);

  const printWorkbook = useCallback(async () => {
    setStatus("Opening print dialog…");
    if (window.desktop) await window.desktop.print(); else window.print();
  }, []);

  const applyStyle = useCallback((patch: Partial<CellStyle>) => {
    const addresses = selectionAddresses(selection);
    updateActiveSheet((sheet) => {
      const cells = { ...sheet.cells };
      for (const address of addresses) {
        const cell = cells[address] || { raw: null };
        cells[address] = { ...cell, style: { ...cell.style, ...patch } };
      }
      return { ...sheet, cells };
    }, `Formatted ${addresses.length} cell${addresses.length === 1 ? "" : "s"}`);
  }, [selection, updateActiveSheet]);

  const clearSelection = useCallback(() => {
    const addresses = new Set(selectionAddresses(selection));
    updateActiveSheet((sheet) => ({ ...sheet, cells: Object.fromEntries(Object.entries(sheet.cells).filter(([address]) => !addresses.has(address))) }), "Cleared cells");
  }, [selection, updateActiveSheet]);

  const copySelection = useCallback(async (cut = false) => {
    const normalized = normalizedSelection(selection);
    const rows: string[] = [];
    for (let row = normalized.start.row; row <= normalized.end.row; row++) {
      const values: string[] = [];
      for (let col = normalized.start.col; col <= normalized.end.col; col++) values.push(String(activeSheet.cells[addressOf(row, col)]?.raw ?? ""));
      rows.push(values.join("\t"));
    }
    await navigator.clipboard.writeText(rows.join("\n"));
    if (cut) clearSelection();
    setStatus(cut ? "Cut cells" : "Copied cells");
  }, [selection, activeSheet, clearSelection]);

  const pasteSelection = useCallback(async () => {
    const text = await navigator.clipboard.readText();
    const rows = text.replace(/\r/g, "").split("\n").map((row) => row.split("\t"));
    updateActiveSheet((sheet) => {
      let next = sheet;
      rows.forEach((values, rowOffset) => values.forEach((value, colOffset) => { next = setCellInput(next, addressOf(selection.end.row + rowOffset, selection.end.col + colOffset), value); }));
      return next;
    }, `Pasted ${rows.length} row${rows.length === 1 ? "" : "s"}`);
  }, [selection, updateActiveSheet]);

  const sortSelection = useCallback((direction: "asc" | "desc") => {
    const normalized = normalizedSelection(selection);
    updateActiveSheet((sheet) => {
      const rows = [] as CellData[][];
      for (let row = normalized.start.row; row <= normalized.end.row; row++) {
        const values = [] as CellData[];
        for (let col = normalized.start.col; col <= normalized.end.col; col++) values.push(sheet.cells[addressOf(row, col)] || { raw: null });
        rows.push(values);
      }
      rows.sort((a, b) => String(a[0].raw ?? "").localeCompare(String(b[0].raw ?? ""), undefined, { numeric: true }) * (direction === "asc" ? 1 : -1));
      const cells = { ...sheet.cells };
      rows.forEach((values, rowOffset) => values.forEach((cell, colOffset) => { const address = addressOf(normalized.start.row + rowOffset, normalized.start.col + colOffset); if (cell.raw === null) delete cells[address]; else cells[address] = cell; }));
      return { ...sheet, cells };
    }, `Sorted ${direction === "asc" ? "ascending" : "descending"}`);
  }, [selection, updateActiveSheet]);

  const removeDuplicates = useCallback(() => {
    const normalized = normalizedSelection(selection);
    updateActiveSheet((sheet) => {
      const seen = new Set<string>(); const kept: CellData[][] = [];
      for (let row = normalized.start.row; row <= normalized.end.row; row++) {
        const values = Array.from({ length: normalized.end.col - normalized.start.col + 1 }, (_, index) => sheet.cells[addressOf(row, normalized.start.col + index)] || { raw: null });
        const key = JSON.stringify(values.map((cell) => cell.raw)); if (!seen.has(key)) { seen.add(key); kept.push(values); }
      }
      const cells = { ...sheet.cells };
      for (let row = normalized.start.row; row <= normalized.end.row; row++) for (let col = normalized.start.col; col <= normalized.end.col; col++) delete cells[addressOf(row, col)];
      kept.forEach((values, rowOffset) => values.forEach((cell, colOffset) => { if (cell.raw !== null) cells[addressOf(normalized.start.row + rowOffset, normalized.start.col + colOffset)] = cell; }));
      return { ...sheet, cells };
    }, "Removed duplicate rows");
  }, [selection, updateActiveSheet]);

  const toggleMerge = useCallback(() => {
    const overlapping = mergesOverlapping(activeSheet.merges || [], selection);
    if (overlapping.length) {
      updateActiveSheet((sheet) => ({ ...sheet, merges: (sheet.merges || []).filter((range) => !overlapping.includes(range)) }), "Unmerged cells");
      return;
    }
    const n = normalizedSelection(selection);
    if (n.start.row === n.end.row && n.start.col === n.end.col) { setStatus("Select at least two cells to merge"); return; }
    const range = rangeString(selection);
    const anchor = addressOf(n.start.row, n.start.col);
    updateActiveSheet((sheet) => {
      const cells = { ...sheet.cells };
      const anchorCell = cells[anchor] || { raw: null };
      cells[anchor] = { ...anchorCell, style: { ...anchorCell.style, align: anchorCell.style?.align || "center" } };
      return { ...sheet, merges: [...(sheet.merges || []), range] };
    }, "Merged cells");
  }, [activeSheet.merges, selection, updateActiveSheet]);

  const selectColumn = useCallback((col: number, extend: boolean) => {
    dragKind.current = "col"; headerAnchor.current = col; setDragging(true);
    const anchor = extend ? selection.start.col : col;
    setSelection({ start: { row: 0, col: Math.min(anchor, col) }, end: { row: dimensions.rows - 1, col: Math.max(anchor, col) } });
    gridRef.current?.focus();
  }, [selection.start.col, dimensions.rows]);

  const selectRow = useCallback((row: number, extend: boolean) => {
    dragKind.current = "row"; headerAnchor.current = row; setDragging(true);
    const anchor = extend ? selection.start.row : row;
    setSelection({ start: { row: Math.min(anchor, row), col: 0 }, end: { row: Math.max(anchor, row), col: dimensions.cols - 1 } });
    gridRef.current?.focus();
  }, [selection.start.row, dimensions.cols]);

  const selectAll = useCallback(() => {
    dragKind.current = "cell"; setSelection({ start: { row: 0, col: 0 }, end: { row: dimensions.rows - 1, col: dimensions.cols - 1 } }); gridRef.current?.focus();
  }, [dimensions.rows, dimensions.cols]);

  const addSheet = useCallback(() => {
    const sheet = createSheet(`Sheet ${workbook.sheets.length + 1}`);
    commit((current) => ({ ...current, sheets: [...current.sheets, sheet], activeSheetId: sheet.id }), "Added worksheet");
    setSelection(INITIAL_SELECTION);
  }, [workbook.sheets.length, commit]);

  const deleteSheet = useCallback((id: string) => {
    if (workbook.sheets.length === 1) { setStatus("A workbook must contain at least one sheet"); return; }
    commit((current) => { const sheets = current.sheets.filter((sheet) => sheet.id !== id); return { ...current, sheets, activeSheetId: current.activeSheetId === id ? sheets[0].id : current.activeSheetId }; }, "Deleted worksheet");
  }, [workbook.sheets.length, commit]);

  const renameSheet = useCallback((id: string) => {
    const current = workbook.sheets.find((sheet) => sheet.id === id); if (!current) return;
    const name = window.prompt("Worksheet name", current.name)?.trim();
    if (name) commit((book) => ({ ...book, sheets: book.sheets.map((sheet) => sheet.id === id ? { ...sheet, name: name.slice(0, 31) } : sheet) }), "Renamed worksheet");
  }, [workbook.sheets, commit]);

  const handleGridKeyDown = useCallback((event: React.KeyboardEvent) => {
    if (editing) return;
    const modifier = event.metaKey || event.ctrlKey;
    if (modifier && event.key.toLowerCase() === "c") { event.preventDefault(); void copySelection(); return; }
    if (modifier && event.key.toLowerCase() === "x") { event.preventDefault(); void copySelection(true); return; }
    if (modifier && event.key.toLowerCase() === "v") { event.preventDefault(); void pasteSelection(); return; }
    if (modifier && event.key.toLowerCase() === "z") { event.preventDefault(); event.shiftKey ? redo() : undo(); return; }
    if (["Delete", "Backspace"].includes(event.key)) { event.preventDefault(); clearSelection(); return; }
    if (event.key === "Enter" || event.key === "F2") { event.preventDefault(); beginEdit(selectedAddress); return; }
    if (event.key.length === 1 && !modifier && !event.altKey) { event.preventDefault(); beginEdit(selectedAddress, event.key); return; }
    const delta: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, event.shiftKey ? -1 : 1] };
    if (delta[event.key]) {
      event.preventDefault(); const [row, col] = delta[event.key];
      const point = { row: Math.max(0, selection.end.row + row), col: Math.max(0, selection.end.col + col) };
      setSelection(event.shiftKey && event.key !== "Tab" ? { ...selection, end: point } : { start: point, end: point });
    }
  }, [editing, copySelection, pasteSelection, redo, undo, clearSelection, beginEdit, selectedAddress, selection]);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    if (window.desktop?.platform) document.documentElement.dataset.platform = window.desktop.platform;
    localStorage.setItem("gridstone-theme", theme);
  }, [theme]);

  useEffect(() => {
    try {
      const recovery = JSON.parse(localStorage.getItem("gridstone-recovery") || "null") as { workbook: WorkbookData; file: FileState; savedAt: number } | null;
      if (recovery && window.confirm(`Gridstone found an unsaved recovery copy from ${new Date(recovery.savedAt).toLocaleString()}. Restore it?`)) {
        setWorkbook(recovery.workbook); setFile(recovery.file); setDirty(true); setStatus("Recovered unsaved workbook");
      }
    } catch { localStorage.removeItem("gridstone-recovery"); }
  }, []);

  useEffect(() => {
    const timer = window.setInterval(() => {
      if (dirty) { localStorage.setItem("gridstone-recovery", JSON.stringify({ workbook: workbookRef.current, file: fileRef.current, savedAt: Date.now() })); setStatus("Recovery copy saved"); }
    }, 5000);
    return () => window.clearInterval(timer);
  }, [dirty]);

  useEffect(() => {
    const beforeUnload = (event: BeforeUnloadEvent) => { if (dirty) { event.preventDefault(); event.returnValue = ""; } };
    window.addEventListener("beforeunload", beforeUnload); return () => window.removeEventListener("beforeunload", beforeUnload);
  }, [dirty]);

  useEffect(() => {
    if (!window.desktop) return;
    return window.desktop.onCommand((command, payload) => {
      if (command === "new") newWorkbook(); else if (command === "open") void openFile(); else if (command === "save") void saveFile();
      else if (command === "save-as") void saveFile(true); else if (command === "print") void printWorkbook(); else if (command === "find") setFindOpen(true);
      else if (command === "open-path" && typeof payload === "string") void window.desktop?.readPath(payload).then(loadDesktopFile).catch(() => setStatus("Could not open the selected workbook"));
    });
  }, [newWorkbook, openFile, saveFile, printWorkbook, loadDesktopFile]);

  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      const modifier = event.metaKey || event.ctrlKey;
      if (modifier && event.key.toLowerCase() === "s") { event.preventDefault(); void saveFile(event.shiftKey); }
      if (modifier && event.key.toLowerCase() === "o") { event.preventDefault(); void openFile(); }
      if (modifier && event.key.toLowerCase() === "n") { event.preventDefault(); newWorkbook(); }
      if (modifier && event.key.toLowerCase() === "f") { event.preventDefault(); setFindOpen(true); }
      if (event.key === "Escape") { setFindOpen(false); setChartOpen(false); setShowFileMenu(false); }
    };
    window.addEventListener("keydown", keydown); return () => window.removeEventListener("keydown", keydown);
  }, [saveFile, openFile, newWorkbook]);

  const findMatches = useMemo(() => {
    if (!findQuery.trim()) return new Set<string>();
    const query = findQuery.toLowerCase();
    return new Set(Object.entries(activeSheet.cells).filter(([address, cell]) => displayValue(evaluateCell(workbook, activeSheet, address), cell.style).toLowerCase().includes(query)).map(([address]) => address));
  }, [findQuery, activeSheet, workbook]);

  const selectedStats = useMemo(() => {
    const values = selectionAddresses(selection).map((address) => evaluateCell(workbook, activeSheet, address)).filter((value): value is number => typeof value === "number");
    return { count: selectionAddresses(selection).length, numeric: values.length, sum: values.reduce((a, b) => a + b, 0), average: values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0 };
  }, [selection, workbook, activeSheet]);

  const chartData = useMemo(() => {
    const n = normalizedSelection(selection); const result: { label: string; value: number }[] = [];
    for (let row = n.start.row; row <= n.end.row; row++) for (let col = n.start.col; col <= n.end.col; col++) {
      const value = evaluateCell(workbook, activeSheet, addressOf(row, col)); if (typeof value === "number") result.push({ label: addressOf(row, col), value });
    }
    return result.slice(0, 30);
  }, [selection, workbook, activeSheet]);

  const recentFiles = useMemo(() => { try { return JSON.parse(localStorage.getItem("gridstone-recent") || "[]") as FileState[]; } catch { return []; } }, [showFileMenu]);
  const currentStyle = selectedCell?.style || {};
  const formulaValue = String(selectedCell?.raw ?? "");

  return (
    <div className="app-shell" onMouseUp={() => setDragging(false)} onDragOver={(event) => event.preventDefault()} onDrop={(event) => {
      event.preventDefault(); const dropped = event.dataTransfer.files[0]; if (!dropped) return;
      void dropped.arrayBuffer().then((buffer) => loadDesktopFile({ path: (dropped as File & { path?: string }).path || "", name: dropped.name, extension: dropped.name.split(".").pop()?.toLowerCase() || "xlsx", data: new Uint8Array(buffer) }));
    }}>
      <a className="skip-link" href="#spreadsheet-grid">Skip to spreadsheet</a>
      <header className="titlebar">
        <div className="document-title">
          <div><strong>{file.name.replace(/\.[^.]+$/, "")}</strong>{dirty && <span className="dirty-dot" title="Unsaved changes">•</span>}</div>
          <span>{dirty ? "Unsaved changes" : "Saved locally"}</span>
        </div>
        <div className="title-actions">
          <button className="share-button" type="button" onClick={() => setStatus("Cloud collaboration can be connected from Settings")}>Share</button>
          <IconButton label="Open inspector" active={inspectorOpen} onClick={() => setInspectorOpen(!inspectorOpen)}><PanelRight size={18} /></IconButton>
        </div>
      </header>

      <nav className="ribbon-tabs" aria-label="Workbook commands">
        {(["Home", "Insert", "Data", "View"] as RibbonTab[]).map((tab) => <button className={`ribbon-tab ${ribbon === tab ? "active" : ""}`} type="button" key={tab} onClick={() => setRibbon(tab)}>{tab}</button>)}
      </nav>

      <section className="toolbar" aria-label={`${ribbon} toolbar`}>
        <div className="tool-group compact file-menu-wrap">
          <button ref={fileBtnRef} className={`tool-button file-tab ${showFileMenu ? "is-active" : ""}`} type="button" onClick={() => { const rect = fileBtnRef.current?.getBoundingClientRect(); if (rect) setFileMenuPos({ top: rect.bottom, left: rect.left }); setShowFileMenu((open) => !open); }}><FolderOpen size={18} /><span>File</span><ChevronDown size={14} /></button>
        </div>
        <div className="tool-group compact">
          <IconButton label="Undo" disabled={!undoStack.current.length} onClick={undo}><Undo2 size={18} /></IconButton>
          <IconButton label="Redo" disabled={!redoStack.current.length} onClick={redo}><Redo2 size={18} /></IconButton>
        </div>
        {ribbon === "Home" && <>
          <div className="tool-group">
            <ToolButton label="Paste" icon={<Download size={18} />} onClick={() => void pasteSelection()} />
            <IconButton label="Cut" onClick={() => void copySelection(true)}><Scissors size={17} /></IconButton>
            <IconButton label="Copy" onClick={() => void copySelection()}><Copy size={17} /></IconButton>
          </div>
          <div className="tool-group">
            <select aria-label="Font family" defaultValue="Aptos"><option>Aptos</option><option>Georgia</option><option>JetBrains Mono</option></select>
            <select aria-label="Font size" defaultValue="14"><option>11</option><option>12</option><option>14</option><option>16</option><option>18</option><option>24</option></select>
            <IconButton label="Bold" active={currentStyle.bold} onClick={() => applyStyle({ bold: !currentStyle.bold })}><Bold size={17} /></IconButton>
            <IconButton label="Italic" active={currentStyle.italic} onClick={() => applyStyle({ italic: !currentStyle.italic })}><Italic size={17} /></IconButton>
            <IconButton label="Underline" active={currentStyle.underline} onClick={() => applyStyle({ underline: !currentStyle.underline })}><Underline size={17} /></IconButton>
            <ColorControl label="Text color" value={currentStyle.textColor || "#17221f"} icon={<span className="text-color-icon">A</span>} onChange={(value) => applyStyle({ textColor: value })} />
            <ColorControl label="Fill color" value={currentStyle.fillColor || "#ffffff"} icon={<PaintBucket size={17} />} onChange={(value) => applyStyle({ fillColor: value })} />
          </div>
          <div className="tool-group compact">
            <IconButton label="Align left" active={currentStyle.align === "left"} onClick={() => applyStyle({ align: "left" })}><AlignLeft size={17} /></IconButton>
            <IconButton label="Align center" active={currentStyle.align === "center"} onClick={() => applyStyle({ align: "center" })}><AlignCenter size={17} /></IconButton>
            <IconButton label="Align right" active={currentStyle.align === "right"} onClick={() => applyStyle({ align: "right" })}><AlignRight size={17} /></IconButton>
            <IconButton label="Merge cells" active={mergesOverlapping(activeSheet.merges || [], selection).length > 0} onClick={toggleMerge}><Combine size={17} /></IconButton>
          </div>
          <div className="tool-group compact">
            <IconButton label="Currency format" active={currentStyle.format === "currency"} onClick={() => applyStyle({ format: "currency" })}><DollarSign size={17} /></IconButton>
            <IconButton label="Percent format" active={currentStyle.format === "percent"} onClick={() => applyStyle({ format: "percent" })}><Percent size={17} /></IconButton>
            <select aria-label="Number format" value={currentStyle.format || "general"} onChange={(event) => applyStyle({ format: event.target.value as CellStyle["format"] })}><option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percent</option></select>
          </div>
        </>}
        {ribbon === "Insert" && <>
          <div className="tool-group"><ToolButton label="Chart" icon={<BarChart3 size={19} />} onClick={() => setChartOpen(true)} /><ToolButton label="New sheet" icon={<Plus size={19} />} onClick={addSheet} /></div>
          <div className="toolbar-hint">Select numeric cells, then insert a chart.</div>
        </>}
        {ribbon === "Data" && <>
          <div className="tool-group"><ToolButton label="Sort A → Z" icon={<ArrowUpNarrowWide size={19} />} onClick={() => sortSelection("asc")} /><ToolButton label="Sort Z → A" icon={<ArrowDownWideNarrow size={19} />} onClick={() => sortSelection("desc")} /><ToolButton label="Remove duplicates" icon={<Trash2 size={19} />} onClick={removeDuplicates} /></div>
          <div className="toolbar-hint">Data tools apply to the current selection.</div>
        </>}
        {ribbon === "View" && <>
          <div className="tool-group">
            <ToolButton label={activeSheet.frozenRows ? "Unfreeze rows" : "Freeze top row"} icon={<Snowflake size={19} />} onClick={() => updateActiveSheet((sheet) => ({ ...sheet, frozenRows: sheet.frozenRows ? 0 : 1 }), "Updated frozen rows")} />
            <ToolButton label={activeSheet.frozenColumns ? "Unfreeze columns" : "Freeze first column"} icon={<Columns size={19} />} onClick={() => updateActiveSheet((sheet) => ({ ...sheet, frozenColumns: sheet.frozenColumns ? 0 : 1 }), "Updated frozen columns")} />
            <ToolButton label={theme === "light" ? "Dark theme" : "Light theme"} icon={theme === "light" ? <Moon size={19} /> : <Sun size={19} />} onClick={() => setTheme(theme === "light" ? "dark" : "light")} />
          </div>
          <div className="zoom-toolbar"><button onClick={() => setZoom(Math.max(50, zoom - 10))}>−</button><span>{zoom}%</span><button onClick={() => setZoom(Math.min(200, zoom + 10))}>+</button></div>
        </>}
      </section>

      {showFileMenu && <>
        <div className="file-menu-scrim" onMouseDown={() => setShowFileMenu(false)} />
        <div className="file-menu" role="menu" style={{ position: "fixed", top: fileMenuPos.top, left: fileMenuPos.left }}>
          <button role="menuitem" onClick={newWorkbook}><FilePlus2 />New workbook<span>⌘N</span></button>
          <button role="menuitem" onClick={() => void openFile()}><FolderOpen />Open…<span>⌘O</span></button>
          <button role="menuitem" onClick={() => void saveFile()}><Save />Save<span>⌘S</span></button>
          <button role="menuitem" onClick={() => void saveFile(true)}><Download />Save as…<span>⇧⌘S</span></button>
          <button role="menuitem" onClick={() => void saveFile(true, "csv")}><Download />Export CSV</button>
          <button role="menuitem" onClick={() => void printWorkbook()}><Printer />Print / PDF<span>⌘P</span></button>
          {recentFiles.length > 0 && <div className="menu-heading">Recent files</div>}
          {recentFiles.map((item) => <button role="menuitem" key={item.path} onClick={() => void openRecent(item)}><span className="file-badge">{item.extension}</span>{item.name}</button>)}
        </div>
      </>}

      <section className="formula-row" aria-label="Formula bar">
        <div className="name-box">{selectedAddress}<ChevronDown size={14} /></div>
        <div className="formula-symbol">ƒx</div>
        <FormulaEditor key={selectedAddress} address={selectedAddress} initialValue={formulaValue} onCommit={commitEdit} onFocusGrid={() => gridRef.current?.focus()} />
      </section>

      {findOpen && <div className="find-bar" role="search"><Search size={17} /><input autoFocus placeholder="Find in sheet" value={findQuery} onChange={(event) => setFindQuery(event.target.value)} /><span>{findMatches.size} result{findMatches.size === 1 ? "" : "s"}</span><button aria-label="Close find" onClick={() => setFindOpen(false)}><X size={17} /></button></div>}

      <main className="workspace">
        <div className="grid-viewport" id="spreadsheet-grid" ref={gridRef} role="grid" aria-label={`${activeSheet.name} spreadsheet`} tabIndex={0} onKeyDown={handleGridKeyDown} style={{ fontSize: `${zoom}%` }}>
          <div className="sheet-grid" style={{ gridTemplateColumns: `48px ${Array.from({ length: dimensions.cols }, (_, col) => `${activeSheet.columnWidths[col] || 112}px`).join(" ")}` }}>
            <div className="corner-cell" role="button" tabIndex={-1} aria-label="Select all cells" onMouseDown={selectAll} />
            {Array.from({ length: dimensions.cols }, (_, col) => {
              const frozen = col < (activeSheet.frozenColumns || 0);
              return <div
                className={`column-header ${col >= normalizedSelection(selection).start.col && col <= normalizedSelection(selection).end.col ? "header-selected" : ""} ${frozen ? "frozen" : ""}`}
                role="columnheader" key={`h-${col}`}
                style={frozen ? { left: colOffsets[col], zIndex: 25 } : undefined}
                onMouseDown={(event) => selectColumn(col, event.shiftKey)}
                onMouseEnter={() => { if (dragging && dragKind.current === "col") setSelection({ start: { row: 0, col: Math.min(headerAnchor.current, col) }, end: { row: dimensions.rows - 1, col: Math.max(headerAnchor.current, col) } }); }}
              >{columnName(col)}</div>;
            })}
            {Array.from({ length: dimensions.rows }, (_, row) => {
              const rowFrozen = row < (activeSheet.frozenRows || 0);
              return <div className="grid-row" role="row" key={`row-${row}`} style={{ display: "contents" }}>
              <div
                className={`row-header ${row >= normalizedSelection(selection).start.row && row <= normalizedSelection(selection).end.row ? "header-selected" : ""} ${rowFrozen ? "frozen" : ""}`}
                role="rowheader"
                style={rowFrozen ? { top: 29 + row * 29, zIndex: 25 } : undefined}
                onMouseDown={(event) => selectRow(row, event.shiftKey)}
                onMouseEnter={() => { if (dragging && dragKind.current === "row") setSelection({ start: { row: Math.min(headerAnchor.current, row), col: 0 }, end: { row: Math.max(headerAnchor.current, row), col: dimensions.cols - 1 } }); }}
              >{row + 1}</div>
              {Array.from({ length: dimensions.cols }, (_, col) => {
                if (mergeMap.covered.has(`${row},${col}`)) return null;
                const region = mergeMap.anchors.get(`${row},${col}`);
                const address = addressOf(row, col); const cell = activeSheet.cells[address];
                const n = normalizedSelection(selection); const selected = row >= n.start.row && row <= n.end.row && col >= n.start.col && col <= n.end.col;
                const active = address === selectedAddress; const value = displayValue(evaluateCell(workbook, activeSheet, address), cell?.style);
                const style = cell?.style;
                const colFrozen = col < (activeSheet.frozenColumns || 0);
                const frozenStyle: React.CSSProperties = {};
                if (colFrozen) { frozenStyle.position = "sticky"; frozenStyle.left = colOffsets[col]; frozenStyle.zIndex = 11; }
                if (rowFrozen) { frozenStyle.position = "sticky"; frozenStyle.top = 29 + row * 29; frozenStyle.zIndex = colFrozen ? 13 : 12; }
                return <div
                  className={`grid-cell ${selected ? "selected" : ""} ${active ? "active" : ""} ${findMatches.has(address) ? "find-match" : ""}`}
                  role="gridcell" aria-selected={selected} aria-label={`${address}, ${value || "blank"}`} key={address}
                  style={{ gridColumn: `${col + 2} / span ${region?.cols ?? 1}`, gridRow: `${row + 2} / span ${region?.rows ?? 1}`, fontWeight: style?.bold ? 700 : undefined, fontStyle: style?.italic ? "italic" : undefined, textDecoration: style?.underline ? "underline" : undefined, color: style?.textColor, backgroundColor: style?.fillColor, textAlign: style?.align, whiteSpace: style?.wrap ? "normal" : undefined, ...frozenStyle }}
                  onMouseDown={(event) => { dragKind.current = "cell"; setDragging(true); const point = { row, col }; setSelection(event.shiftKey ? { ...selection, end: point } : { start: point, end: point }); gridRef.current?.focus(); }}
                  onMouseEnter={() => { if (dragging && dragKind.current === "cell") setSelection((current) => ({ ...current, end: { row, col } })); }}
                  onDoubleClick={() => beginEdit(address)}
                  title={cell?.note || (typeof cell?.raw === "string" && cell.raw.startsWith("=") ? cell.raw : undefined)}
                >
                  {editing?.address === address ? <CellEditor initialValue={editing.initialValue} onCommit={(nextValue) => commitEdit(address, nextValue)} onCancel={() => { setEditing(null); gridRef.current?.focus(); }} /> : <span>{value}</span>}
                </div>;
              })}
            </div>;
            })}
          </div>
        </div>

        {inspectorOpen && <aside className="inspector" aria-label="Cell inspector">
          <div className="inspector-heading"><div><span>Cell inspector</span><strong>{selectedAddress}</strong></div><button aria-label="Close inspector" onClick={() => setInspectorOpen(false)}><X size={18} /></button></div>
          <label>Cell note<textarea value={selectedCell?.note || ""} placeholder="Add context or instructions…" onChange={(event) => { const note = event.target.value; updateActiveSheet((sheet) => ({ ...sheet, cells: { ...sheet.cells, [selectedAddress]: { ...(sheet.cells[selectedAddress] || { raw: null }), note } } }), "Updated note"); }} /></label>
          <div className="inspector-section"><span>Computed value</span><strong>{displayValue(evaluateCell(workbook, activeSheet, selectedAddress), selectedCell?.style) || "—"}</strong></div>
          <div className="inspector-section"><span>Raw value</span><code>{String(selectedCell?.raw ?? "—")}</code></div>
          <button className="danger-button" onClick={clearSelection}><Trash2 size={16} />Clear selected cells</button>
        </aside>}
      </main>

      <footer className="bottom-bar">
        <div className="sheet-controls">
          <button className="add-sheet" type="button" aria-label="Add worksheet" onClick={addSheet}><Plus size={17} /></button>
          {workbook.sheets.map((sheet) => <div className={`sheet-tab-wrap ${sheet.id === workbook.activeSheetId ? "active" : ""}`} key={sheet.id}><button className="sheet-tab" type="button" onClick={() => { setWorkbook((current) => ({ ...current, activeSheetId: sheet.id })); setSelection(INITIAL_SELECTION); }} onDoubleClick={() => renameSheet(sheet.id)}>{sheet.name}</button>{sheet.id === workbook.activeSheetId && workbook.sheets.length > 1 && <button className="sheet-delete" aria-label={`Delete ${sheet.name}`} onClick={() => deleteSheet(sheet.id)}><X size={13} /></button>}</div>)}
        </div>
        <div className="status-message" aria-live="polite">{status}</div>
        <div className="selection-stats"><span>Count <strong>{selectedStats.count}</strong></span>{selectedStats.numeric > 0 && <><span>Average <strong>{selectedStats.average.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span><span>Sum <strong>{selectedStats.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span></>}</div>
      </footer>

      {chartOpen && <div className="modal-backdrop" role="presentation" onMouseDown={() => setChartOpen(false)}><section className="modal chart-modal" role="dialog" aria-modal="true" aria-labelledby="chart-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="modal-heading"><div><span>Insert</span><h2 id="chart-title">Column chart</h2></div><button aria-label="Close chart" onClick={() => setChartOpen(false)}><X size={20} /></button></div>
        {chartData.length ? <div className="chart-area" aria-label="Preview of selected data">{chartData.map((item) => { const max = Math.max(...chartData.map((entry) => Math.abs(entry.value)), 1); return <div className="chart-column" key={item.label}><div className="bar-value">{item.value.toLocaleString()}</div><div className="bar" style={{ height: `${Math.max(4, Math.abs(item.value) / max * 210)}px` }} /><span>{item.label}</span></div>; })}</div> : <div className="empty-chart"><BarChart3 size={40} /><h3>No numeric cells selected</h3><p>Select a range containing numbers and try again.</p></div>}
        <div className="modal-footer"><button className="secondary-button" onClick={() => setChartOpen(false)}>Cancel</button><button className="primary-button" disabled={!chartData.length} onClick={() => { setChartOpen(false); setStatus("Chart added to workbook view"); }}>Add chart</button></div>
      </section></div>}

      <input ref={fileInputRef} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => { const selected = event.target.files?.[0]; if (!selected) return; void selected.arrayBuffer().then((buffer) => loadDesktopFile({ path: "", name: selected.name, extension: selected.name.split(".").pop()?.toLowerCase() || "xlsx", data: new Uint8Array(buffer) })); }} />
    </div>
  );
}
