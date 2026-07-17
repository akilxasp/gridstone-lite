import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlignCenter, AlignLeft, AlignRight, ArrowDownWideNarrow, ArrowUpNarrowWide,
  Bold, Check, ChevronDown, Copy, DollarSign, Download, FilePlus2,
  FolderOpen, Italic, Moon, PaintBucket, Percent, Plus,
  Redo2, Save, Scissors, Search, Snowflake, Sun, Trash2, Underline, Undo2, X,
} from "lucide-react";
import {
  CellData, CellStyle, DEFAULT_COLS, DEFAULT_ROWS, Selection, SheetData, WorkbookData,
  addressOf, columnName, createSheet, displayValue, evaluateCell, exportWorkbook,
  importWorkbook, normalizedSelection, pointFromAddress, sampleWorkbook, selectionAddresses,
  setCellInput,
} from "./workbook";

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
  const [editing, setEditing] = useState<EditSession | null>(null);
  const [dragging, setDragging] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => (localStorage.getItem("gridstone-theme") as "light" | "dark") || "light");
  const [status, setStatus] = useState("Ready");
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState("");
  const [showFileMenu, setShowFileMenu] = useState(false);
  const [zoom, setZoom] = useState(100);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const undoStack = useRef<WorkbookData[]>([]);
  const redoStack = useRef<WorkbookData[]>([]);
  const workbookRef = useRef(workbook);
  const fileRef = useRef(file);

  useEffect(() => { workbookRef.current = workbook; }, [workbook]);
  useEffect(() => { fileRef.current = file; }, [file]);

  const activeSheet = useMemo(() => workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) || workbook.sheets[0], [workbook]);
  const selectedAddress = addressOf(selection.end.row, selection.end.col);
  const selectedCell = activeSheet.cells[selectedAddress];

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
      else if (command === "save-as") void saveFile(true); else if (command === "find") setFindOpen(true);
      else if (command === "open-path" && typeof payload === "string") void window.desktop?.readPath(payload).then(loadDesktopFile).catch(() => setStatus("Could not open the selected workbook"));
    });
  }, [newWorkbook, openFile, saveFile, loadDesktopFile]);

  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      const modifier = event.metaKey || event.ctrlKey;
      if (modifier && event.key.toLowerCase() === "s") { event.preventDefault(); void saveFile(event.shiftKey); }
      if (modifier && event.key.toLowerCase() === "o") { event.preventDefault(); void openFile(); }
      if (modifier && event.key.toLowerCase() === "n") { event.preventDefault(); newWorkbook(); }
      if (modifier && event.key.toLowerCase() === "f") { event.preventDefault(); setFindOpen(true); }
      if (event.key === "Escape") { setFindOpen(false); setShowFileMenu(false); }
    };
    window.addEventListener("keydown", keydown); return () => window.removeEventListener("keydown", keydown);
  }, [saveFile, openFile, newWorkbook]);

  const dimensions = useMemo(() => {
    let rows = DEFAULT_ROWS, cols = DEFAULT_COLS;
    Object.keys(activeSheet.cells).forEach((address) => { const point = pointFromAddress(address); if (point) { rows = Math.max(rows, point.row + 20); cols = Math.max(cols, point.col + 8); } });
    return { rows: Math.min(rows, 500), cols: Math.min(cols, 100) };
  }, [activeSheet.cells]);

  const findMatches = useMemo(() => {
    if (!findQuery.trim()) return new Set<string>();
    const query = findQuery.toLowerCase();
    return new Set(Object.entries(activeSheet.cells).filter(([address, cell]) => displayValue(evaluateCell(workbook, activeSheet, address), cell.style).toLowerCase().includes(query)).map(([address]) => address));
  }, [findQuery, activeSheet, workbook]);

  const selectedStats = useMemo(() => {
    const values = selectionAddresses(selection).map((address) => evaluateCell(workbook, activeSheet, address)).filter((value): value is number => typeof value === "number");
    return { count: selectionAddresses(selection).length, numeric: values.length, sum: values.reduce((a, b) => a + b, 0), average: values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0 };
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
        <div className="brand-mark" aria-hidden="true"><span>G</span></div>
        <div className="document-title">
          <div><strong>{file.name.replace(/\.[^.]+$/, "")}</strong>{dirty && <span className="dirty-dot" title="Unsaved changes">•</span>}</div>
          <span>{dirty ? "Unsaved changes" : "Saved locally"}</span>
        </div>
      </header>

      <nav className="ribbon-tabs" aria-label="Workbook commands">
        <div className="file-menu-wrap">
          <button className={`ribbon-tab file-tab ${showFileMenu ? "active" : ""}`} type="button" onClick={() => setShowFileMenu(!showFileMenu)}>File</button>
          {showFileMenu && <div className="file-menu" role="menu">
            <button role="menuitem" onClick={newWorkbook}><FilePlus2 />New workbook<span>⌘N</span></button>
            <button role="menuitem" onClick={() => void openFile()}><FolderOpen />Open…<span>⌘O</span></button>
            <button role="menuitem" onClick={() => void saveFile()}><Save />Save<span>⌘S</span></button>
            <button role="menuitem" onClick={() => void saveFile(true)}><Download />Save as…<span>⇧⌘S</span></button>
            <button role="menuitem" onClick={() => void saveFile(true, "csv")}><Download />Export CSV</button>
            {recentFiles.length > 0 && <div className="menu-heading">Recent files</div>}
            {recentFiles.map((item) => <button role="menuitem" key={item.path} onClick={() => void openRecent(item)}><span className="file-badge">{item.extension}</span>{item.name}</button>)}
          </div>}
        </div>
      </nav>

      <section className="toolbar" aria-label="Toolbar">
        <div className="tool-group compact">
          <IconButton label="Undo" disabled={!undoStack.current.length} onClick={undo}><Undo2 size={18} /></IconButton>
          <IconButton label="Redo" disabled={!redoStack.current.length} onClick={redo}><Redo2 size={18} /></IconButton>
        </div>
        <div className="tool-group">
          <ToolButton label="Paste" icon={<Download size={18} />} onClick={() => void pasteSelection()} />
          <IconButton label="Cut" onClick={() => void copySelection(true)}><Scissors size={17} /></IconButton>
          <IconButton label="Copy" onClick={() => void copySelection()}><Copy size={17} /></IconButton>
        </div>
        <div className="tool-group">
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
        </div>
        <div className="tool-group compact">
          <IconButton label="Currency format" active={currentStyle.format === "currency"} onClick={() => applyStyle({ format: "currency" })}><DollarSign size={17} /></IconButton>
          <IconButton label="Percent format" active={currentStyle.format === "percent"} onClick={() => applyStyle({ format: "percent" })}><Percent size={17} /></IconButton>
          <select aria-label="Number format" value={currentStyle.format || "general"} onChange={(event) => applyStyle({ format: event.target.value as CellStyle["format"] })}><option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percent</option><option value="date">Date</option></select>
        </div>
        <div className="tool-group compact">
          <IconButton label="Sort ascending" onClick={() => sortSelection("asc")}><ArrowUpNarrowWide size={17} /></IconButton>
          <IconButton label="Sort descending" onClick={() => sortSelection("desc")}><ArrowDownWideNarrow size={17} /></IconButton>
          <IconButton label="Remove duplicates" onClick={removeDuplicates}><Trash2 size={17} /></IconButton>
        </div>
        <div className="tool-group compact">
          <IconButton label="Freeze top row" active={Boolean(activeSheet.frozenRows)} onClick={() => updateActiveSheet((sheet) => ({ ...sheet, frozenRows: sheet.frozenRows ? 0 : 1 }), "Updated frozen rows")}><Snowflake size={17} /></IconButton>
          <IconButton label={theme === "light" ? "Dark theme" : "Light theme"} onClick={() => setTheme(theme === "light" ? "dark" : "light")}>{theme === "light" ? <Moon size={17} /> : <Sun size={17} />}</IconButton>
          <div className="zoom-toolbar"><button onClick={() => setZoom(Math.max(50, zoom - 10))}>−</button><span>{zoom}%</span><button onClick={() => setZoom(Math.min(200, zoom + 10))}>+</button></div>
        </div>
      </section>

      <section className="formula-row" aria-label="Formula bar">
        <div className="name-box">{selectedAddress}<ChevronDown size={14} /></div>
        <div className="formula-symbol">ƒx</div>
        <FormulaEditor key={selectedAddress} address={selectedAddress} initialValue={formulaValue} onCommit={commitEdit} onFocusGrid={() => gridRef.current?.focus()} />
      </section>

      {findOpen && <div className="find-bar" role="search"><Search size={17} /><input autoFocus placeholder="Find in sheet" value={findQuery} onChange={(event) => setFindQuery(event.target.value)} /><span>{findMatches.size} result{findMatches.size === 1 ? "" : "s"}</span><button aria-label="Close find" onClick={() => setFindOpen(false)}><X size={17} /></button></div>}

      <main className="workspace">
        <div className="grid-viewport" id="spreadsheet-grid" ref={gridRef} role="grid" aria-label={`${activeSheet.name} spreadsheet`} tabIndex={0} onKeyDown={handleGridKeyDown} style={{ fontSize: `${zoom}%` }}>
          <div className="sheet-grid" style={{ gridTemplateColumns: `48px ${Array.from({ length: dimensions.cols }, (_, col) => `${activeSheet.columnWidths[col] || 112}px`).join(" ")}` }}>
            <div className="corner-cell" aria-hidden="true" />
            {Array.from({ length: dimensions.cols }, (_, col) => <div className="column-header" role="columnheader" key={`h-${col}`}>{columnName(col)}</div>)}
            {Array.from({ length: dimensions.rows }, (_, row) => <div className="grid-row" role="row" key={`row-${row}`} style={{ display: "contents" }}>
              <div className="row-header" role="rowheader">{row + 1}</div>
              {Array.from({ length: dimensions.cols }, (_, col) => {
                const address = addressOf(row, col); const cell = activeSheet.cells[address];
                const n = normalizedSelection(selection); const selected = row >= n.start.row && row <= n.end.row && col >= n.start.col && col <= n.end.col;
                const active = address === selectedAddress; const value = displayValue(evaluateCell(workbook, activeSheet, address), cell?.style);
                const style = cell?.style;
                return <div
                  className={`grid-cell ${selected ? "selected" : ""} ${active ? "active" : ""} ${findMatches.has(address) ? "find-match" : ""}`}
                  role="gridcell" aria-selected={selected} aria-label={`${address}, ${value || "blank"}`} key={address}
                  style={{ fontWeight: style?.bold ? 700 : undefined, fontStyle: style?.italic ? "italic" : undefined, textDecoration: style?.underline ? "underline" : undefined, color: style?.textColor, backgroundColor: style?.fillColor, textAlign: style?.align, whiteSpace: style?.wrap ? "normal" : undefined }}
                  onMouseDown={(event) => { setDragging(true); const point = { row, col }; setSelection(event.shiftKey ? { ...selection, end: point } : { start: point, end: point }); gridRef.current?.focus(); }}
                  onMouseEnter={() => { if (dragging) setSelection((current) => ({ ...current, end: { row, col } })); }}
                  onDoubleClick={() => beginEdit(address)}
                  title={typeof cell?.raw === "string" && cell.raw.startsWith("=") ? cell.raw : undefined}
                >
                  {editing?.address === address ? <CellEditor initialValue={editing.initialValue} onCommit={(nextValue) => commitEdit(address, nextValue)} onCancel={() => { setEditing(null); gridRef.current?.focus(); }} /> : <span>{value}</span>}
                </div>;
              })}
            </div>)}
          </div>
        </div>

      </main>

      <footer className="bottom-bar">
        <div className="sheet-controls">
          <button className="add-sheet" type="button" aria-label="Add worksheet" onClick={addSheet}><Plus size={17} /></button>
          {workbook.sheets.map((sheet) => <div className={`sheet-tab-wrap ${sheet.id === workbook.activeSheetId ? "active" : ""}`} key={sheet.id}><button className="sheet-tab" type="button" onClick={() => { setWorkbook((current) => ({ ...current, activeSheetId: sheet.id })); setSelection(INITIAL_SELECTION); }} onDoubleClick={() => renameSheet(sheet.id)}>{sheet.name}</button>{sheet.id === workbook.activeSheetId && workbook.sheets.length > 1 && <button className="sheet-delete" aria-label={`Delete ${sheet.name}`} onClick={() => deleteSheet(sheet.id)}><X size={13} /></button>}</div>)}
        </div>
        <div className="status-message" aria-live="polite">{status}</div>
        <div className="selection-stats"><span>Count <strong>{selectedStats.count}</strong></span>{selectedStats.numeric > 0 && <><span>Average <strong>{selectedStats.average.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span><span>Sum <strong>{selectedStats.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span></>}</div>
      </footer>

      <input ref={fileInputRef} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => { const selected = event.target.files?.[0]; if (!selected) return; void selected.arrayBuffer().then((buffer) => loadDesktopFile({ path: "", name: selected.name, extension: selected.name.split(".").pop()?.toLowerCase() || "xlsx", data: new Uint8Array(buffer) })); }} />
    </div>
  );
}
