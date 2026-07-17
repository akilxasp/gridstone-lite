# Gridstone

Gridstone is a local-first spreadsheet desktop application for macOS and Windows. It uses Electron, React, TypeScript, and SheetJS.

## Included in this build

- Native open, save, save-as, print, recent-file, drag-and-drop, and file-association workflows
- `.xlsx`, `.xls`, `.csv`, and `.tsv` import; `.xlsx` and `.csv` export
- Multiple worksheets with add, rename, switch, and delete controls
- Direct cell and formula-bar editing
- Keyboard navigation, range selection, copy, cut, paste, clear, undo, and redo
- Common formulas, arithmetic, comparisons, cell/range references, and circular-reference detection
- Text, fill, emphasis, alignment, number, currency, and percentage formatting
- Sort, duplicate removal, find, charts, notes, print/PDF, dark mode, zoom, frozen rows, and selection statistics
- Local recovery copies, unsaved-change protection, and offline operation
- Windows NSIS/portable and macOS DMG/ZIP packaging configuration

## Development

```bash
npm install
npm run dev
```

Run verification:

```bash
npm test
npm run build
```

Create installers:

```bash
npm run dist:mac
npm run dist:mac:universal
npm run dist:win
```

`dist:mac:universal` creates one macOS installer for Apple Silicon (including M3 Pro) and Intel x64 Macs (including Intel i7) running macOS 11 or later. The local universal build uses an ad-hoc signature for internal testing. Public distribution still requires an Apple Developer ID signature and notarization.

Windows artifacts should be produced on Windows or in a Windows CI runner. Public distribution also requires a Windows code-signing certificate.

## Excel compatibility boundary

This is a working spreadsheet application, not a complete clone of Microsoft Excel. Full parity requires separate engines and substantial additional product work for VBA/macros, Power Query, pivot tables, advanced chart editing, real-time collaboration, password-protected workbooks, every Excel function, perfect preservation of unsupported workbook objects, and enterprise cloud connectors.
