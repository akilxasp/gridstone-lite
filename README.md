# Gridstone

A local-first spreadsheet desktop application for macOS and Windows, built with Electron, React, TypeScript, and SheetJS.

## Features

- Native open, save, save-as, print, recent-file, drag-and-drop, and file-association workflows
- `.xlsx`, `.xls`, `.csv`, and `.tsv` import; `.xlsx` and `.csv` export
- Multiple worksheets with add, rename, switch, and delete controls
- Direct cell and formula-bar editing
- Keyboard navigation, range selection, copy, cut, paste, clear, undo, and redo
- Formulas, arithmetic, comparisons, cell/range references, and circular-reference detection
- Text, fill, emphasis, alignment, number, currency, and percentage formatting
- Sort, duplicate removal, find, charts, notes, print/PDF, dark mode, zoom, frozen rows, and selection statistics
- Row, column, and select-all header selection with click-drag ranges
- Freeze top row and/or first column, and merge/unmerge cell ranges
- Automatic in-app updates from GitHub Releases (download and install without re-downloading manually)
- Local recovery copies, unsaved-change protection, and offline operation
- Windows NSIS/portable and macOS DMG/ZIP packaging

## Development

```bash
npm install
npm run dev
```

Run tests and build:

```bash
npm test
npm run build
```

## Installers

```bash
npm run dist:mac            # Apple Silicon DMG
npm run dist:mac:x64        # Intel x64 DMG
npm run dist:mac:universal  # Universal DMG
npm run dist:win            # Windows NSIS/portable
```

The universal build supports Apple Silicon and Intel x64 Macs on macOS 11 or later, using an ad-hoc signature for internal testing. Public distribution requires an Apple Developer ID signature and notarization.

Build Windows artifacts on Windows or a Windows CI runner. Public distribution requires a Windows code-signing certificate.
