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

## Automatic updates

The app checks GitHub Releases on launch (and from **File → Check for updates…** or the app/Help menu) and downloads new versions in the background using [`electron-updater`](https://www.electron.build/auto-update). When a build finishes downloading, an in-app banner offers **Restart & install**. Updates are wired to the release feed declared in `build.publish` in `package.json` (`akilxasp/gridstone-lite`).

Publish a release that installed apps will pick up:

```bash
export GH_TOKEN=<a GitHub token with repo scope>
npm run release
```

`npm run release` builds the renderer and runs `electron-builder --publish always`, which uploads the installers plus the `latest.yml` / `latest-mac.yml` update metadata (and `.blockmap` files for delta downloads) to a GitHub Release for the current `version` in `package.json`. Bump `version` before each release. Run it per target OS (macOS for `dmg`/`zip`, Windows for `nsis`) so every platform's metadata lands on the same release.

Update mechanism per platform:

- **Windows:** full silent auto-update via `electron-updater` (Squirrel). New versions download in the background; the in-app banner offers **Restart & install**. Unsigned installers work but show SmartScreen warnings, so a code-signing certificate is recommended. The `portable` target does not auto-update.
- **macOS:** Squirrel.Mac requires a signed, notarized build, which needs a paid Apple Developer ID. To avoid that, macOS uses a **manual download-and-open** flow instead: the app queries the latest GitHub release, and if it's newer, the banner offers **Download update**. It downloads the matching `.dmg` to the Downloads folder and opens it; the user drags Gridstone into Applications and relaunches. No signing or Apple account required. If you later obtain a Developer ID, sign + notarize the build and switch macOS back to the Squirrel path in `electron/main.cjs` for fully silent updates.

## Excel compatibility

Gridstone is a working spreadsheet application, not an Excel clone. It does not support VBA/macros, Power Query, pivot tables, advanced chart editing, real-time collaboration, password-protected workbooks, every Excel function, or enterprise cloud connectors.
