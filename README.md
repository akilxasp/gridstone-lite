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
- Row, column, and select-all header selection with click-drag ranges
- Freeze top row and/or first column, and merge/unmerge cell ranges
- Automatic in-app updates from GitHub Releases (download and install without re-downloading manually)
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

## Automatic updates

The app checks GitHub Releases on launch (and from **File → Check for updates…** or the app/Help menu) and downloads new versions in the background using [`electron-updater`](https://www.electron.build/auto-update). When a build finishes downloading, an in-app banner offers **Restart & install**. Updates are wired to the release feed declared in `build.publish` in `package.json` (`akilxasp/gridstone-lite`).

Publish a release that installed apps will pick up:

```bash
export GH_TOKEN=<a GitHub token with repo scope>
npm run release
```

`npm run release` builds the renderer and runs `electron-builder --publish always`, which uploads the installers plus the `latest.yml` / `latest-mac.yml` update metadata (and `.blockmap` files for delta downloads) to a GitHub Release for the current `version` in `package.json`. Bump `version` before each release. Run it per target OS (macOS for `dmg`/`zip`, Windows for `nsis`) so every platform's metadata lands on the same release.

Signing requirements for auto-update:

- **macOS:** Squirrel.Mac only installs updates from a signed, notarized build, and updates are delivered via the `zip` target (already configured). The ad-hoc signature used for local universal builds cannot auto-update — a real Developer ID signature and notarization are required for the feature to work for end users.
- **Windows:** NSIS builds auto-update unsigned, but an unsigned installer shows SmartScreen warnings; a code-signing certificate is recommended. The `portable` target does not auto-update.

## Excel compatibility boundary

This is a working spreadsheet application, not a complete clone of Microsoft Excel. Full parity requires separate engines and substantial additional product work for VBA/macros, Power Query, pivot tables, advanced chart editing, real-time collaboration, password-protected workbooks, every Excel function, perfect preservation of unsupported workbook objects, and enterprise cloud connectors.
