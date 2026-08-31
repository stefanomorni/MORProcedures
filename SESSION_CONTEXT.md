# Session Context: MORProcedures

- **Date**: 2026-08-30
- **Branch**: `main`
- **State**: Scripts scaffolded, publish-out pipeline in place

## What This Project Is

`MORProcedures.xlam` is a **universal Excel add-in** loaded in every Excel session. It hosts:
- **Mask/Unmask ribbon** — available in all open workbooks (`customUI14.xml`)
- **Shared VBA utilities** — `DataMasker.cls`, `RibbonCallbacks.bas`, and supporting modules

## Architecture

- **Canonical add-in**: `MORProcedures.xlam` — symlink → `D:\Cloud\OneDrive\...\MORProcedures.xlam`
- **Dev workbook**: `MORProcedures-dev.xlsm` — visible `.xlsm` used for active development (does NOT exist yet — create via Save As from the add-in or start fresh)
- **Scripts**: symlinks → `Office-Automation-Framework\scripts\` (single-writer SSoT)
- **Launchers**: 8 `.bat` files at project root — all target `MORProcedures-dev.xlsm`

> **Single-writer rule**: all script edits go to `Office-Automation-Framework/scripts/` only.
> Child `scripts/` entries are Windows symlinks to the SSoT. See `Office-Automation-Framework/docs/SINGLE-WRITER-RULE.md`.

## Add-In Exception — Publish-Out Pipeline

**Never run `run-vba-watcher.bat` or import into `MORProcedures.xlam` directly.**
The loaded add-in has `IsAddin=True` — VBA project is locked. Direct import causes corruption.

### Development workflow
1. Open `MORProcedures-dev.xlsm` in Excel (visible workbook)
2. Run `run-vba-watcher.bat` — live bidirectional sync with `vba-files\`
3. Edit VBA freely; watcher syncs to disk automatically
4. When ready to ship: run `run-addin-publish.bat`
   - Exports current VBA state → `vba-files\`
   - SaveAs `xlAddIn` to `_stage_MORProcedures.xlam`
   - Copies stage → `MORProcedures.xlam` (canonical)
   - Reload add-in in Excel if already loaded

### Creating MORProcedures-dev.xlsm (first time)
**Option A — In Excel VBA Editor (Fastest):**
1. Open Excel (so `MORProcedures.xlam` is loaded).
2. Press `Alt + F11` to open the VBA Editor.
3. In the Project Explorer, select `ThisWorkbook` under `MORProcedures`.
4. In the Properties window (`F4`), change `IsAddin` from `True` to `False` (the workbook immediately unhides in Excel).
5. Back in Excel, File → Save As → `MORProcedures-dev.xlsm` (Excel Macro-Enabled Workbook, `*.xlsm`) in `D:\Cloud\Coding\Projects\MORProcedures\`.
6. In VBA Editor, set `IsAddin` back to `True` on the `.xlam` (or simply close Excel without saving changes to the original `.xlam`).
7. Run `run-vba-export.bat` to verify disk sync against `MORProcedures-dev.xlsm`.

**Option B — Via Excel Add-Ins Manager:**
1. File → Options → Add-ins → Manage: Excel Add-ins → Go... → Uncheck `MORProcedures`.
2. Open `MORProcedures.xlam` directly, open VBA editor (`Alt + F11`), set `IsAddin = False`.
3. File → Save As → `MORProcedures-dev.xlsm`.
4. Close and re-check the add-in under Excel Add-Ins.

## Launchers

| File | Purpose |
|---|---|
| `run-vba-export.bat` | Export VBA from dev `.xlsm` → `vba-files\` |
| `run-vba-import.bat` | Import `vba-files\` → dev `.xlsm` |
| `run-vba-watcher.bat` | Live bidirectional watcher (dev `.xlsm` only) |
| `run-pq-export.bat` | Provisioned; add-ins have no PQ — inert |
| `run-pq-import.bat` | Provisioned; add-ins have no PQ — inert |
| `run-pq-watcher.bat` | Provisioned; add-ins have no PQ — inert |
| `run-ribbon-export.bat` | Export ribbon XML from dev `.xlsm` → `ribbon\` |
| `run-ribbon-import.bat` | Import `ribbon\customUI14.xml` → dev `.xlsm` |
| `run-addin-publish.bat` | **Publish pipeline**: dev `.xlsm` → staged → `MORProcedures.xlam` |

## Accomplishments — 2026-08-30
1. Git divergence resolved (`DataMasker.cls` preserved, rebased onto origin)
2. `scripts\` folder created with 5 symlinks → OAF SSoT
3. Full launcher suite (9 `.bat` files) written including `run-addin-publish.bat`
4. `SESSION_CONTEXT.md` created with full architecture + publish-out workflow docs
