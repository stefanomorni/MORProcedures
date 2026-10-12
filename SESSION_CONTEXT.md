# Session Context: MORProcedures

- **Date**: 2026-10-12
- **Branch**: `main`
- **State**: Row highlighting StopIfTrue fix published to canonical .xlam

> This is a **short delta pointer**, not a handoff packet. Read in this order:
> 1. `PROJECT_HANDBOOK.md` — durable intent, constraints, architecture, decision log
> 2. `PROJECT_PROGRESS.md` — full lossless resume packet for the last session
> 3. This file — only what changed *since* that packet

## Orientation

`MORProcedures.xlam` is a **universal Excel add-in** loaded in every Excel session: the
**MOR Procedures** ribbon tab (scramble / restore selection, deterministic `CLI-XXXX`
pseudonyms), plus a legacy shared-helper library.

- **Canonical add-in**: `MORProcedures.xlam` — gitignored binary, registered with Excel.
- **Dev workbook**: `MORProcedures-dev.xlsm` — **now exists**; all launchers target it.
- **Scripts**: `scripts\` are NTFS symlinks → `Office-Automation-Framework\scripts\` (SSoT).
- **Launchers**: **9** `.bat` files at project root (8 targeted at the dev `.xlsm`;
  `run-addin-publish.bat` produces the add-in).

## Workflow law (unchanged, non-negotiable)

Never run the live watcher against the `.xlam` — a loaded add-in has `IsAddin = True` and
locks its VBA project. Develop in `MORProcedures-dev.xlsm` with `run-vba-watcher.bat`;
publish with `run-addin-publish.bat`. See `PROJECT_HANDBOOK.md` §3.

## Delta since the last session packet

- **Row highlighting rewritten and published**: `UserInteraction.bas` now implements a robust, self-healing, and self-contained table row highlighter using native Conditional Formatting.
  - Safe inspection of FormatConditions (fixed crash/silent false-positive when `UniqueValues` or `IconSet` rules are present on the table).
  - Self-healing: automatically calls `fc.ModifyAppliesToRange` when rows are added or deleted in the table.
  - Portable 1-line calling signature: `MORProcedures.Highlight_Selected_Table_Row Target` (no `Table` object or `MORFunctions` dependency needed).
  - Preserved backward compatibility: alias `Highlight_Selected_Tabe_Row` still accepts legacy 2-argument calls.
  - **Fixed 2026-10-12**: Explicitly set `existingRule.StopIfTrue = False` in `EnsureTableHighlightRule` so row highlight conditional formatting never suppresses downstream column-level font rules (such as `HasLink` styling or icons).
  - Published to canonical `MORProcedures.xlam` and verified working in `_Registro.xlsm`.

## Next action

- Feature work or testing in daily Excel workflow.
- Clean and commit sibling projects (`Registro`, `PMS 3.2`).
