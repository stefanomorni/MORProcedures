# Session Context: MORProcedures

- **Date**: 2026-09-27
- **Branch**: `main`
- **State**: Development re-initiated — framework alignment + toolchain repairs verified

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

Nothing yet — this file was rewritten on 2026-09-27 to remove stale claims
(it previously asserted `RibbonCallbacks.bas` existed on disk before it did, quoted both
"8" and "9" launcher counts, described a `vba_manifest.json` mechanism that never existed,
and said the dev workbook did not exist).

## Next action

Confirm in Excel (dev workbook open): ribbon **MOR Procedures** tab → *Scramble Selection*
→ values become `CLI-XXXX`; *Restore Selection* reverts; `Ctrl+Shift+M` / `Ctrl+Shift+U`
do the same. Then republish the canonical `.xlam` when an interactive Excel session is
available.
