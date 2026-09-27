# PROJECT_PROGRESS — MORProcedures

**Date:** 2026-09-27 (end of session)
**Branch / HEAD:** `main` — working tree contains uncommitted changes (see below); **pushed to origin: NO** (prior `79eec6e` still unpushed).
**Phase:** Development re-initiation — framework alignment + toolchain repairs. **Verified working.**
**Note:** Excel was running in a separate Windows session (`RDP-Tcp#0`); COM automation was performed by launching isolated Excel instances.

> Prior chat/transcript is **not** required to resume. Read this packet + `PROJECT_HANDBOOK.md` §5 decision log.

---

## What changed this session

### 1. Framework alignment (was: project was missing the add-in scaffold)
- Created `MORProcedures-dev.xlsm` (the dev workbook every launcher requires — it did **not** exist).
- Added `PROJECT_HANDBOOK.md` (Deep), this progress packet, `AGENTS.md` (directive contract), `Repair-ProjectSymlinks.ps1`, `.vscode/tasks.json`, `power-queries/`.

### 2. Toolchain defects fixed
- **Publish pipeline was broken.** `run-addin-publish.bat` embedded inline Python with a raw string ending in a backslash (`pathlib.Path(r'%~dp0')`) → `SyntaxError`; step [2/3] could never run. Logic moved to new `scripts/publish_addin.py` (SSoT in OAF). **Verified end-to-end** against a throwaway target (stage → publish → cleanup), canonical `.xlam` untouched.
- **`vba-config.toml`** `file` retargeted from the locked `.xlam` to `MORProcedures-dev.xlsm`.
- **`run-pq-watcher.bat`** passed a stray `watch` subcommand that `pq_sync_watch.py` does not accept (workbook is positional) → argparse error, watcher died instantly. Fixed.
- **`.gitignore`** extended (`*.log`, `*_temp.log`, `~$*`, `*.tmp`, `_stage_*.xlam`); removed 8 committed xvba export logs + stale `vba_metadata.json`.

### 3. Framework (OAF) fixes — these are the higher-value ones
- **`vba_sync_watch.py` encoding bug (root cause of duplicate modules):** the Safe-Sync branch hardcoded `codecs.open(..., "utf-8-sig")`. Five legacy modules are **cp1252** (`Ambiente`, `Editing`, `GestioneFiles`, `Grafici`, `Menus`), so the read raised `UnicodeDecodeError`, and the except-branch fell back to `VBComponents.Import` — which **appends** a new component rather than replacing. Result: a duplicate `Menus1` was created and the module stayed stale. Fix: new `_read_vba_source()` helper (UTF-8 → cp1252, raises if neither), plus the fallback now refuses to `Import` an already-existing component. Replaced the old `is_doc_module`/`Type == 100` special case (superseded).
- **`Promote-SSoT.ps1` was completely unrunnable:** `[CmdletBinding(SupportsShouldProcess)]` plus a manual `[switch]$WhatIf` → "parameter defined multiple times" at parse time. Fixed. Also added `publish_addin.py` to `$CoreScripts`.
- Added `publish_addin.py` + new encoding tests (17/17 pass; both OAF test suites green).
- `templates/Excel/.gitignore` extended to match.

### 4. Tri-surface parity reconciliation (VBA)
The `.xlam` and `vba-files\` had diverged in **both** directions — neither was a superset:

| Direction | Modules | Action |
|---|---|---|
| Add-in ahead | `AutoMacros.bas` (shortcut registration), `UserInteraction.bas` (rewritten crosshair highlighter) | Written to disk |
| Add-in only | `RibbonCallbacks.bas` | Written to disk (disk was missing it entirely, though `SESSION_CONTEXT.md` claimed it existed) |
| Disk ahead | `Menus.bas` (`FaceId` / `BeginGroup` support) | Imported into workbook |
| Disk empty, add-in populated | `AutoMacros.bas` was a 2-line stub | (covered above) |
| Cosmetic | `Editing`, `FrmEtichette`, `Utilita`, `Visualizzazione` — `.names` vs `.Names` casing (9 sites) | Disk normalised to workbook casing (zero COM ops) |

**Result: 21/21 components at parity** (single residual difference is a trailing newline in `Menus.bas`).

---

## Non-negotiables introduced/confirmed

- Never run the live watcher against the `.xlam`; publish-out only (`run-addin-publish.bat`).
- Never `VBComponents.Import` an **existing** component — duplicates. Safe-Sync only.
- Legacy modules may be **cp1252**; do not assume UTF-8 when writing encoders.
- Both the `.xlam` **and** `MORProcedures-dev.xlsm` are gitignored — text sources only.
- Shortcuts `Ctrl+Shift+M` / `Ctrl+Shift+U` are registered by `AutoMacros.Auto_Open` and **depend on module ordering**: `AutoMacros` appears first among standard modules in the VBA project (Project Explorer ordering). Preserve that ordering; a reorder could stop `Auto_Open` firing.

---

## Done inventory

| Item | Anchor |
|---|---|
| Framework alignment audit (KB/OAF vs project) | this session |
| Dev workbook created from canonical `.xlam` | `MORProcedures-dev.xlsm` |
| Publish pipeline repaired + verified | `scripts/publish_addin.py`, `run-addin-publish.bat` |
| Encoding tolerance + duplicate-safe fallback | `OAF/scripts/vba_sync_watch.py` |
| `Promote-SSoT.ps1` parse bug fixed | `OAF/Promote-SSoT.ps1` |
| Handbook / AGENTS / progress / repair / tasks / power-queries | project root |
| 17/17 unit tests + ribbon tests green | `OAF/tests/` |

## Next (exactly one)

**Push the unpushed commits.** `79eec6e` was already ahead before this session, and this session added more across two repos. Then confirm in Excel that the ribbon Mask/Unmask buttons and the two keyboard shortcuts work against the dev workbook before the next publish.

---

## Open risks (not next)

1. **The canonical `.xlam` has NOT been republished** with the reconciled sources. It currently holds `AutoMacros`/`UserInteraction` but **not** the disk-ahead `Menus` changes. Run `run-addin-publish.bat` from an interactive Excel session when ready. (Deliberate: publishing replaces the live add-in.)
2. **`MORProcedures` may no longer auto-load in Excel.** `HKCU\Software\Microsoft\Office\16.0\Excel\Addins` has no `MORProcedures` subkey and there is no `OPEN` value in `Excel\Options` — so the add-in appears to load via transient session state. If it stops loading, re-enable via File → Options → Add-Ins → Excel Add-Ins → Browse → `%APPDATA%\Microsoft\AddIns\MORProcedures.xlam`.
3. `run-addin-publish.bat` requires Excel reachable from the invoking shell's Windows session — a cross-session Excel makes COM attach fail (now reported clearly, not a traceback).
4. `AppData\Roaming\Microsoft\AddIns\` holds stale lock files (`~$MORFunctions.xlam` etc.) and legacy `.dotm` templates — harmless, not cleaned.
5. Legacy modules (`Menus`, `Editing`, …) predate the Code-Quality directives; audit only if touched.

---

## Verify before new feature work

```powershell
# 1) Symlinks intact
powershell -NoProfile -File .\Repair-ProjectSymlinks.ps1 -WhatIf

# 2) Interpreter resolves to the project env
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\Resolve-OfficePython.ps1 -ProjectRoot .

# 3) Framework suites (pytest not installed in xlpy)
mamba run -n xlpy python ..\Office-Automation-Framework\tests\test_strip_vba_header.py
mamba run -n xlpy python ..\Office-Automation-Framework\tests\test_ribbon_sync_images.py
```

Excel-side smoke test (dev workbook open):
1. Ribbon → **MOR Procedures** tab → *Scramble Selection* → selection becomes `CLI-XXXX`.
2. *Restore Selection* → original values return.
3. `Ctrl+Shift+M` / `Ctrl+Shift+U` do the same.

## Open (not next)

- Consider converting the 5 cp1252 modules to UTF-8 (changes bytes; low value).
- Compare with `MORFunctions` (sibling add-in) to see if it needs the same alignment pass.