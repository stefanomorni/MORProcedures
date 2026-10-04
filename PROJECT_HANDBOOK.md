---
project: MORProcedures
type: project_handbook
handbook_maturity: deep
habits: on
created: 2026-09-27
updated: 2026-09-27
source_code:
  - "%PROJECTS_ROOT%\\MORProcedures"
---

# Project Handbook: MORProcedures

> **Habits: ON** · **Handbook: Deep**  
> Durable project handbook (intent, constraints, architecture). Not a session log.  
> Session deltas belong in `SESSION_CONTEXT.md` / `PROJECT_PROGRESS.md`.

## 1. Purpose (plain language)

- **What this is for**: A universal Excel add-in (`MORProcedures.xlam`) that is loaded in **every** Excel session and provides shared procedures — most prominently one-click **data masking / unmasking** of selected ranges, plus a legacy library of menu, file, graphics, and formatting helpers used across Stefano's workbooks.
- **Who it serves**: Stefano as the Excel operator; the add-in is invisible infrastructure rather than a workbook he opens deliberately.
- **Analogy**: A toolbox bolted onto Excel itself — it is not a document, it is an appliance. That distinction drives every constraint below.

## 2. Use cases & workflows

- **Must support**:
  - Scramble / restore the current selection via ribbon buttons, with deterministic, reversible pseudonyms (`CLI-XXXX`).
  - Provide shared helper procedures to any open workbook (menu construction, file handling, array↔range transfer, user dialogs, status-bar messaging).
  - Load silently in every Excel session without requiring the user to open a workbook.
  - Keep all VBA as plain text in git and sync it through the Office Automation Framework.
- **Nice to have**: Extend the ribbon with more procedure groups; keep the legacy module set compiling cleanly on modern Excel.
- **Out of scope / non-goals**:
  - Being a document the user opens — it is an add-in, never a data workbook.
  - Holding project-specific business logic (that belongs in PMS 3.2 / GP-Dashboard / PMMS).
  - Power Query usage — add-ins have no meaningful PQ surface.

## 3. Constraints & restrictions

- **Hard rules**:
  - **Add-in workflow law**: never run the live VBA watcher against `MORProcedures.xlam`. A loaded add-in has `IsAddin = True` which locks its VBA project; `VBComponents.Import` corrupts it. Develop in `MORProcedures-dev.xlsm`, publish with `run-addin-publish.bat`. See `%PROJECTS_ROOT%\Office-Automation-Framework\docs\SINGLE-WRITER-RULE.md` §Sanctioned Exception.
  - **Script SSoT**: `scripts\` are NTFS symlinks into `%PROJECTS_ROOT%\Office-Automation-Framework\scripts\`. Never edit a `.py` here; commit in the framework and run `Promote-SSoT.ps1`.
  - **CRLF** for every `.bas` / `.cls` / `.frm` — LF breaks `VBComponents.Import`.
  - **Encoding**: some legacy modules are cp1252 (ANSI) rather than UTF-8, because VBA exports inherit the VBE ANSI codepage. `vba_sync_watch.py` decodes UTF-8 first, then cp1252.
  - **Never use `VBComponents.Import` to update an existing component** — it appends `Name1`, `Name2`, … instead of replacing. Safe-Sync (delete-lines + `AddFromString`) is the only correct path for existing modules.
  - Config-first: no hardcoded paths or workbook names inside logic.
  - Add-in binary and dev workbook are gitignored — **text sources only** in version control.
- **Soft preferences**: Italian naming/UI retained in the legacy library (e.g. `FinestraInformativa`, `AggiungiComandoAMenu`); new work may use English identifiers.
- **Forbidden approaches**: Decorative Business-pack ribbon icons; guessed `imageMso`; committing `*.log` / staging artefacts.
- **Office ribbon icons**: icon SSoT is `%INFRASTRUCTURE_ROOT%\Assets\Office-Ribbon-Icons\` (Axialis **Pure Flat**; see its `README.md`). Embed via `ribbon_sync.py` using `image="<stem>"`.

## 3b. Runtime / tooling

- **Python / mamba env**: `xlpy` (Office COM, watchers, ribbon sync, xlwings).
- **How to run**: project `run-*.bat` launchers, or `mamba run -n xlpy python ...`.
- **Config SSoT**: `vba-config.toml` → `python_env = "xlpy"`; `file` points at **`MORProcedures-dev.xlsm`** (never the locked `.xlam`).
- **Resolution order for agents/scripts**:
  1. `vba-config.toml` `python_env`
  2. else `OFFICE_PYTHON_ENV`
  3. else framework fallback `xlpy`
- **Note**: `pytest` is not installed in `xlpy`; the framework tests in `Office-Automation-Framework\tests\` are written to also run as plain scripts (`python tests/test_*.py`).

## 4. Architecture now

- **Where code lives**: `%PROJECTS_ROOT%\MORProcedures\` — `vba-files\`, `ribbon\`, `scripts\` (symlinks), `power-queries\` (provisioned, inert), `.vscode\tasks.json`, launchers, `vba-config.toml`.
- **Deliverables**:
  - `MORProcedures.xlam` — canonical add-in (symlink → OneDrive junction at `%APPDATA%\Microsoft\AddIns\`). Gitignored.
  - `MORProcedures-dev.xlsm` — visible dev workbook mirroring the add-in's VBA project. Gitignored.
- **Key components** (`vba-files\`, 21 components):
  - `DataMasker.cls` — business-logic tier: deterministic masking vault, 2D `Variant`/`.Value2` bulk transfers, multi-area support.
  - `RibbonCallbacks.bas` — presentation tier: `Ribbon_OnLoad`, `OnRibbonMaskClick`, `OnRibbonUnmaskClick`, `QuickMaskSelection` / `QuickUnmaskSelection`.
  - `AutoMacros.bas` — `Auto_Open` / `Auto_Close` registering `Ctrl+Shift+M` / `Ctrl+Shift+U`.
  - Legacy library: `Ambiente`, `DDECalls`, `Editing`, `GestioneFiles`, `Grafici`, `Menus`, `OnTimeProcedure`, `UserInteraction`, `Utilita`, `Visualizzazione`, `Windekis`, `Xdebug`, plus 4 UserForms and document modules.
- **Bridges**: COM via `win32com` / `xlwings` (`xlpy`); OneDrive junction for the installed add-in; GoodSync for `%CODING_ROOT%`.
- **External systems**: Excel desktop (`MORProcedures.xlam` registered as an Excel Add-In), Windows Credential Manager / `.secrets` vault file for the Python masking tool.

## 5. Decision log (append-only)

| Date | Decision | Why | Still valid? |
|------|----------|-----|--------------|
| 2026-08-30 | Publish-out pipeline (`-dev.xlsm` → stage → `.xlam`) instead of live watcher | Loaded add-in locks its VBA project; direct import corrupts it | Yes |
| 2026-08-30 | Symlink SSoT for `scripts\` | Eliminates per-project script drift | Yes |
| 2026-09-18 | Axialis Pure Flat PNG icons replace `imageMso`; project `python_env` resolved first | Blank/ugly icons; `python` on PATH could hit Inkscape | Yes |
| 2026-09-27 | `vba-config.toml` `file` retargeted from `.xlam` to `MORProcedures-dev.xlsm` | Config pointed at the locked add-in, contradicting the publish-out rule | Yes |
| 2026-09-27 | Publish pipeline moved from inline `python -c` into `scripts/publish_addin.py` | Inline raw-string path `r'%~dp0'` was a Python SyntaxError; step [2/3] always failed | Yes |
| 2026-09-27 | `vba_sync_watch.py` decodes UTF-8 then cp1252 | cp1252-safe modules crashed Safe-Sync, and the `Import` fallback spawned duplicate modules | Yes |
| 2026-09-27 | Duplicate-safe fallback: never `Import` an already-existing component | Prevents `Menus1`-style duplicates | Yes |
| 2026-09-27 | Disk `vba-files\` reconciled to the add-in + disk-ahead modules imported back | `.xlam` held newer `AutoMacros`/`UserInteraction`; disk held newer `Menus`; two modules were empty on disk | Yes |
| 2026-09-27 | Robust self-healing table row highlighting in `UserInteraction.bas` | Fix silent false-positive in FormatConditions loop when UniqueValues/IconSet rules exist; auto-adapt AppliesTo on table row resize; native Target.ListObject resolution | Yes |

## 6. Open questions & risks

- `MORFunctions` is a sibling universal add-in; if it shares the same publish-out pattern, consider pooling the launcher/publish tooling (already shared via SSoT).
- `Menus.bas` and four other legacy modules remain cp1252; converting them to UTF-8 would simplify the toolchain but changes bytes — deferred as low value.
- The legacy library is untouched by the Code-Quality directives (`.Select` usage, cell-by-cell loops may exist). Refactor opportunistically, not wholesale.
- `run-addin-publish.bat` requires Excel to be reachable from the shell's Windows session; Excel launched in another session makes COM attach fail (script now reports this clearly rather than crashing).

## 7. Validated commands / recipes

```powershell
# Resolve the project interpreter
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\Resolve-OfficePython.ps1 -ProjectRoot .

# Day-to-day development (dev workbook open in Excel)
.\run-vba-watcher.bat

# Force a structural export after add/rename/delete of modules
.\run-vba-export.bat

# Publish to the canonical add-in
.\run-addin-publish.bat

# Framework tests (pytest is not in xlpy — run as scripts)
mamba run -n xlpy python ..\Office-Automation-Framework\tests\test_strip_vba_header.py
```

## 8. Related docs

- USER_GUIDE: `%PROJECTS_ROOT%\Office-Automation-Framework\docs\USER_GUIDE.md`
- Single-writer rule: `%PROJECTS_ROOT%\Office-Automation-Framework\docs\SINGLE-WRITER-RULE.md`
- SOP: `%PROJECTS_ROOT%\Office-Automation-Framework\docs\SOP-Office-Automation-Framework.md`
- Resume packet: `PROJECT_PROGRESS.md`
- Session delta: `SESSION_CONTEXT.md`
- LEARN / KB notes:
  - `%KB_ROOT%\25-Knowledge\Development\LEARN-OAF-AddIn-PublishOut-Pipeline.md`
  - `%KB_ROOT%\25-Knowledge\Development\LEARN-OAF-SymlinkSSoT-Architecture.md`
  - `%KB_ROOT%\25-Knowledge\Development\LEARN-OAF-Ribbon-CustomPNG-and-ProjectPythonEnv.md`