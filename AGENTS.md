# MORProcedures — Agent Directives

Universal Excel add-in (`MORProcedures.xlam`) hosted in every Excel session.
Coding root: `%PROJECTS_ROOT%\MORProcedures`.

## Workflow law — read before touching VBA

This project is an **add-in**, not a workbook. The live watcher **cannot** target the
`.xlam`: a loaded add-in has `IsAddin = True` and locks its VBA project — direct
`VBComponents.Import` corrupts it.

1. Develop in `MORProcedures-dev.xlsm` with `run-vba-watcher.bat`.
2. Publish with `run-addin-publish.bat` (export → stage as `xlAddIn` → copy to canonical `.xlam`).
3. Never run `vba_sync_watch.py watch` against `MORProcedures.xlam` (the script refuses).

See `PROJECT_HANDBOOK.md` §3 and `%PROJECTS_ROOT%\Office-Automation-Framework\docs\SINGLE-WRITER-RULE.md`.

## Script SSoT

`scripts/` contains NTFS **symlinks** into `%PROJECTS_ROOT%\Office-Automation-Framework\scripts\`.
Never edit a `.py` here — commit the change in the framework, then run `..\Office-Automation-Framework\Promote-SSoT.ps1`.

## Directive contract (mandatory)

Refer to and enforce the following workstation directives:

- `%KB_ROOT%\25-Knowledge\Code-Quality-Directives\01-universal-principles.md`
- `%KB_ROOT%\25-Knowledge\Code-Quality-Directives\05-excel-vba.md`

Non-negotiables for this add-in:

- **Three-tier separation**: `RibbonCallbacks.bas` is presentation only; masking logic lives in `DataMasker.cls`.
- **Bulk memory arrays**: range data is moved via 2D `Variant` / `.Value2`, never cell-by-cell.
- **State restoration**: guarantee `ScreenUpdating`, `EnableEvents`, `Calculation` restore on every exit path.
- **`Option Explicit`** in every module; no `.Select` / `.Activate`.
- **CRLF line endings** for every `.bas`, `.cls`, `.frm` — LF breaks `VBComponents.Import`.
- Config-first: no hardcoded paths or workbook names in logic.

## Ribbon icons

Icon SSoT: `%INFRASTRUCTURE_ROOT%\Assets\Office-Ribbon-Icons\` (Axialis **Pure Flat**; see its `README.md`).
Use `image="<stem>"` + `scripts\ribbon_sync.py` to embed. No decorative Business-pack icons, no guessed `imageMso`.

## Runtime

- Project env: `vba-config.toml` → `python_env = "xlpy"`.
- Resolve with `scripts\Resolve-OfficePython.ps1`; never assume bare `python` (Inkscape shadows it on PATH).