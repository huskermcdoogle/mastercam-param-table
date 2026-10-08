# Handoff - Parameter Table Tool

Read this first. It is written so a fresh session can pick up without the history.

## What this is
A Mastercam 2026 C++ Add-In (C-Hook). **Dump** writes a formatted `.xlsx` of every supported
operation's parameters; the operator edits it in Excel and saves; **Load** reads the saved
workbook back into the operations (CSV still accepted). Its own git repo (this folder).

## State (2026-10-08)
- **Proven on real parts in Mastercam:** dump (15 ops / 11 kinds on a real mill-turn part),
  load (feeds, speeds, manual text written back), xlsx read-back from Excel-saved files, the
  preview window, X-style coolant decode and labels, NCI path walk (feed length = backplot to 4
  decimals; feed time = Mastercam's cycle time within seconds on every op with moves), the live
  estimate (checked in real Excel against edits).
- **Also proven in Mastercam:** dump window (`src\DumpDialog.*`, settings in
  `HKCU\Software\ParamTableTool`), file name rules (`src\FileRules.*`), Tools sheet with lathe
  tool pictures (`src\ToolPictures.*`, needs `MCTool.lib`), canned-drill time, coolant
  write-back (codes, dialog).
- **Macro layer (2026-10-08):** `vba\ParamTable.bas` + `vba\Sheet1.cls` + `vba\ribbon.xml`,
  compiled by `tools\build_vba.ps1` into `res\vbaProject.bin` (embedded as RCDATA). Dump window
  "Include macros" -> `.xlsm` (code names ThisWorkbook / Sheet1 are written by the writer).
  `tools\check_macros.ps1`: 21 checks pass in real Excel. Manual-text editor is a UserForm
  (`vba\TextEditor.vb`, controls made by build_vba.ps1); Set selected is coolant-aware. The
  form's MSForms reference embeds C:\Users\<name>\...\MSForms.exd - `tools\scrub_vba.py`
  (run by build_vba.ps1, needs python) overwrites the name in place.
- **Not implemented:** mill tool pictures (mill tools need `WriteMillToolImage`).
- Real parts used for testing live in `tests\` locally and are NOT tracked (gitignored).

## Layout of the code
- Kinds/columns: `src\LatheFields.*` (one table per kind; `AllTables`). Mill kinds too.
- Dump: `src\Dump.cpp` (columns, groups, validations, edit tracking, live estimate formulas).
- Workbook: `src\Xlsx.cpp` (writer: stored zip, hidden "Dumped" sheet, conditional formats,
  validations, Tools sheet + drawing) and `src\XlsxRead.cpp` (own DEFLATE + zip + sheet reader).
- Load: `src\Load.cpp` -> `src\Plan.*` (SDK-free rules) -> `src\Preview.*` (the window).
- Paths/estimate: `src\Paths.*` (NCI walk), `src\Estimate.*` (formula generator, SDK-free).

## Facts paid for (do not re-derive)
- Lathe NCI: X is a RADIUS. Lines store (X,0,Z); ARCS store (X,Z,0) - Z in the 2nd slot. Arcs
  G18, swept Z->X. Each section's 1002 line sets speed: rpm<0 = CSS. Feed sign per move: - per rev.
- `CalcCycleTime` returns seconds. Backplot's rapid counts moves into/out of the op (ours does not).
- X coolant = negative canned-text codes: -(when*1000 + 100 + k), when 3/4/5 = before/with/after,
  k odd = coolant (k+1)/2 on. Labels come from the machine definition entity (machine group).
- Excel: `INDEX(...,0)` inside a cell formula is cut by implicit intersection - use array maths.
- Prime turning keeps its own feeds/speeds in its prm struct (`pt_*` columns).
- Feed/rpm signs: feed - = per rev; rpm - = CCW. The sheet shows size + a word column.

## Build / run
- `python tools\genproj.py` then `build.bat` (MSVC v143 14.40 pinned). Close Mastercam first.
- `tests\run_tests.bat` (6 SDK-free tests); `tools\check_estimate.ps1` (real-Excel formula check).
- `deploy.ps1` (the build machine), `package.ps1` (the release zip: drop `ParamTable.ft` and
  the `ParamTable\` folder into Add-Ins - install needs no scripts).

## Deployment notes
- Install is a file copy; nothing needs running. Macros (a future, optional .xlsm) need VBA
  project trust on the BUILD machine only, to compile the macro source.

## Next (agreed)
1. Run a macro dump from Mastercam; try the ribbon tab.
2. Mill tool pictures.

## Traps
- Backslashes die in bash heredocs - write files with the Write/Edit tools, or use '/'.
- Dropbox locks freshly written files; test build output lives in %TEMP%\ParamTableTests.
- The PowerShell tool's working directory drifts - use absolute paths / Set-Location.
- Driving Excel by COM: a VBA MsgBox/InputBox hangs it invisibly - keep testable work in
  quiet Functions, and run checks under a timeout.
- The VBA project is PUBLIC: scan res\vbaProject.bin for user names/paths before committing.
