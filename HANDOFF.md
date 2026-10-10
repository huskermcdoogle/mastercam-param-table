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
- **Built, NOT yet run in Mastercam:** mill tool pictures (`ToolPictures::MillTool`: the tool
  manager's `CreateITlBitmapFactory ()->Create (TlToolMill, ...)`, deliberately not the setup
  sheet's `WriteMillToolImage` - setup-sheet code is what raised the same-tool-number prompt);
  mill `mrr` (`src\MillMrr.*`: ae x ap x feed/min, formulas checked in real Excel); TRANSFORM
  rows (read-only table + `xf_detail` / `xf_instances` / `xf_sources` from `GetSourceOpIDs`,
  their own NCI walked for time and flips; `src\Xform.*`).
- **Summary page (2026-10-09, not yet run from Mastercam):** `src\Summary.*` (SDK-free) builds
  the first sheet - part totals (moved off Tools), a batch with insert cost, and three ranked
  lists. Ranking is LIVE: one hidden working column (I) of single-cell ARRAY formulas,
  `MATCH(LARGE(key,n),key,0)` with key = value - ROW/1e9 (ties keep sheet order); the visible
  cells INDEX by it. `tools\check_summary.ps1` drives it in real Excel. Sheet order is now
  Summary, Lathe params, Tools, Dumped: the filter's localSheetId follows; code names unchanged.
- **Dump window:** find box (highlights matches via tree custom draw; Tick / Untick matches)
  and saved ticks (`HKCU\Software\ParamTableTool\Selections\<part file name>`, one value per
  name, op numbers). Logic in `src\Pick.*` (SDK-free, `tests\pick_test.cpp`).
- **Planning macros (2026-10-09):** `vba\Planner.bas` + `vba\PlanBox.vb` - hit a target time,
  even out flips, apply to every op of a tool, filters, what-if scenarios (hidden sheet
  "Scenario store", compare on a "Scenarios" sheet), round-insert chip thinning in the
  calculator. A plan is worked out on the sheet's OWN live formulas: trial values in, the
  sheet recalculated, est_seconds / flips_part / mrr_avg read, old values back (events off,
  calculation manual) - then OK writes through TryWrite. `tools\check_macros.ps1`: 94 checks.
  Bulk-edit results go to the status bar (the window already previewed them), not a box.
- **Macro rework (2026-10-09, not yet run by hand in Excel - only through the self-tests):**
  `vba\Panel.bas` is what every command and window shares: an undo journal (every write goes
  through `ParamTable.TryWrite` / `Panel.Journal`; `BeginEdit` ... `EndEdit` around a command
  or a window's Apply; Ctrl+Z via `Application.OnUndo`; 20 deep; a cell typed over since is
  left alone), op rows (`OpRows` - one cell is itself, hidden rows skipped), `ShowColumns`
  (open the columns' groups, select those rows' cells), a registry of modeless windows that
  follow the selection (`Opened` / `Closed`; `FollowRows` / `Refresh` on each window;
  `Following` while a window moves the sheet itself), Go to / Find op, the row mark (a
  conditional format on the hidden name `PT_Row`), `ShowHelp` (help\<topic>.html via the
  name `PT_Help` the dump writes, else Mastercam's UserDir). Op windows: `TextEditor.vb`
  (Comments & text, tabs on a MultiPage), `CoolantPicker.vb` (V9 + X-style), `Calculator.vb`
  (Speed & feed + chip thinning; Tools columns "Insert size (IC)" / "Entering angle" from the
  holder ISO style letter, kept typed-over in .ptconfig), `InspectBox.vb` (replaces Even out
  flips). `vba\Checks.bas`: Program check (ignore keys in a hidden sheet + .ptconfig
  `[ignore]`), Slowest ops, Change report. Copy from op absorbed To all ops of tool; percent
  is signed everywhere (+10% more, -10% less). `tools\check_macros.ps1` draws every ribbon
  icon (Office's orange "no picture" dot fails) - GoalSeek has none.
- **History, actual vs estimated, regenerate on purpose (2026-10-10, built and tested in real
  Excel; NOT yet run from Mastercam):** `src\History.*` (SDK-free, `tests\history_test.cpp`)
  keeps `<part>.pthistory` beside the part - append-only lines `when | kind | name: value ...`
  (format in History.h): dump / load / why / undo / regen / actual. The Summary has "From the
  machine" typed cells (labels in `Summary::Machine`; the time's seconds in the hidden column I,
  parsed with LEFT/MID - TIMEVALUE wraps at 24 h); a dump and a load harvest them and the Change
  report's Why column (each once - keys in History.cpp). The dump writes a "History" page
  (`Xlsx::Sheet::pages`: parts `page<n>.xml`, ids `rIdP<n>`, sheetIds 100+, after the Summary -
  the filter's localSheetId follows; Print_Titles for the heading row). Load: an unticked box
  in the preview (`Preview::Options::choice`) regenerates the written ops; every load writes a
  load record (Impact before -> after, insert cost from `History::InsertCost`, each written op's
  estimate), undo writes an undo record (complete: yes takes the load's saving back on the
  sheet). `src\Regen.*`: `operation_manager (&op, OPMGR_NCI_REGEN, &eptr, &succf)` (MCMill.lib,
  added to the link) on dirty ops only, refused while `OperationInProcess (0)`, then
  `CalcCycleTime` per op -> regen records. New command `LatheParamsRegenEntry` (TIP 7, the
  load's icons, ribbon key R). **A real run must confirm:** that OPMGR_NCI_REGEN regenerates
  synchronously (the log says how long each call took and whether Mastercam was still busy
  after), that it clears db.nci_flag in the op list and the database (logged when not), that
  the Operation Manager shows the ops clean afterwards, and that CalcCycleTime then reads the
  new toolpath's time.
- **Not implemented:** mill tool pictures (mill tools need `WriteMillToolImage`); radial chip
  thinning for mill ops (the sheet carries no cutter diameter / flutes).
- Real parts used for testing live in `tests\` locally and are NOT tracked (gitignored).
- **Batch dump (2026-10-10, built and unit-tested, NOT yet run in Mastercam):** "Parameter Table -
  dump a folder of parts" (`ParamTableBatchEntry`, TIP 7, ribbon key B). `src\Batch.cpp` (SDK, thin):
  refuses while `IFileManager::IsFileDirty ()`; opens each part with `IFileManager::Open (false, path)`
  (no save prompt; nothing regenerated) and checks `GetCurrentFileName ()`; `Dump::Quiet` dumps it -
  `Dump::Run`'s body was MOVED (not rewritten) into `DumpPart` in Dump.cpp, which both call; reopens
  the first part (`Open`, or `New (false)` when none was saved). Everything the SDK said goes to
  `Shop_compare_<date>.log` beside `Shop_compare_<date>.xlsx`. `src\Shop.*` (SDK-free,
  `tests\shop_test.cpp`) reads each dump's sheet as written into Parts / Inserts / Ops / About - the
  Inserts sheet's names (insert, op_kind, material, sfm_median, feed_median, depth_median,
  edge_time_median, n_ops, parts) are read by the standards library: do not rename them.
  `src\BatchDialog.*` (Ui.h, no SDK): the window and a progress window; Mastercam's window is
  disabled during the batch and only the progress window's messages and paints are pumped between
  parts (no Mastercam timer - autosave - can fire mid-batch). A real run must confirm: Open/New
  return values and dialogs; whether a dump leaves a part dirty (logged before/after); the material
  (`op_group ogi.pg2.matl_name`, every group logged to each part's ParamTable.log).

## Layout of the code
- Kinds/columns: `src\LatheFields.*` (one table per kind; `AllTables`). Mill kinds too.
- Dump: `src\Dump.cpp` (columns, groups, validations, edit tracking, live estimate formulas).
- Workbook: `src\Xlsx.cpp` (writer: stored zip, hidden "Dumped" sheet, conditional formats,
  validations, Tools sheet + drawing) and `src\XlsxRead.cpp` (own DEFLATE + zip + sheet reader).
- Load: `src\Load.cpp` -> `src\Plan.*` (SDK-free rules) -> `src\Preview.*` (the window;
  tick rules in `src\PreviewTicks.cpp`, SDK-free). Impact line: `src\Impact.*` reads
  cycle_time_raw / est_seconds / flips_part (+ the Dumped sheet's flips_part) - SDK-free.
- History and regenerate: `src\History.*` (SDK-free: the file, the harvest, the History page),
  `src\Regen.*` (the SDK calls, thin; `Regen::LastLoad` is the command).
- Undo last load (`LatheParamsUndoEntry`, `Load::UndoLast`): `src\Undo.*` (SDK-free) writes
  the load's log lines and reads them back. A load now opens with
  `load begin: <part>  <-  <sheets>`; loads logged before that line existed cannot be undone.
  Not yet run in Mastercam (built and unit-tested only; the window checked in a harness).
- Paths/estimate: `src\Paths.*` (NCI walk), `src\Estimate.*` (formula generator, SDK-free).
- The windows `src\Preview.cpp` and `src\DumpDialog.cpp` include `src\Ui.h` (MFC, NO SDK, no
  precompiled header); their one SDK need, Mastercam's main window, is `Ui::Host` in
  `src\Util.cpp`. `tests\dialog_shots.cpp` (built by `tests\build_dialog_shots.bat`, also by
  run_tests) shows the real windows without Mastercam for the manual's pictures - it gives
  its own `Ui::Host` and in-memory saved ticks, and puts the preview's lines together the way
  `Load.cpp` does (change one, change the other).

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
- `tests\run_tests.bat` (SDK-free tests; set PT_TEST_OUT to build them elsewhere); `tools\check_estimate.ps1` (real-Excel formula check).
- `deploy.ps1` (the build machine), `package.ps1` (the release zip: drop `ParamTable.ft` and
  the `ParamTable\` folder into Add-Ins - install needs no scripts).

## Deployment notes
- Install is a file copy; nothing needs running. Macros (a future, optional .xlsm) need VBA
  project trust on the BUILD machine only, to compile the macro source.

## Next (agreed)
1. Run a macro dump from Mastercam; try the ribbon tab.
2. Mill tool pictures.
3. In Mastercam: load with the regenerate box ticked on a small part, then "regenerate last
   load" on one with an op regenerated by hand; read ParamTable.log's `regen:` lines and the
   next dump's History sheet. Manual pictures still to make: help\images\NEEDED.txt.

## Traps
- Backslashes die in bash heredocs - write files with the Write/Edit tools, or use '/'.
- Dropbox locks freshly written files; test build output lives in %TEMP%\ParamTableTests.
- The PowerShell tool's working directory drifts - use absolute paths / Set-Location.
- Driving Excel by COM: a VBA MsgBox/InputBox hangs it invisibly - keep testable work in
  quiet Functions, and run checks under a timeout.
- The VBA project is PUBLIC: scan res\vbaProject.bin for user names/paths before committing.
- A VBA compile error in a COM-driven Excel shows no dialog you can see: open the built
  `ParamTableVba\vba.xlsm`, run VBE command 578 (Debug > Compile) and read
  `VBE.ActiveCodePane.GetSelection` - it stops on the bad line.
- `tools\build_vba.ps1` builds in `%PT_TEST_OUT%\ParamTableVba` when PT_TEST_OUT is set, so
  worktrees do not trip over each other.
