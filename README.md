# Mastercam Parameter Table Tool

A **Mastercam 2026** add-in (C-Hook) that puts a part's toolpath parameters into one Excel
workbook, lets you edit them there, and loads the changes back into the operations - with a
preview of every change before anything is written.

Built for lathe and mill-turn work: change feeds, speeds, depths of cut, stock to leave and
coolant across dozens of operations in a spreadsheet instead of one dialog at a time - and see
what each change does to the cycle time before you load it.

## What it does

**Dump** - a window asks which operations (all, or only those selected in the Operation
Manager; by kind), where to save and what to call the file, then writes one `.xlsx`:

- **One row per operation**, in Operation Manager order; the most-used parameters first
  (feeds and speeds, depth of cut, stock to leave, coolant), the rest in collapsible groups.
- **Readable values.** A feed is its size plus `per rev` / `per min`; a speed is its size
  plus `CSS` / `RPM` and `CW` / `CCW`; coolant is the machine's own coolant names. Dropdowns
  wherever there is a fixed set of choices.
- **Either coolant style.** A machine set up for X-style coolant gets coolant before / with /
  after the move (several at once); a machine set up for V9 coolant gets its one setting -
  Off, Flood, Mist or Thru-tool. Each operation follows its own machine.
- **Checked as you type.** Every editable cell carries the limits the load enforces, and a
  tooltip saying what the column is. Read-only cells, and cells that do not apply to that kind
  of operation, refuse input and say why.
- **Edits highlighted.** A hidden copy of the dumped values lets every changed cell light up,
  and a `changes` column counts each row's edits - sorting and filtering do not confuse it.
- **Toolpath stats** from the NCI: cycle time, cut and rapid length, travel range, the length
  run at each feed rate.
- **A live cycle-time estimate.** `est_cycle_time` starts equal to Mastercam's cycle time and
  moves as you edit feed, speed, CSS/RPM, max spindle speed or feed mode on that row. The model
  walks the toolpath (lines and arcs, CSS turned to RPM at each diameter, the max-RPM cap, each
  section's own spindle setting) and reproduces Mastercam's cycle time to within seconds.
- **A Tools sheet** with each lathe tool's picture, linked from the `tool` column.
- **Optional macros** (an `.xlsm`, off unless ticked): a *Parameter Table* ribbon tab -
  set / scale / copy across selected cells in a window that shows how many cells will change
  or be refused before anything is written, revert cells, rows or everything to the dump, a
  Changes sheet listing every edit, a live speed and feed calculator filled from the row
  (SFM / RPM at a diameter, per rev / per minute, warns past max_ss), tick boxes for
  coolants (each row checked against its own machine's list), a big resizable editor for
  manual-entry text with a live count against the 3,111-character limit (double-click the
  cell) - and, as you type, linked amount / percent pairs kept in step and gentle warnings in the status bar (CSS with no max RPM, a
  per-minute feed that looks like a per-rev one). The sheet loads back the same without them.
  Planning tools on the same tab, each showing every op before -> after before it writes:
  - *Hit a target time* - pick ops or a tool, type the time you want (`12:30`) or a cut
    (`10%`); the feeds (and speeds, if ticked) scale by one factor, each inside its own
    limits, worked out on the sheet's own live estimate - time, flips and MRR per op, and
    which cells hit a limit.
  - *Even out flips* - the insp_time (or feeds) that give each tool the same flips per part
    with every edge used evenly, or one flip fewer, or a count you type.
  - *To all ops of tool* - copy the selected columns of one op into every op of its tool.
  - *What-if scenarios* - save the sheet's edits under a name, restore one, and compare
    them side by side (cycle time, cutting time, flips, MRR, every changed parameter).
  - *Show* - only the changed rows, only this tool's ops, or all (the sheet's own filter).
  - The calculator also does round-insert chip thinning for dynamic turning (chip
    thickness <-> the feed to program at a depth) and can set the row's feed.
- Never overwrites a file - a name that is taken gets ` (2)`.

**Load** - pick the edited, saved workbook (or a CSV). A preview window lists every change
grouped by operation - old value struck through in red, new value in green - and every
refused row with the reason. **Apply** writes; **Cancel** writes nothing.

### Operation kinds

Lathe: rough, finish, dynamic rough, prime turning, face, groove, plunge rough, drill,
manual entry. Mill (live tooling): contour, drill, 2D dynamic / peel mill. Other kinds are
listed by the dump as "not read yet".

## Safety rules (what a load may change)

- An **empty cell means "leave it alone"**, never zero.
- A row is matched to its operation by `op_idn` **and checked against its kind**.
- **Only a cell that differs is a change**; an untouched sheet writes nothing.
- A row with **any** bad cell is **refused whole** - half a row applied is a state nobody chose.
- Read-only columns are reported if edited and never written.
- **Every old value is written to `ParamTable.log` first** - there is no undo, the log is the
  way back. Changed operations are marked for regeneration.

## Install

1. Download the release zip and unzip it.
2. Close Mastercam.
3. Copy `ParamTable.ft` **and** the `ParamTable` folder into
   `Documents\My Mastercam 2026\Mastercam\Add-Ins`.
4. Start Mastercam. The two functions appear under **Customize** as
   *Parameter Table Tool* - "Lathe params - dump to Excel" and "Lathe params - load".

`ParamTable\SHA256SUMS.txt` has the hash of each file; `BUILD-INFO.txt` says which commit and
toolset built it. The add-in uses only Mastercam's own DLLs, MFC and Windows - no networking.

## Use

1. Save the part. Run **dump**; choose operations, folder and name; the workbook opens.
2. Edit in Excel. Hover a cell for what it is; watch `est_cycle_time` as you change feeds.
3. **Save** the workbook (as `.xlsx` - no "Save As CSV").
4. Run **load**; it starts at the sheet you last dumped. Read the preview, then Apply.
5. Regenerate the changed operations, check them, save the part.

## Limits

- Values are the stored values, unconverted (lengths in the part's units, some angles in
  radians - the tooltip says so).
- The time estimate leaves out rapids, dwells and tool changes (they are in the calibration,
  so the estimate starts at Mastercam's figure; only the feed-and-speed part moves).
- Mill tool pictures are not drawn yet.

## Build

Visual Studio with the **v143 14.40** toolset (the 2026 SDK does not build with v145) and the
Mastercam 2026 SDK at `C:\Program Files\Mastercam 2026\sdk`.

```
python tools\genproj.py
build.bat
tests\run_tests.bat
```

`tests\run_tests.bat` builds and runs the SDK-free tests (CSV, load rules, workbook writer and
reader - including its own DEFLATE - file naming, and the estimate formula);
`tools\check_estimate.ps1` checks the estimate formula in real Excel.

The macros are plain text in `vba\`; `tools\build_vba.ps1` compiles them with Excel into
`res\vbaProject.bin` (needs *Trust access to the VBA project object model*), which the add-in
embeds. `tools\check_macros.ps1` drives them in real Excel. `deploy.ps1` installs on
the build machine; `package.ps1` makes the release zip.

## Layout

| | |
|---|---|
| `src\LatheFields.*` | every operation kind's columns, limits and where each lives in the SDK |
| `src\Dump.cpp`, `src\DumpDialog.*` | the dump and its window |
| `src\Load.cpp`, `src\Plan.*`, `src\Preview.*` | the load, its rules (SDK-free), its preview |
| `src\Xlsx.cpp`, `src\XlsxRead.cpp` | the workbook writer and reader (no libraries) |
| `src\Paths.*`, `src\Estimate.*` | the NCI walk and the live estimate formula |
| `src\Coolant.*` | X-style coolant: codes, machine labels, write-back |
| `vba\` | the workbook macros (module, sheet events, ribbon XML) |

## Licence

MIT - see [LICENSE](LICENSE). Not affiliated with or endorsed by CNC Software; Mastercam is
their trademark. Building needs their Mastercam 2026 SDK, which is not included here.
