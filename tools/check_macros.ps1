<#
    Real-Excel check of the workbook macros: opens macro_sample.xlsm (tests\macro_test.exe)
    with macros enabled and drives them - two-way links, cross-checks, the cell rules on
    bulk writes, scale, copy, the change list and revert; the planning tools (target time,
    even out flips, apply to a tool, filters, scenarios) and the chip-thinning calculator.
    No window is ever shown (a dialog would hang a hidden Excel): each window has a quiet
    self-test Function that sets it up, types into it and reads it back.

    Usage:  powershell -File tools\check_macros.ps1 [<dir holding macro_sample.xlsm>]
#>
param([string] $Dir = "")
$ErrorActionPreference = "Stop"
if (-not $Dir) { $Dir = if ($env:PT_TEST_OUT) { $env:PT_TEST_OUT } else { Join-Path $env:TEMP "ParamTableTests" } }
$src = Join-Path $Dir "macro_sample.xlsm"
$f = Join-Path $Dir "macro_check.xlsm"
Copy-Item -LiteralPath $src -Destination $f -Force

$failed = 0
function Check ($ok, $what) {
    if ($ok) { "  ok    $what" } else { "  FAIL  $what"; $script:failed++ }
}

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1          # msoAutomationSecurityLow: macros run
    $wb = $xl.Workbooks.Open($f)
    $ws = $wb.Worksheets.Item("Lathe params")
    $xl.EnableEvents = $true
    # The Summary opens first; the macros' sheets keep their names and code names.
    Check ($wb.Worksheets.Item(1).Name -eq "Summary" -and $wb.ActiveSheet.Name -eq "Summary" -and $ws.CodeName -eq "Sheet1") "Summary first and open, main sheet still Sheet1 ($($ws.CodeName))"
    "opened: $($xl.Windows.Item(1).Caption)  sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ', ') +
        "  modules: " + (($wb.VBProject.VBComponents | ForEach-Object { $_.Name }) -join ', ')

    # The ribbon: every button's icon is one Office has (an unknown imageMso is a blank
    # button), and every onAction is a ribbon callback in the macros.
    $ribbon = [xml] (Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "vba\ribbon.xml") -Raw)
    $code = (Get-ChildItem (Join-Path (Split-Path -Parent $PSScriptRoot) "vba\*.bas") | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    $badIcon = @(); $badAction = @()
    foreach ($b in $ribbon.SelectNodes("//*[@onAction]")) {
        try { [void] $xl.CommandBars.GetImageMso($b.imageMso, 16, 16) }
        catch { if ($_.Exception.Message -notlike "*Catastrophic*") { $badIcon += $b.imageMso } }   # a real icon fails only to marshal
        if ($code -notmatch ("Public Sub " + $b.onAction + "\(control As IRibbonControl\)")) { $badAction += $b.onAction }
    }
    Check ($badIcon.Count -eq 0) "every ribbon icon exists ($($badIcon -join ', '))"
    Check ($badAction.Count -eq 0) "every ribbon button has its callback ($($badAction -join ', '))"

    # A calculated (untracked) cell that moves with an edit is drawn blue.
    $before = $ws.Range("M3").DisplayFormat.Interior.Color
    $ws.Range("L3").Value2 = 0.3
    $xl.Calculate()
    $after = $ws.Range("M3").DisplayFormat.Interior.Color
    Check ($after -eq 16706267 -and $before -ne $after) "a calculated cell that moved is blue ($before -> $after)"
    $ws.Range("L3").Value2 = 0.25
    $xl.Calculate()
    Check ($ws.Range("M3").DisplayFormat.Interior.Color -eq $before) "and back to plain when the edit is undone"

    # Two-way link: amount -> percent (the formula), percent typed -> amount.
    $ws.Range("L3").Value2 = 0.2
    Check ([math]::Abs($ws.Range("M3").Value2 - 40) -lt 1e-9) "stepover 0.25 -> 0.2 gives percent 40 (got $($ws.Range('M3').Value2))"
    $ws.Range("M3").Value2 = 30
    Check ([math]::Abs($ws.Range("L3").Value2 - 0.15) -lt 1e-9) "percent typed 30 gives stepover 0.15 (got $($ws.Range('L3').Value2))"
    $ws.Range("L3").Value2 = 0.1
    Check ([math]::Abs($ws.Range("M3").Value2 - 20) -lt 1e-9) "then stepover 0.1 updates the typed percent to 20 (got $($ws.Range('M3').Value2))"

    # Cross-check in the status bar.
    $ws.Range("K3").Value2 = 0
    Check ("$($xl.StatusBar)" -like "*CSS with no max_ss*") "CSS with max_ss 0 is flagged: '$($xl.StatusBar)'"
    $ws.Range("K3").Value2 = 3500
    Check ($xl.StatusBar -eq $false -or "$($xl.StatusBar)" -eq "False") "flag clears when fixed"

    # Bulk writes obey each cell's rule.
    $xl.Run("ParamTable.SetCells", $ws.Range("G3:G4"), "0.012") | Out-Null
    Check ($ws.Range("G3").Value2 -eq 0.012 -and $ws.Range("G4").Value2 -eq 0.012) "set selected feeds to 0.012"
    $xl.Run("ParamTable.SetCells", $ws.Range("G3"), "-1") | Out-Null
    Check ($ws.Range("G3").Value2 -eq 0.012) "a negative feed is refused (cell rule)"
    $xl.Run("ParamTable.SetCells", $ws.Range("A3"), "99") | Out-Null
    Check ($ws.Range("A3").Value2 -eq 2) "op_idn (read-only) is refused"
    $xl.Run("ParamTable.SetCells", $ws.Range("L4"), "0.3") | Out-Null
    Check ($null -eq $ws.Range("L4").Value2) "a FINISH row's stepover (does not apply) is refused"
    $xl.Run("ParamTable.ScaleCells", $ws.Range("I3:I4"), 110) | Out-Null
    Check ($ws.Range("I3").Value2 -eq 220 -and $ws.Range("I4").Value2 -eq 330) "scale speeds 110% (got $($ws.Range('I3').Value2), $($ws.Range('I4').Value2))"
    $xl.Run("ParamTable.CopyCells", $ws.Range("J4:K4"), 2) | Out-Null
    Check ($ws.Range("J4").Text -eq "CSS" -and $ws.Range("K4").Value2 -eq 3500) "copy speed_mode/max_ss from op 2"

    # Change list and revert.
    $xl.Calculate()
    $n = $xl.Run("ParamTable.ListChanges")
    Check ($n -ge 5) "change list finds the edits ($n)"
    Check ($ws.Range("E3").Value2 -gt 0) "changes count on row 3 ($($ws.Range('E3').Value2))"
    $back = $xl.Run("ParamTable.RevertCells", $ws.Range("A3:N4"), $false)
    $xl.Calculate()
    Check ($ws.Range("G3").Value2 -eq 0.01 -and $ws.Range("I4").Value2 -eq 300) "revert puts values back ($back cells)"
    Check ($ws.Range("E3").Value2 -eq 0 -and $ws.Range("E4").Value2 -eq 0) "and the change counts return to 0"

    # Undo (Panel): a command's cells come back - its linked cells too; a cell typed over
    # since is left alone.
    $g4 = $ws.Range("G4").Value2
    $xl.Run("Panel.BeginEdit", "test set")
    $xl.Run("ParamTable.SetCells", $ws.Range("G3:G4"), "0.015") | Out-Null
    $n = $xl.Run("Panel.EndEdit")
    $u = $xl.Run("Panel.UndoLast")
    Check ($n -eq 2 -and $u -eq "2 0" -and $ws.Range("G3").Value2 -eq 0.01 -and $ws.Range("G4").Value2 -eq $g4) "undo puts a command's 2 cells back ($n, '$u')"
    $xl.Run("Panel.BeginEdit", "test set")
    $xl.Run("ParamTable.SetCells", $ws.Range("G3:G4"), "0.015") | Out-Null
    $xl.Run("Panel.EndEdit") | Out-Null
    $ws.Range("G3").Value2 = 0.02
    $u = $xl.Run("Panel.UndoLast")
    Check ($u -eq "1 1" -and $ws.Range("G3").Value2 -eq 0.02 -and $ws.Range("G4").Value2 -eq $g4) "undo leaves a cell typed over since ('$u')"
    $ws.Range("G3").Value2 = 0.01
    $l3 = $ws.Range("L3").Value2; $m3 = $ws.Range("M3").Value2
    $xl.Run("Panel.BeginEdit", "test link")
    $xl.Run("ParamTable.SetCells", $ws.Range("L3"), "0.2") | Out-Null
    $n = $xl.Run("Panel.EndEdit")
    $u = $xl.Run("Panel.UndoLast")
    Check ($n -eq 2 -and $ws.Range("L3").Value2 -eq $l3 -and $ws.Range("M3").Value2 -eq $m3) "undo takes back the stepover and the percent that followed it ($n, '$u')"
    $u = $xl.Run("Panel.UndoLast")
    Check ($u -eq "0 0") "nothing left to undo ('$u')"

    # Op rows, focus, Find op, Go to, the row mark.
    $t = $xl.Run("Panel.OpRowsSelfTest", $ws.Range("C3,G5,B3"))
    Check ($t -eq "ops 2, 9||3|5") "the ops of a selection, each once, in order ('$t')"
    $ws.Activate()
    $t = $xl.Run("Panel.ShowColumnsSelfTest", "feed|speed", $ws.Range("A3,A5"))
    Check ($t -eq "True|G3,I3,G5,I5") "a window's columns: those rows' cells selected ('$t')"
    $ws.Range("B3").Select()
    $found = $xl.Run("Panel.FindOp", "T3", $true)
    Check ($found -and $xl.ActiveCell.Row -eq 4) "Find op: T3 goes to op 7 (row $($xl.ActiveCell.Row))"
    $found = $xl.Run("Panel.FindOp", "face", $true)
    Check ($found -and $xl.ActiveCell.Row -eq 5) "Find op: words from the comment go to 'Rough face' (row $($xl.ActiveCell.Row))"
    $found = $xl.Run("Panel.FindOp", "11", $true)
    Check ($found -and $xl.ActiveCell.Row -eq 6) "Find op: an op number (row $($xl.ActiveCell.Row))"
    $xl.Run("Panel.GoToGroup", "Depth")
    Check ($ws.Columns.Item(8).Hidden -and -not $ws.Columns.Item(13).Hidden -and $xl.ActiveCell.Column -eq 12) "Go to Depth: Feeds folds, Depth opens, cursor in it"
    $xl.Run("Panel.GoToGroup", "")
    Check (-not $ws.Columns.Item(8).Hidden) "Go to All columns opens them again"
    $ws.Range("C4").Select()
    $mark = $wb.Names.Item("PT_Row").RefersTo
    $blue = $ws.Range("G4").DisplayFormat.Borders.Item(8).Color      # xlEdgeTop
    Check ($mark -eq "=4" -and $blue -eq 15426341) "the row mark follows the cursor ($mark, top border $blue)"

    # Coolant: Set selected on coolant cells - several at once, each row's own machine.
    $r = $xl.Run("ParamTable.SetCoolantCells", $ws.Range("O3:O4"), "Flood + Mist")
    Check ($r -eq "2 0" -and $ws.Range("O3").Text -eq "Flood + Mist" -and $ws.Range("O4").Text -eq "Flood + Mist") "coolant 'Flood + Mist' set on both rows ($r)"
    $r = $xl.Run("ParamTable.SetCoolantCells", $ws.Range("O3:O4"), "Thru-tool")
    Check ($r -eq "1 1" -and $ws.Range("O3").Text -eq "Thru-tool" -and $ws.Range("O4").Text -eq "Flood + Mist") "Thru-tool set where the machine has it, refused where not ($r)"

    # The coolant window (not shown): boxes from every selected row's machine, * where not all have it.
    $t = $xl.Run("ParamTable.CoolantPickerSelfTest", $ws.Range("O3:O4"), "1,3")
    Check ($t -eq "Flood|Mist|Thru-tool  *#Flood + Thru-tool") "coolant window: union with * on Thru-tool, ticks 1+3 ('$t')"
    $t = $xl.Run("ParamTable.CoolantPickerSelfTest", $ws.Range("O3:O4"), "")
    Check ($t -like "*#none") "rows that differ start unticked ('$t')"
    $t = $xl.Run("ParamTable.CoolantPickerSelfTest", $ws.Range("O4"), "")
    Check ($t -eq "Flood|Mist#Flood + Mist") "one cell starts ticked as it is now ('$t')"

    # The manual-text editor (no window): counter, limit, line breaks as CR LF.
    $t = $xl.Run("ParamTable.EditorSelfTest", "G4 X1.`r`nM01")
    Check ($t -eq "11 / 3,111 characters,  2 lines|True|11") "editor counts '$t'"
    $t = $xl.Run("ParamTable.EditorSelfTest", "a`nb")
    Check ($t -like "*|True|4") "a bare line feed becomes CR LF ('$t')"
    $t = $xl.Run("ParamTable.EditorSelfTest", ("x" * 4000))
    Check ($t -like "*over the limit by 889*|False|4000") "4000 characters: over the limit, OK off ('$t')"
    # A one-line comment (an op's, an inspection stop's): its own limit, no line breaks.
    $t = $xl.Run("ParamTable.EditorSelfTest", "ROTATE`r`nINSERT", 119, $true)
    Check ($t -eq "13 / 119 characters|True|13") "a comment: line break gone, its own limit ('$t')"
    $t = $xl.Run("ParamTable.EditorSelfTest", ("x" * 120), 119, $true)
    Check ($t -like "*over the limit by 1|False|120") "a comment past 119: OK off ('$t')"

    # The Set / Scale / Copy window (not shown): a live line of what OK would do.
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("G3:G4"), "0.02")
    Check ($t -like "2 cell(s) will change.|True|0.02") "set window: 2 feeds will change ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("G3:G4"), "-1")
    Check ($t -like "0 cell(s) will change,  2 refused*|False|-1") "a negative feed: both refused, OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("G3:G4"), "")
    Check ($t -like "Type or pick*|False|") "nothing typed: OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("H3:H4"), "#1")
    Check ($t -like "2 cell(s) will change.|True|per min") "both feed_mode cells share a list: pick from it ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("A3,G3"), "0.03")
    Check ($t -like "1 cell(s) will change,  1 refused*|True|0.03") "read-only op_idn counted as refused ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "110")
    Check ($t -like "2 cell(s) will change.*200 -> 220*300 -> 330|True|110") "scale window shows 200 -> 220 ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "abc")
    Check ($t -like "Type a percent*|False|abc") "scale by 'abc': OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G4:I4"), "#0")
    Check ($t -like "2 cell(s) will change,  1 already that.|True|2") "copy window: pick op 2 from the list ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("J4:K4"), "#0")
    Check ($t -like "0 cell(s) will change,  2 already that.|False|2") "copying what is already there: OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("J4:K4"), "")
    Check ($t -like "Pick the operation*|False|*") "copy with nothing picked: OK off ('$t')"
    Check ($ws.Range("G3").Value2 -eq 0.01 -and $ws.Range("I3").Value2 -eq 200) "the windows wrote nothing"

    # The Speed & feed window (not shown): filled from the row, the boxes follow each other.
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=14")
    Check ($t -like "55|200|0.01|0.55|This op's max spindle speed: 3500 RPM.") "speed & feed from op 2: 14 dia at 200 SFM = 55 RPM, 0.55 per min ('$t')"
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=0.2")
    Check ($t -like "3820|*|0.01|35.0|Above this op's max spindle speed of 3500 RPM*") "a small diameter goes past max_ss, says so, and the feed per minute is at 3500 ('$t')"
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=14;txtRpm=100")
    Check ($t -like "100|366.5|0.01|1.0|*") "typing RPM gives the surface speed back ('$t')"
    # The window's header and moves: no op, op 2, next / previous (hidden rows skipped),
    # following the selection. Its own moves select the op's feed and speed on the sheet.
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("A1"), "", "lblInfo|cboOp")
    Check ($t -like "Select an op (any cell in its row)*|") "not on an op: says to select one ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "", "cboOp|lblInfo")
    Check ($t -like "op 2  -  T1 DYNAMIC  -  Rough OD|From op 2's row, in inches*CSS: type a diameter*") "the header names the op; units from its row ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "!btnNext", "cboOp|txtRev")
    Check ($t -like "op 7  -  T3 FINISH*|0.008" -and $xl.Selection.Address($false, $false) -eq "G4:K4") "next: op 7, its feed and speed selected on the sheet ('$t', $($xl.Selection.Address($false, $false)))"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "!btnNext;!btnNext;!btnPrev", "cboOp")
    Check ($t -like "op 7  *") "next, next, previous: op 7 ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "!btnPrev", "cboOp|lblResult")
    Check ($t -like "op 2  *|This is the first op.") "previous from the first op stays and says so ('$t')"
    [void] $xl.Run("Planner.FilterTool", "1")
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "!btnNext", "cboOp")
    Check ($t -like "op 9  *") "next skips a hidden row: op 9 ('$t')"
    [void] $xl.Run("Planner.ShowAllOps")
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "@G5", "cboOp|txtRev")
    Check ($t -like "op 9  *|0.012") "follows the selection to op 9 ('$t')"
    # Units follow the op's units column (no tick box): a metric op works in mm and m/min.
    $ws.Range("N3").Value2 = "mm"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "txtDia=100", "txtRpm|lblDia|lblSurf|lblRev")
    $ws.Range("N3").Value2 = "in"
    Check ($t -eq "637|Diameter (mm)|Surface speed (m/min)|Feed per rev (mm)") "a metric op: 200 m/min at 100 mm = 637 RPM ('$t')"

    # The Tools page: live sums from the main sheet, the inserts table under it.
    $tl = $wb.Worksheets.Item("Tools")
    $xl.Calculate()
    Check ($tl.Range("E2").Value2 -eq 200 -and $tl.Range("E3").Value2 -eq 300) "Tools page sums per tool ($($tl.Range('E2').Value2), $($tl.Range('E3').Value2))"
    Check ($tl.Range("D6").Value2 -eq 500 -and $tl.Range("E6").Value2 -eq 125) "inserts table: 500 per part, 125 inserts at 4 edges ($($tl.Range('D6').Value2), $($tl.Range('E6').Value2))"
    $tl.Range("C6").Value2 = 2
    Check ($tl.Range("E6").Value2 -eq 250) "2 edges per insert: 250 inserts ($($tl.Range('E6').Value2))"
    Check ($ws.Range("P4").NumberFormat -eq "@") "a Text-formatted column keeps what is typed ($($ws.Range('P4').NumberFormat))"

    # Calculators.
    $rpm = $xl.Run("ParamTable.RpmFromSurface", 14, 200, $false)
    Check ([math]::Abs($rpm - 54.567) -lt 0.01) "200 SFM at 14 dia = 54.57 RPM (got $rpm)"

    # ================================================= the planning tools (Planner.bas)
    # The sample: tool 1 = ops 2, 9, 11 (rows 3, 5, 6), tool 3 = op 7 (row 4). est_seconds
    # (S) = fixed + cut time; the cut time (T) goes as dumped feed / feed x dumped speed / speed.
    # flips_part (U) = INT(cut time / insp_time); op 9's feed (G5) has a ceiling of 0.013.
    $all = $ws.Range("A3:W6")
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $xl.Calculate()
    function Total ($rows) { $t = 0; foreach ($r in $rows) { $t += $ws.Range("S$r").Value2 }; $t }
    Check ((Total 3, 5, 6) -eq 1910) "tool 1's ops take 1910 s as dumped ($(Total 3, 5, 6))"

    # ---- Hit a target time.
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3,A5,A6"), "10%", $false, $false)
    "        plan: $p"
    Check ($p -like "1|*31:50  ->  28:39*") "10% off tool 1: 31:50 -> 28:39 ('$($p.Substring(0, [math]::Min(90, $p.Length)))...')"
    Check ($p -like "*held at a limit: op 9 feed (max 0.013)*") "op 9's feed is held at its ceiling and said so"
    Check ($ws.Range("G3").Value2 -eq 0.01 -and (Total 3, 5, 6) -eq 1910) "the plan wrote nothing yet"
    $rows = $xl.Run("Planner.PlanCount")
    $r = $xl.Run("Planner.ApplyPlan")
    $xl.Calculate()
    Check ($r -eq "$rows 0") "OK writes the $rows planned cells ($r)"
    Check ([math]::Abs((Total 3, 5, 6) - 1719) -lt 1.5) "and the ops now take 28:39 = 1719 s ($(Total 3, 5, 6))"
    Check ($ws.Range("G5").Value2 -eq 0.013 -and $ws.Range("G3").Value2 -gt 0.0112) "op 9 at 0.013, the others went further (op 2 $($ws.Range('G3').Value2))"
    Check ($ws.Range("S4").Value2 -eq 330) "tool 3's op is untouched"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "20%", $true, $false)
    [void] $xl.Run("Planner.ApplyPlan"); $xl.Calculate()
    Check ([math]::Abs((Total 3) - 528) -lt 1.5 -and $ws.Range("I3").Value2 -gt 200) "speeds too: op 2 20% faster = 528 s, speed $($ws.Range('I3').Value2), feed $($ws.Range('G3').Value2)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "", $false, $false)
    Check ($p -like "0|op 2 takes 11:00 now*") "no target yet: says the time now ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "0:50", $false, $false)
    Check ($p -like "0|Out of reach: 1:00 of the 11:00*") "a target under the time feeds cannot move is refused ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "abc", $false, $false)
    Check ($p -like "0|'abc' is not a time*") "a target that is not a time ('$p')"
    $t = $xl.Run("Planner.PlanSelfTest", "time", $ws.Range("G3"), 1, "30:00", 0, $false, $false)
    "        window: $t"
    Check ($t -like "*31:50  ->  30:00*|True|op;tool;time;*") "the window: op 2's tool picked, 30:00 typed - before -> after per op, OK on"
    Check ($ws.Range("G3").Value2 -eq 0.01) "the window wrote nothing"

    # ---- Even out flips: tool 1 inspects on ops 2 and 9 every 8:00 - 600 s and 900 s of cut, 1 + 1 flips.
    $p = $xl.Run("Planner.PlanFlips", $ws.Range("A3,A5,A6"), "", "insp_time")
    "        plan: $p"
    Check ($p -like "1|2 cell(s) will change.*insp_time 8:00 -> 7:35*") "even out: same 2 flips, the shortest edge time 7:35 ('$p')"
    Check ($ws.Range("Q3").Text -eq "8:00") "the plan wrote nothing yet"
    [void] $xl.Run("Planner.ApplyPlan"); $xl.Calculate()
    Check ($ws.Range("Q3").Text -eq "7:35" -and $ws.Range("Q5").Text -eq "7:35" -and $ws.Range("Q6").Text -eq "") "applied: 7:35 on both inspected ops, op 11 (no inspection) left blank"
    Check (($ws.Range("U3").Value2 + $ws.Range("U5").Value2) -eq 2) "still 2 flips"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $p = $xl.Run("Planner.PlanFlips", $ws.Range("A3,A5,A6"), "One flip fewer per tool", "insp_time")
    Check ($p -like "1|*flips 2 -> 1,  insp_time 8:00 -> 10:05*") "one flip fewer: 10:05 ('$p')"
    $p = $xl.Run("Planner.PlanFlips", $ws.Range("A3,A5,A6"), "", "feed")
    "        plan: $p"
    Check ($p -like "1|*feeds x0.9*") "even out by feed: slower feeds, same flips ('$p')"
    [void] $xl.Run("Planner.ApplyPlan"); $xl.Calculate()
    Check (($ws.Range("U3").Value2 + $ws.Range("U5").Value2) -eq 2 -and $ws.Range("G3").Value2 -lt 0.01) "applied: 2 flips at feed $($ws.Range('G3').Value2)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $p = $xl.Run("Planner.PlanFlips", $ws.Range("A4"), "", "insp_time")
    Check ($p -like "0|*Tool 3: no insp_time to change*") "a tool that does not inspect: nothing to do ('$p')"
    $t = $xl.Run("Planner.PlanSelfTest", "flips", $ws.Range("G3"), 1, "Even out - same flips, every edge cut the same time", 0, $false, $false)
    Check ($t -like "*7:35*|True|op;tool;insp_time;*") "the flips window shows the plan, OK on ('$($t.Substring(0, [math]::Min(80, $t.Length)))...')"

    # ---- Apply to every op of the tool.
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "tool", $ws.Range("G3:I3"), "#0")
    Check ($t -like "4 cell(s) will change,  2 already that.*Into ops 9, 11.|True|2") "to all ops of tool 1: 4 cells in ops 9 and 11 ('$t')"
    $t = $xl.Run("ParamTable.EditListSelfTest", "tool", $ws.Range("G3:I3"), "")
    Check ($t -like "4#op 9    feed    0.012  ->  0.01") "the window lists each cell old -> new ('$t')"
    $r = $xl.Run("Planner.ToolCopyCells", $ws.Range("G3:I3"), 2)
    Check ($r -eq "6 0" -and $ws.Range("G5").Value2 -eq 0.01 -and $ws.Range("I6").Value2 -eq 200 -and $ws.Range("G4").Value2 -eq 0.008) "copied into ops 9 and 11, not tool 3's op 7 ($r)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)

    # ---- Filters (the sheet's own AutoFilter on row 2).
    $n = $xl.Run("Planner.FilterTool", "1")
    Check ($n -eq 3 -and $ws.Rows(4).Hidden) "show tool 1's ops: 3 rows, op 7 hidden"
    $n = $xl.Run("Planner.ShowAllOps")
    Check ($n -eq 4 -and $ws.AutoFilterMode) "show all: 4 rows, the filter dropdowns stay"
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("G4"), "0.009")
    $n = $xl.Run("Planner.FilterChanged")
    Check ($n -eq 1 -and -not $ws.Rows(4).Hidden -and $ws.Rows(3).Hidden) "only changed rows: op 7 ($n)"
    [void] $xl.Run("Planner.ShowAllOps")
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)

    # ---- What-if scenarios.
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("G3"), "0.011")
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scensave", $ws.Range("A3"), "fast")
    Check ($t -like "Saves the 1 edited cell(s) on the sheet as 'fast'.|True|fast") "save window ('$t')"
    Check ($xl.Run("Planner.SaveScenario", "fast") -eq 1) "saved 'fast' (1 cell)"
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("G3"), "0.009")
    Check ($xl.Run("Planner.SaveScenario", "slow") -eq 1) "saved 'slow'"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scensave", $ws.Range("A3"), "slow")
    Check ($t -like "*REPLACES the one saved*|True|slow") "saving under a name taken says it replaces it ('$t')"
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("I3"), "210")
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scenload", $ws.Range("A3"), "fast")
    Check ($t -like "Sets 1 cell(s) from 'fast'*the 2 edit(s) on the sheet now go back*not in any scenario*|True|fast") "restore window warns the edits now are not saved ('$t')"
    $r = $xl.Run("Planner.RestoreScenario", "fast")
    Check ($r -eq "1 0 2" -and $ws.Range("G3").Value2 -eq 0.011 -and $ws.Range("I3").Value2 -eq 200) "restore 'fast': feed 0.011, the other edit back to the dump ($r)"
    $n = $xl.Run("Planner.CompareScenarios")
    $sc = $wb.Worksheets.Item("Scenarios")
    Check ($sc.Range("A3").Text -eq "As dumped" -and $sc.Range("C3").Text -eq "37:20" -and $sc.Range("A4").Text -eq "fast" -and $sc.Range("C4").Text -eq "36:25" -and $sc.Range("C5").Text -eq "38:27") "compare: dump 37:20, fast 36:25, slow 38:27 ($($sc.Range('C3').Text), $($sc.Range('C4').Text), $($sc.Range('C5').Text))"
    Check ($sc.Range("A6").Text -eq "Now (the sheet)" -and $sc.Range("F3").Value2 -eq 2) "and the sheet now; flips 2 as dumped"
    Check ($sc.Range("D10").Text -eq "feed" -and $sc.Range("E10").Text -eq "0.01" -and $sc.Range("F10").Text -eq "0.011" -and $sc.Range("G10").Text -eq "0.009" -and $sc.Range("H10").Text -eq "0.011") "parameters side by side: dumped 0.01, fast 0.011, slow 0.009, now 0.011"
    [void] $xl.Run("Planner.DeleteScenario", "slow")
    $names = $xl.Run("Planner.ScenarioNames")
    Check (@($names).Count -eq 1 -and @($names)[0] -eq "fast") "deleted 'slow': only 'fast' left"
    Check ($wb.Worksheets.Item("Scenario store").Visible -eq 0) "the store sheet is hidden"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)

    # ---- Chip thinning (the Speed & feed window's second tab). The Tools page gives each
    # tool's insert size and entering angle: T1 no angle (op 2 is DYNAMIC with tool radius
    # 0.5 - a 1.0 round, stepover 0.25), T3 a 95-degree holder (op 7, 0.008 per rev).
    $f = $xl.Run("Planner.ChipFactor", 1, 0.25)
    Check ([math]::Abs($f - 0.8660) -lt 0.0001) "a 1.0 round insert at 0.25 deep: the chip is 0.866 x the feed ($f)"
    $f = $xl.Run("Planner.AngleChipFactor", 45)
    Check ([math]::Abs($f - 0.7071) -lt 0.0001 -and $xl.Run("Planner.AngleChipFactor", 90) -eq 1) "a 45-degree entering angle: 0.71 x the feed; 90: no thinning ($f)"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "")
    Check ($t -like "1|0.25|0.00866|0.01|Entering angle 60*0.87 x the feed*|This op's feed is 0.01 per rev - no change") "from op 2: its 0.01 feed makes a 0.00866 chip ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "", "optRound|lblInsert|lblNow")
    Check ($t -like "1|T1 insert: CNMG 432.  A round insert*|Now: feed 0.01 per rev makes a 0.00866 chip.") "op 2 is taken as a round insert ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "txtHex=0.01")
    Check ($t -like "1|0.25|0.01|0.01155|*|Set this op's feed:  0.01  ->  0.01155 per rev") "a 0.01 chip wants 0.01155 per rev ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G3"), "txtHex=0.01;!btnSetChip", "lblResult|txtFn")
    Check ($t -eq "op 2 feed 0.01 -> 0.01155 per rev - Ctrl+Z to undo|0.01155" -and $ws.Range("G3").Value2 -eq 0.01155) "Set this op's feed: op 2 feed 0.01155, said in the window ('$t')"
    $u = $xl.Run("Panel.UndoLast")
    Check ($u -eq "1 0" -and $ws.Range("G3").Value2 -eq 0.01) "one undo puts the feed back ('$u')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "txtAp=0.6")
    Check ($t -like "*no thinning*") "past half the insert: no thinning ('$t')"
    # A straight edge: op 7's holder enters at 95 (lead -5); type 45, or a US lead angle.
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "", "optLead|txtKr|txtLead|txtHex|lblThin|lblInsert")
    Check ($t -like "1|95|-5|0.00797|Entering angle 95* (lead -5*): the chip is 1.00 x the feed*|T3 insert: CNMG 432.  Holder's entering angle 95* (Tools page).") "op 7: a 95-degree holder, chip 0.00797 at 0.008 ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "txtKr=45", "txtLead|txtHex|lblThin")
    Check ($t -like "45|0.00566|Entering angle 45* (lead 45*): the chip is 0.71 x the feed - program 1.41 x the chip you want.") "a 45-degree entering angle: chip 0.00566 ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "txtLead=15", "txtKr|txtHex")
    Check ($t -eq "75|0.00773") "a 15-degree US lead angle is a 75 entering angle: chip 0.00773 ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "txtKr=45;txtHex=0.008;!btnSetChip", "lblResult")
    Check ($t -eq "op 7 feed 0.008 -> 0.01131 per rev - Ctrl+Z to undo" -and $ws.Range("G4").Value2 -eq 0.01131) "a 0.008 chip at 45: op 7 feed 0.01131 ('$t')"
    [void] $xl.Run("Panel.UndoLast")
    Check ($ws.Range("G4").Value2 -eq 0.008) "and undone"
    # A per-minute feed: per rev at the op's RPM, and written back per minute.
    $ws.Range("J4").Value2 = "RPM"; $ws.Range("I4").Value2 = 1000; $ws.Range("H4").Value2 = "per min"; $ws.Range("G4").Value2 = 8
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "", "txtRev|txtFn|lblNow")
    Check ($t -eq "0.00800|0.00800|Now: feed 0.008 per rev makes a 0.00797 chip.") "8 per min at 1000 RPM is 0.008 per rev ('$t')"
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G4"), "txtKr=45;txtHex=0.008;!btnSetChip", "lblResult")
    Check ($t -eq "op 7 feed 8 -> 11.31 per min - Ctrl+Z to undo" -and $ws.Range("G4").Value2 -eq 11.31) "written back per minute: 11.31 ('$t')"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    # A feed the cell refuses (op 9's ceiling 0.013): nothing written, said why.
    $t = $xl.Run("ParamTable.CalcProbe", $ws.Range("G5"), "txtKr=90;txtFn=0.02;!btnSetChip", "lblResult")
    Check ($t -like "Refused - op 9 feed: 0.02 is outside*" -and $ws.Range("G5").Value2 -eq 0.012) "a refused feed is left alone and said why ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G5"), "")
    Check ($t -like "1|||0.012|Type the holder's entering angle*") "op 9: no angle on the Tools page - asks for it ('$t')"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $t = $xl.Run("ParamTable.EditListSelfTest", "set", $ws.Range("G3:G4"), "0.02")
    Check ($t -like "2#op 2    feed    0.01  ->  0.02") "set window lists each cell old -> new ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("G3:I3"), "")
    Check ($t -like "Type or pick*|False|") "set window on mixed columns, nothing typed: OK off ('$t')"

    $wb.Close($false)
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_macros: $failed FAILED"; exit 1 }
"check_macros: all checks passed"
