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

    # The calculator (not shown): filled from the row, the boxes follow each other.
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=14")
    Check ($t -like "55|200|0.01|0.55|max_ss on this row: 3500 RPM.") "calculator from op 2: 14 dia at 200 SFM = 55 RPM, 0.55 per min ('$t')"
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=0.2")
    Check ($t -like "3820|*Above this row's max_ss*") "a small diameter goes past max_ss and says so ('$t')"
    $t = $xl.Run("ParamTable.CalcSelfTest", $ws.Range("G3"), "txtDia=14;txtRpm=100")
    Check ($t -like "100|366.5|0.01|1.0|*") "typing RPM gives the surface speed back ('$t')"

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

    # ---- Round-insert chip thinning (the calculator, from op 2: tool radius 0.5 = a 1.0 round, stepover 0.25).
    $f = $xl.Run("Planner.ChipFactor", 1, 0.25)
    Check ([math]::Abs($f - 0.8660) -lt 0.0001) "a 1.0 round insert at 0.25 deep: the chip is 0.866 x the feed ($f)"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "")
    Check ($t -like "1|0.25|0.00866|0.01|Entering angle 60*0.87 x the feed*|Set op 2 feed:  0.01  ->  0.01 per rev") "from op 2: its 0.01 feed makes a 0.00866 chip ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "txtHex=0.01")
    Check ($t -like "1|0.25|0.01|0.01155|*") "a 0.01 chip wants 0.01155 per rev ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "txtHex=0.01;!btnUseFeed;!btnSetRow")
    Check ($t -like "*|Done - op 2 feed is 0.01155" -and $ws.Range("G3").Value2 -eq 0.01155) "Use as feed, Set the row: op 2 feed 0.01155 ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G3"), "txtAp=0.6")
    Check ($t -like "*no thinning*") "past half the insert: no thinning ('$t')"
    $t = $xl.Run("ParamTable.ChipSelfTest", $ws.Range("G5"), "")
    Check ($t -like "||*Type the insert diameter*") "a row with no round insert: asks for the diameter ('$t')"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $t = $xl.Run("ParamTable.EditListSelfTest", "set", $ws.Range("G3:G4"), "0.02")
    Check ($t -like "2#op 2    feed    0.01  ->  0.02") "set window lists each cell old -> new ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "set", $ws.Range("G3:I3"), "")
    Check ($t -like "Type or pick*|False|") "set window on mixed columns, nothing typed: OK off ('$t')"

    $wb.Close($false)

    # ================================================= the continuous-improvement tools (Checks.bas)
    # ci_sample.xlsm (tests\macro_test.exe): eight ops laid out as a dump lays them out, each
    # Program check rule finding one thing, and "no-comment|op 3" ignored already (the hidden
    # list a dump writes from the part's .ptconfig).
    $cf = Join-Path $Dir "ci_check.xlsm"
    Copy-Item -LiteralPath (Join-Path $Dir "ci_sample.xlsm") -Destination $cf -Force
    $wb = $xl.Workbooks.Open($cf)
    $ws = $wb.Worksheets.Item("Lathe params")
    $tl = $wb.Worksheets.Item("Tools")
    $sm = $wb.Worksheets.Item("Summary")
    # The findings' keys (the hidden column I), top to bottom.
    function Keys ($sh) { $k = @(); for ($r = 6; $r -le 60; $r++) { $v = $sh.Cells.Item($r, 9).Text; if ($v -eq "") { break }; $k += $v }; , $k }
    function RowOfKey ($sh, $key) { for ($r = 6; $r -le 60; $r++) { if ($sh.Cells.Item($r, 9).Text -eq $key) { return $r } }; 0 }
    # A row of a sheet whose column A reads $text.
    function RowOfA ($sh, $text) { for ($r = 1; $r -le 200; $r++) { if ($sh.Cells.Item($r, 1).Text -eq $text) { return $r } }; 0 }
    function RowOfB ($sh, $text) { for ($r = 1; $r -le 200; $r++) { if ($sh.Cells.Item($r, 2).Text -eq $text) { return $r } }; 0 }

    # ---- Program check.
    $n = $xl.Run("Checks.ProgramCheck")
    $e = $xl.Run("Checks.LastError")
    $pc = $wb.Worksheets.Item("Program check")
    Check ($n -eq 9 -and $e -eq "") "program check: 9 findings ($n) '$e'"
    Check ($pc.Range("A2").Text -eq "9 findings - 1 ignored (shown greyed at the bottom)") "the count line ('$($pc.Range('A2').Text)')"
    $k = Keys $pc
    $want = @("needs-regen|op 3", "feed-unit|op 5", "css-no-max|op 2", "edge-time|T5|RPGV 1204", "air|op 4",
              "mixed-speeds|T1|ROUGH", "tool-changes|T1|Main", "coolant-off|op 3", "no-comment|op 3")
    Check (($k -join ",") -eq ($want -join ",")) "most useful first, the ignored one last: $($k -join ', ')"
    Check ($pc.Range("B6").Text -eq "Look at" -and $pc.Range("B11").Text -eq "FYI" -and $pc.Range("H14").Text -eq "yes" -and $pc.Range("H13").Text -eq "") "severity, and the ignored one marked yes"
    $bad = @()
    for ($r = 6; $r -le 14; $r++) { if ($pc.Cells.Item($r, 5).Text -eq "" -or $pc.Cells.Item($r, 6).Text -eq "" -or $pc.Cells.Item($r, 7).Text -eq "") { $bad += $r } }
    Check ($bad.Count -eq 0) "every finding says what, why and how to fix ($($bad -join ', '))"
    $h = $pc.Range("C6").Hyperlinks.Item(1).SubAddress
    Check ($h -eq "'Lathe params'!F5" -and $pc.Range("C6").Text -eq "op 3") "the op links to its cell: needs_regen of op 3 ($h)"
    $h = $pc.Range("D9").Hyperlinks.Item(1).SubAddress
    Check ($h -eq "'Tools'!A6" -and $pc.Range("D9").Text -eq "T5  RPGV 1204") "the tool links to its Tools row ($h, '$($pc.Range('D9').Text)')"
    Check ($pc.Range("E9").Text -like "T5 cuts about 24:00 on each edge - T6, the other RPGV 1204 tool in this program, runs 10:00.*") "edge time against the insert's other tool: '$($pc.Range('E9').Text)'"
    Check ($pc.Range("E10").Text -eq "op 4 cuts air 40% of its cutting time - 5:20 per part, 53:20 per batch of 10.") "air, per part and per batch: '$($pc.Range('E10').Text)'"
    Check ($pc.Range("E11").Text -like "T1 runs its ROUGH ops at different speeds: speed 500 - 600 SFM (op 1 600, op 4 500).") "mixed speeds: '$($pc.Range('E11').Text)'"
    Check ($pc.Range("E12").Text -eq "T1 is put in 3 times in 'Main' (op 1;  op 4;  op 6) - 2 tool changes per part could be saved.") "tool changes: '$($pc.Range('E12').Text)'"
    Check ($pc.Range("E8").Text -like "op 2 runs CSS*max_ss is 0." -and $pc.Range("E7").Text -like "op 5's feed is 0.5 per min*") "CSS with no cap; a per-minute feed that looks per rev"
    Check ($pc.Range("H6").Validation.Formula1 -eq "yes" -and $pc.Columns.Item(9).Hidden) "Ignore is a yes dropdown; the key column is hidden"
    Check ($pc.Range("G1").Text -like "?*") "a ? link to the manual ('$($pc.Range('G1').Text)')"

    # Ignore one (css), un-ignore the other (no-comment) - taken in at the next check.
    $pc.Range("H8").Value2 = "yes"
    $pc.Range("H14").Value2 = ""
    Check ($pc.Range("E8").DisplayFormat.Font.Color -eq 10526880) "a finding set to yes greys at once"
    $n = $xl.Run("Checks.ProgramCheck")
    $pc = $wb.Worksheets.Item("Program check")
    $k = Keys $pc
    Check ($n -eq 9 -and $k[8] -eq "css-no-max|op 2" -and $pc.Range("H14").Text -eq "yes" -and $k[7] -eq "no-comment|op 3" -and $pc.Range("H13").Text -eq "") "ignored css goes to the bottom, no-comment comes back ($($k -join ', '))"
    $ik = $xl.Run("Checks.IgnoredKeys")
    Check ($ik -eq "css-no-max|op 2" -and $wb.Worksheets.Item("Ignored findings").Visible -eq 0) "the hidden list holds just css ('$ik')"
    Check ($pc.Range("E14").DisplayFormat.Font.Color -eq 10526880 -and $pc.Range("E6").DisplayFormat.Font.Color -ne 10526880) "the ignored line is grey, the others not"

    # Fixed: the finding goes; its ignore stays (it would stay ignored if it came back).
    $ws.Range("M4").Value2 = 3000
    $n = $xl.Run("Checks.ProgramCheck")
    $pc = $wb.Worksheets.Item("Program check")
    $k = Keys $pc
    Check ($n -eq 8 -and -not ($k -contains "css-no-max|op 2") -and $pc.Range("A2").Text -eq "8 findings") "max_ss typed on op 2: 8 findings, no CSS one ('$($pc.Range('A2').Text)')"
    Check ($xl.Run("Checks.IgnoredKeys") -eq "css-no-max|op 2") "its ignore is kept"

    # A usual edge time typed on Tools: T1's CNMG 432 edge (9:00, its longest) against 6:00.
    $tl.Range("J10").Value2 = "6:00"
    $n = $xl.Run("Checks.ProgramCheck")
    $pc = $wb.Worksheets.Item("Program check")
    $r = RowOfKey $pc "edge-time|T1|CNMG 432"
    Check ($n -eq 9 -and $r -gt 0 -and $pc.Cells.Item($r, 5).Text -like "T1 cuts about 9:00 on each edge - the usual for CNMG 432 is 6:00 (Tools page), so 50% longer.*") "usual edge time 6:00: T1 flagged ('$(if ($r) { $pc.Cells.Item($r, 5).Text })')"

    # The part's .ptconfig: what a dump keeps from this workbook, saved - the hidden list and
    # an Ignore set on the sheet since the last check, and the usual edge time.
    $r = RowOfKey $pc "no-comment|op 3"
    $pc.Cells.Item($r, 8).Value2 = "yes"
    $saved = Join-Path $Dir "CIPART_lathe_params_saved.xlsm"
    if (Test-Path -LiteralPath $saved) { Remove-Item -LiteralPath $saved -Force }
    $wb.SaveCopyAs($saved)
    $harvest = @(& (Join-Path $Dir "partconfig_test.exe") --harvest $saved)
    Check (($harvest -contains "ignore: css-no-max|op 2") -and ($harvest -contains "ignore: no-comment|op 3") -and ($harvest -contains "usual: CNMG 432 = 6:00")) "a dump keeps the ignores and the usual edge time ($($harvest -join '; '))"
    $pc.Cells.Item($r, 8).Value2 = ""

    # ---- Slowest ops: 1:13:00 in all; ops 7, 4, 8 and 1 make 80% of it.
    $s = $xl.Run("Checks.SlowestOps")
    Check ($s -like "4 of 8 ops make 80% of the cycle time (1:01:00 of 1:13:00) - longest first: op 7 25:00, op 4 15:00, op 8 11:00, op 1 10:00.*" -and $xl.Run("Checks.LastError") -eq "") "slowest ops: '$s'"
    Check ($xl.Run("Planner.VisibleOps") -eq 4 -and $ws.Rows.Item(4).Hidden -and -not $ws.Rows.Item(9).Hidden -and $ws.Range("A3").Value2 -eq 1) "the sheet shows those 4, in their own order"
    $sl = $wb.Worksheets.Item("Slowest ops")
    Check ($sl.Range("B5").Text -eq "op 7" -and $sl.Range("F5").Text -eq "25:00" -and $sl.Range("H8").Text -eq "83.6%" -and $sl.Range("B5").Hyperlinks.Count -eq 1) "ranked on the Slowest ops sheet, with links ($($sl.Range('H8').Text))"
    Check ($xl.Run("Planner.ShowAllOps") -eq 8) "Show all brings every op back"

    # ---- Change report: op 1 faster (feed up, its time 10:00 -> 9:00), op 4 one flip fewer.
    $ws.Range("I3").Value2 = 0.014
    $ws.Range("T3").Value2 = 540
    $ws.Range("S3").Value2 = 480
    $ws.Range("Q6").Value2 = 0
    $xl.Calculate()
    $n = $xl.Run("Checks.ChangeReport")
    $rp = $wb.Worksheets.Item("Change report")
    Check ($n -eq 2 -and $xl.Run("Checks.LastError") -eq "") "change report: 2 changed cells (op 1 feed, op 2 max_ss) ($n)"
    Check ($rp.Range("A1").Text -eq "Change report - CIPART.mcam") "its heading names the part ('$($rp.Range('A1').Text)')"
    $r = RowOfA $rp "Cycle time per part"
    Check ($r -gt 0 -and $rp.Cells.Item($r, 5).Text -eq "1:13:00" -and $rp.Cells.Item($r, 6).Text -eq "1:12:00" -and $rp.Cells.Item($r, 7).Text -eq "-1:00  (-1.4%)") "cycle time 1:13:00 -> 1:12:00 ('$(if ($r) { $rp.Cells.Item($r, 7).Text })')"
    $r = RowOfA $rp "Cycle time per batch of 10"
    Check ($r -gt 0 -and $rp.Cells.Item($r, 5).Text -eq "12:10:00" -and $rp.Cells.Item($r, 6).Text -eq "12:00:00") "per batch of 10: 12:10:00 -> 12:00:00"
    $r = RowOfA $rp "Insert flips per part"
    Check ($r -gt 0 -and $rp.Cells.Item($r, 5).Text -eq "5" -and $rp.Cells.Item($r, 6).Text -eq "4") "insert flips 5 -> 4"
    $r = RowOfA $rp "Insert cost per part (average)"
    $sr = RowOfB $sm "Insert cost per part (average)"
    Check ($r -gt 0 -and $rp.Cells.Item($r, 5).Text -match "14[.,]38" -and $rp.Cells.Item($r, 6).Text -match "11[.,]25" -and [math]::Abs($sm.Cells.Item($sr, 3).Value2 - 11.25) -lt 1e-9) "insert cost per part 14.38 -> 11.25, as the Summary says ($(if ($r) { $rp.Cells.Item($r, 6).Text }) / $($sm.Cells.Item($sr, 3).Value2))"
    $r = RowOfA $rp "Insert cost per batch of 10 (whole inserts)"
    $sr = RowOfB $sm "Insert cost per batch (whole inserts)"
    Check ($r -gt 0 -and $rp.Cells.Item($r, 5).Text -match "160[.,]00" -and $rp.Cells.Item($r, 6).Text -match "122[.,]50" -and [math]::Abs($sm.Cells.Item($sr, 3).Value2 - 122.5) -lt 1e-9) "per batch 160.00 -> 122.50, as the Summary says"
    $fr = 0; for ($q = 1; $q -le 80; $q++) { if ($rp.Cells.Item($q, 4).Text -eq "feed") { $fr = $q; break } }
    Check ($fr -gt 0 -and $rp.Cells.Item($fr, 3).Text -eq "Feed" -and $rp.Cells.Item($fr, 5).Text -eq "0.012" -and $rp.Cells.Item($fr, 6).Text -eq "0.014" -and $rp.Cells.Item($fr - 1, 1).Text -eq "op 1" -and $rp.Cells.Item($fr - 1, 7).Text -eq "10:00 -> 9:00  (-1:00)") "the change under its op: Feed 0.012 -> 0.014, op 1 10:00 -> 9:00"
    $rp.Cells.Item($fr, 8).Value2 = "tested on the floor"
    [void] $xl.Run("Checks.ChangeReport")
    $rp = $wb.Worksheets.Item("Change report")
    Check ($rp.Cells.Item($fr, 8).Text -eq "tested on the floor") "a Why typed in stays when the report is made again"
    Check ($rp.PageSetup.Orientation -eq 2 -and $rp.PageSetup.FitToPagesWide -eq 1 -and $rp.PageSetup.PrintTitleRows -ne "") "printable: landscape, one page wide, the heading on every page ($($rp.PageSetup.Orientation), $($rp.PageSetup.FitToPagesWide), $($rp.PageSetup.PrintTitleRows))"
    Check ($ws.Range("I3").Value2 -eq 0.014 -and $ws.Range("E3").Value2 -eq 1) "the tools wrote nothing on the sheet"

    $wb.Close($false)
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_macros: $failed FAILED"; exit 1 }
"check_macros: all checks passed"
