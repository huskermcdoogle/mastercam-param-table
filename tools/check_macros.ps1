<#
    Real-Excel check of the workbook macros: opens macro_sample.xlsm (tests\macro_test.exe)
    with macros enabled and drives them - two-way links, cross-checks, the cell rules on
    bulk writes, scale, copy, the change list and revert; the planning tools (target time,
    inspection and inserts, copy into a tool's ops, filters, scenarios) and the
    chip-thinning calculator.
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

    # The ribbon: every icon really draws something, and every onAction is a ribbon
    # callback in the macros. An unknown imageMso is a blank button - and some names
    # Office knows have no picture at all (GoalSeek): Office hands back an orange dot
    # for those, which on the ribbon is blank too. So each icon is drawn to a file (by a
    # throwaway workbook's macro) and compared with that dot.
    $ribbon = [xml] (Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "vba\ribbon.xml") -Raw)
    $code = (Get-ChildItem (Join-Path (Split-Path -Parent $PSScriptRoot) "vba\*.bas") | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    $badIcon = @(); $badAction = @()
    $iconDir = Join-Path $env:TEMP ("pt_icons_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $iconDir | Out-Null
    $iconIds = @("GoalSeek") + @($ribbon.SelectNodes("//*[@imageMso]") | ForEach-Object { $_.imageMso } | Select-Object -Unique)
    $tmp = $xl.Workbooks.Add()
    try {
        $mod = $tmp.VBProject.VBComponents.Add(1)
        $mod.CodeModule.AddFromString("Public Sub IconSave(ByVal ids As String, ByVal folder As String)`r`n" +
            "    Dim v As Variant, p As Object`r`n" +
            "    For Each v In Split(ids, "","")`r`n" +
            "        Set p = Nothing`r`n" +
            "        On Error Resume Next`r`n" +
            "        Set p = Application.CommandBars.GetImageMso(CStr(v), 32, 32)`r`n" +
            "        If Not p Is Nothing Then stdole.SavePicture p, folder & ""\"" & v & "".bmp""`r`n" +
            "        On Error GoTo 0`r`n" +
            "    Next`r`n" +
            "End Sub")
        $xl.Run("'" + $tmp.Name + "'!IconSave", ($iconIds -join ","), $iconDir)
    } finally { $tmp.Close($false) }
    $dot = [IO.File]::ReadAllBytes((Join-Path $iconDir "GoalSeek.bmp"))
    foreach ($id in $iconIds | Select-Object -Skip 1) {
        $f = Join-Path $iconDir "$id.bmp"
        if (-not (Test-Path -LiteralPath $f)) { $badIcon += "$id (unknown)"; continue }
        $bytes = [IO.File]::ReadAllBytes($f)
        if ($bytes.Length -eq $dot.Length -and -not (Compare-Object $bytes $dot -SyncWindow 0)) { $badIcon += "$id (no picture)" }
    }
    Remove-Item -LiteralPath $iconDir -Recurse -Force -ErrorAction SilentlyContinue
    foreach ($b in $ribbon.SelectNodes("//*[@onAction]")) {
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
    $xl.Run("ParamTable.ScaleCells", $ws.Range("I3:I4"), 10) | Out-Null
    Check ($ws.Range("I3").Value2 -eq 220 -and $ws.Range("I4").Value2 -eq 330) "scale speeds +10% (got $($ws.Range('I3').Value2), $($ws.Range('I4').Value2))"
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
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "+10%")
    Check ($t -like "10% more:  2 cell(s) will change.*200 -> 220*300 -> 330|True|+10%") "scale window: +10% is 10% more, 200 -> 220 ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "-10%")
    Check ($t -like "10% less:  2 cell(s) will change.*200 -> 180*300 -> 270|True|-10%") "scale window: -10% is 10% less, 200 -> 180 ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "110")
    Check ($t -like "'110' - say which way: +110% for more, -110% for less.|False|110") "scale by 110 (no sign): asks which way, OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "scale", $ws.Range("I3:I4"), "abc")
    Check ($t -like "*is not a percent*|False|abc") "scale by 'abc': OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G4:I4"), "#0")
    Check ($t -like "2 cell(s) will change,  1 already that.*Into op 7.|True|2") "copy window: pick op 2, into the selected row ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("J4:K4"), "#0")
    Check ($t -like "0 cell(s) will change,  2 already that.|False|2") "copying what is already there: OK off ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("J4:K4"), "")
    Check ($t -like "Pick the op to copy from.*|False|*") "copy with nothing picked: OK off ('$t')"
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
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3,A5,A6"), "-10%", $false, $false)
    "        plan: $p"
    Check ($p -like "1|*31:50  ->  28:39*") "-10% on tool 1: 31:50 -> 28:39 ('$($p.Substring(0, [math]::Min(90, $p.Length)))...')"
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
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "-20%", $true, $false)
    [void] $xl.Run("Planner.ApplyPlan"); $xl.Calculate()
    Check ([math]::Abs((Total 3) - 528) -lt 1.5 -and $ws.Range("I3").Value2 -gt 200) "speeds too: op 2 20% faster = 528 s, speed $($ws.Range('I3').Value2), feed $($ws.Range('G3').Value2)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "", $false, $false)
    Check ($p -like "0|op 2 takes 11:00 now*") "no target yet: says the time now ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "0:50", $false, $false)
    Check ($p -like "0|Out of reach: 1:00 of the 11:00*") "a target under the time feeds cannot move is refused ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "abc", $false, $false)
    Check ($p -like "0|'abc' is not a time*") "a target that is not a time ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "10%", $false, $false)
    Check ($p -like "0|'10%' - say which way: +10% for more, -10% for less.*") "10% with no sign: asks which way ('$p')"
    $p = $xl.Run("Planner.PlanTargetTime", $ws.Range("A3"), "+10%", $false, $false)
    Check ($p -like "1|*11:00  ->  12:06  (+10.0%)*") "+10% is 10% more time: 11:00 -> 12:06 ('$p')"

    # The window (not shown): opens with the main cutting feed only; Enter previews, never
    # applies - Apply is its own button, one Ctrl+Z.
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), -1, "30:00", $true, $false, "")
    Check ($t -like "Press Enter*|False|*|main") "the window starts on the main cutting feed only; typed, Apply waits for Enter ('$t')"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "30:00", $true, $false, "enter,enter")
    "        window: $t"
    Check ($t -like "*31:50  ->  30:00*|True|op;tool;time;*") "Enter twice: op 2's tool, 30:00 - before -> after per op, Apply on"
    Check ($ws.Range("G3").Value2 -eq 0.01 -and (Total 3, 5, 6) -eq 1910) "and Enter twice wrote nothing"
    $rows = @($t.Split("|")[2] -split " // ")
    Check (@($rows | Where-Object { $_.Split(";")[5].Length -gt 62 }).Count -eq 0) "the list's changes column fits (no sideways scrolling)"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "-10%", $true, $false, "enter,apply")
    $xl.Calculate()
    Check ($t -like "*|Done: 3 cell(s) changed - the ops now take 28:39 (were 31:50). Ctrl+Z to undo.|*" -and [math]::Abs((Total 3, 5, 6) - 1719) -lt 1.5) "Apply writes the plan and says so in the window ($(Total 3, 5, 6))"
    $u = $xl.Run("Panel.UndoLast")
    $xl.Calculate()
    Check ($u -eq "3 0" -and (Total 3, 5, 6) -eq 1910) "one Ctrl+Z takes the whole plan back ('$u')"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "-10%", $true, $false, "enter,set G5=0.0125,apply")
    Check ($t -like "*|True|*|The sheet changed since this was worked out - here it is again. Check it, then Apply.|*" -and $ws.Range("G5").Value2 -eq 0.0125 -and $ws.Range("G3").Value2 -eq 0.01) "a feed typed over after the preview: Apply works it out again instead of writing"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 0, "-20%", $true, $true, "enter")
    Check ($t -like "*|True|op;tool;time;flips / part *;*|* Speeds raised: the flips here count cut time only*") "speeds too: the flips column is marked and the note says why ('$($t.Substring(0, [math]::Min(60, $t.Length)))...')"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "30:00", $true, $false, "enter", $ws.Range("A4"))
    Check ($t -like "*|Not applied yet - Apply, or Revert to follow the sheet again.||op 2*") "a plan not applied: the window stays on op 2 when op 7 is clicked"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "", $true, $false, "", $ws.Range("A4"))
    Check ($t -like "op 7 takes 5:30 now*|op 7  -  T3 FINISH  -  Finish OD|*") "no plan: the window follows to op 7 ('$t')"
    $t = $xl.Run("Planner.PlanSelfTest", $ws.Range("G3"), 1, "30:00", $true, $false, "enter,revert")
    Check ($t -like "* now (*|False|*") "Revert drops the plan ('$($t.Substring(0, [math]::Min(60, $t.Length)))...')"
    Check ($ws.Range("G3").Value2 -eq 0.01) "the window wrote nothing"

    # ---- Inspection & inserts: tool 1 inspects on ops 2 and 9 every 8:00 - 600 s and 900 s
    # of cut, 1 + 1 flips; the shortest edge time that keeps 2 flips is 7:31 (to 5 s: 7:35).
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("G3"), $false, 0)
    "        window: $t"
    Check ($t -like "Apply sets the new edge time on every op of T1*flips per part 2 -> 2  (the whole part 2 -> 2).|True|T1;CNMG 432;2;8:00;82%;7:35;2;Same flips - each edge cuts 0:25 less.;ticked|*") "even: T1, its insert, 2 flips, 8:00 now, last edge 82% - 7:35 keeps 2 flips"
    Check ($ws.Range("Q3").Text -eq "8:00") "the window wrote nothing yet"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("G3"), $false, 0, "", "apply")
    $xl.Calculate()
    Check ($t -like "*|Done: edge time set on ops 2, 9 (T1) - flips per part 2 -> 2. Ctrl+Z to undo.|*") "Apply says what it did, in the window ('$($t.Split('|')[3])')"
    Check ($ws.Range("Q3").Text -eq "7:35" -and $ws.Range("Q5").Text -eq "7:35" -and $ws.Range("Q6").Text -eq "") "applied: 7:35 on every op of T1 that inspects; op 11 (no inspection) left blank"
    Check (($ws.Range("U3").Value2 + $ws.Range("U5").Value2) -eq 2) "still 2 flips"
    Check ($t -like "*T1;CNMG 432;2;7:35;*;Already even - every edge cuts about the same.;*") "after Apply the window reads the sheet again: already even"
    $u = $xl.Run("Panel.UndoLast")
    Check ($u -eq "2 0" -and $ws.Range("Q3").Text -eq "8:00" -and $ws.Range("Q5").Text -eq "8:00") "one Ctrl+Z puts both back ('$u')"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 0, "", "set Q5=6:00,apply")
    Check ($t -like "*|The sheet changed since this was worked out - here it is again. Check it, then Apply.|*" -and $ws.Range("Q3").Text -eq "8:00" -and $ws.Range("Q5").Text -eq "6:00") "an edge time typed over after the list was made: Apply works it out again instead of writing"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3,A5"), $false, 1)
    Check ($t -like "*2 -> 1  (the whole part 2 -> 1). / 1 of them would run each edge longer than now - check the insert can take it.|True|T1;CNMG 432;2;8:00;82%;10:05;1;Edge runs 2:05 longer than now - check the insert can take it.;ticked|*") "one flip fewer: 10:05, and the longer edge is warned about"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 2)
    Check ($t -like "*|True|T1;CNMG 432;2;8:00;82%;5:05;3;Each edge cuts 2:55 less.;ticked|*") "one flip more: 5:05, 3 flips ('$($t.Split('|')[2])')"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $true, 0)
    Check ($t -like "*|T1;*;ticked // T3;CNMG 432;0;;;;;No flips - it does not stop to turn the insert.;||Every tool on the sheet") "every tool: T3 does not stop, nothing to tick ('$($t.Split('|')[2])')"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 0, "1")
    Check ($t -like "Tick the tools to change.|False|*;8:00;82%;7:35;2;*;|*") "unticked: Apply off"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 1, "", "", $ws.Range("A4"))
    Check ($t -like "*|Not applied yet - Apply, or Revert to follow the sheet again.|op 2  -  T1*") "a goal changed here: the window stays on op 2 when op 7 is clicked"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 0, "", "", $ws.Range("A4"))
    Check ($t -like "*|T3;*||op 7  -  T3 FINISH  -  Finish OD") "nothing changed here: it follows to op 7's tool ('$t')"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3"), $false, 0, "", "", $ws.Range("A5"))
    Check ($t -like "*|True|T1;CNMG 432;2;8:00;82%;7:35;2;*;ticked||op 9  -  T1 ROUGH  -  Rough face") "another op of the same tool: the heading follows, the list stays ('$($t.Split('|')[4])')"
    # A tool whose flips the toolpath fixes - stops every 5 in of cut: no edge time moves them.
    $u4 = $ws.Range("U4").Formula
    $ws.Range("X2").Value2 = "insp_do_stop"; $ws.Range("Y2").Value2 = "insp_dist_on"; $ws.Range("Z2").Value2 = "insp_dist"
    $ws.Range("X3").Value2 = 1; $ws.Range("X4").Value2 = 1; $ws.Range("X5").Value2 = 1; $ws.Range("X6").Value2 = 0
    $ws.Range("Y4").Value2 = 1; $ws.Range("Z4").Value2 = 5
    $ws.Range("U4").Formula = "=2"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A4"), $false, 0)
    Check ($t -like "Nothing to change*|False|T3;CNMG 432;2;;;;;Flips come from stops in the toolpath (every 5 in of cut) - change them in Mastercam.;|*") "fixed by the toolpath: says so plainly, no suggestion ('$t')"
    $t = $xl.Run("Planner.InspectSelfTest", $ws.Range("A3,A4"), $false, 0)
    Check ($t -like "*|True|T1;CNMG 432;2;8:00;82%;7:35;2;*;ticked // T3;*;Flips come from stops in the toolpath*") "beside it, T1 still plans on its inspecting ops"
    [void] $ws.Range("X2:Z6").ClearContents()
    $ws.Range("U4").Formula = $u4
    $xl.Calculate()
    Check ($ws.Range("U4").Value2 -eq 0 -and $ws.Range("Q3").Text -eq "8:00" -and $ws.Range("G3").Value2 -eq 0.01) "the inspection window wrote nothing it was not asked to"

    # ---- Copy from op, into every other op of the tool.
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G3:I3"), "#0")
    Check ($t -like "4 cell(s) will change,  2 already that.*Into ops 9, 11.|True|2") "only op 2's row selected: copies into tool 1's other ops 9 and 11 ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G3:I3"), "#0", 0)
    Check ($t -like "Nothing to copy into*|False|2") "into the selected rows, with only op 2's row selected: nothing to do ('$t')"
    $t = $xl.Run("ParamTable.EditListSelfTest", "copy", $ws.Range("G3:I3"), "#0", 1)
    Check ($t -like "4#op 9    feed    0.012  ->  0.01") "the window lists each cell old -> new ('$t')"
    $r = $xl.Run("Planner.ToolCopyCells", $ws.Range("G3:I3"), 2)
    Check ($r -eq "6 0" -and $ws.Range("G5").Value2 -eq 0.01 -and $ws.Range("I6").Value2 -eq 200 -and $ws.Range("G4").Value2 -eq 0.008) "copied into ops 9 and 11, not tool 3's op 7 ($r)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)
    # Units: a feed is not copied into a row whose feed is per min, unless per rev / per min goes too.
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("H4"), "per min")
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G4"), "#0")
    Check ($t -like "0 cell(s) will change,  1 not copied - the units differ*|False|2") "op 2's per rev feed into op 7's per min feed: refused ('$t')"
    $t = $xl.Run("ParamTable.EditListSelfTest", "copy", $ws.Range("G4"), "#0")
    Check ($t -eq "1#op 7    feed    not copied - op 2 is per rev, op 7 is per min - select the per rev / per min column too") "and listed with why ('$t')"
    $t = $xl.Run("ParamTable.EditListSelfTest", "copy", $ws.Range("H4"), "#0")
    Check ($t -eq "1#op 7    feed_mode    not copied - op 7's feed is a per min value - select the feed too") "per rev / per min alone, without the feed: refused too ('$t')"
    $t = $xl.Run("ParamTable.EditWindowSelfTest", "copy", $ws.Range("G4:H4"), "#0")
    Check ($t -like "2 cell(s) will change.*Into op 7.|True|2") "the feed with its per rev / per min: copied ('$t')"
    $r = $xl.Run("ParamTable.CopyCells", $ws.Range("G4"), 2)
    Check ($r -eq "0 1" -and $ws.Range("G4").Value2 -eq 0.008) "and the copy itself refuses it ($r)"
    [void] $xl.Run("ParamTable.RevertCells", $all, $false)

    # ---- The dumped values are found by column NAME: a column inserted on the sheet does not
    # make every cell after it look edited.
    [void] $xl.Run("ParamTable.SetCells", $ws.Range("G3"), "0.011")
    [void] $ws.Columns.Item(4).Insert()
    $n = $xl.Run("Planner.EditCount")
    $c = $xl.Run("ParamTable.ListChanges")
    $back = $xl.Run("ParamTable.RevertCells", $ws.Range("A3:X6"), $false)
    Check ($n -eq 1 -and $c -eq 1 -and $back -eq 1 -and $ws.Range("H3").Value2 -eq 0.01) "a column inserted: 1 edit found, listed and reverted ($n, $c, $back)"
    [void] $ws.Columns.Item(4).Delete()
    [void] $ws.Activate()
    Check ($ws.Range("G3").Value2 -eq 0.01 -and $ws.Range("D3").Text -eq "Rough OD") "the column taken out again"

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
    $kept = $xl.Run("Planner.AutoSaved")
    $d = @($xl.Run("Planner.ScenarioDetail", $kept))
    Check ($kept -like "(before restore *)" -and $d.Count -eq 2 -and @($xl.Run("Planner.ScenarioNames")).Count -eq 3) "the edits the restore wiped are kept first, as '$kept' ($($d.Count) cells)"
    $r = $xl.Run("Planner.RestoreScenario", "fast")
    Check ($xl.Run("Planner.AutoSaved") -eq "" -and @($xl.Run("Planner.ScenarioNames")).Count -eq 3) "restoring again: the sheet's edits are in 'fast' already - nothing more kept"
    [void] $xl.Run("Planner.DeleteScenario", $kept)
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
