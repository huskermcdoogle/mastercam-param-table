<#
    Real-Excel check of the workbook macros: opens macro_sample.xlsm (tests\macro_test.exe)
    with macros enabled and drives them - two-way links, cross-checks, the cell rules on
    bulk writes, scale, copy, the change list and revert.

    Usage:  powershell -File tools\check_macros.ps1 [<dir holding macro_sample.xlsm>]
#>
param([string] $Dir = "")
$ErrorActionPreference = "Stop"
if (-not $Dir) { $Dir = Join-Path $env:TEMP "ParamTableTests" }
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
    "opened: $($xl.Windows.Item(1).Caption)  sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ', ') +
        "  modules: " + (($wb.VBProject.VBComponents | ForEach-Object { $_.Name }) -join ', ')

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

    $wb.Close($false)
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_macros: $failed FAILED"; exit 1 }
"check_macros: all checks passed"
