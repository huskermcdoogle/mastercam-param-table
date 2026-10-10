<#
    Real-Excel check of the Summary page: opens summary_sample.xlsx (tests\summary_test.exe)
    and checks that it is the first sheet and the one that opens, that nothing on it is an
    error, that its ranked lists follow edits on the main sheet (and survive a sort), and
    that the batch quantity and the Tools page's cost per insert give the insert cost.
    Then history_sample.xlsx (tests\history_test.exe): the "From the machine" cells against
    the estimate, the History sheet after the Summary, and what a dump keeps once Excel
    has saved it.

    Usage:  powershell -File tools\check_summary.ps1 [<dir holding summary_sample.xlsx>]
#>
param([string] $Dir = "")
$ErrorActionPreference = "Stop"
if (-not $Dir) { $Dir = if ($env:PT_TEST_OUT) { $env:PT_TEST_OUT } else { Join-Path $env:TEMP "ParamTableTests" } }
$src = Join-Path $Dir "summary_sample.xlsx"
$f = Join-Path $Dir "summary_check.xlsx"
Copy-Item -LiteralPath $src -Destination $f -Force

$failed = 0
function Check ($ok, $what) {
    if ($ok) { "  ok    $what" } else { "  FAIL  $what"; $script:failed++ }
}
# The Summary row whose column A or B reads $text.
function RowOf ($ws, $text) {
    for ($r = 1; $r -le 120; $r++) {
        if ($ws.Cells.Item($r, 1).Text -eq $text -or $ws.Cells.Item($r, 2).Text -eq $text) { return $r }
    }
    return 0
}

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($f)
    $xl.Calculate()
    $sm = $wb.Worksheets.Item("Summary")
    $ws = $wb.Worksheets.Item("Lathe params")
    $tl = $wb.Worksheets.Item("Tools")
    $names = ($wb.Worksheets | ForEach-Object { $_.Name }) -join ', '
    Check ($wb.Worksheets.Item(1).Name -eq "Summary" -and $wb.ActiveSheet.Name -eq "Summary") "Summary is first and opens ($names)"
    Check ($wb.Worksheets.Item("Dumped").Visible -eq 0) "Dumped still hidden"
    Check (-not $sm.Application.ActiveWindow.DisplayGridlines) "no gridlines on the Summary"

    $errors = @()
    foreach ($c in $sm.UsedRange.Cells) { if ("$($c.Text)" -like "#*" -and "$($c.Text)" -ne "#") { $errors += $c.Address(0, 0) + "=" + $c.Text } }
    Check ($errors.Count -eq 0) "no error values on the Summary ($($errors -join ' '))"
    $man = RowOf $sm "Open the user manual - how every tool works, with examples"
    Check ($man -gt 0 -and $sm.Cells.Item($man, 1).Formula -like '=HYPERLINK("C:\ParamTable\help\index.html",*') "a link to the user manual ($man)"
    Check ($wb.Names.Item("PT_Help").RefersTo -eq '="C:\ParamTable\help"') "the manual's folder recorded for the macros ($($wb.Names.Item('PT_Help').RefersTo))"

    $lg = RowOf $sm "Longest operations (estimated time, now)"
    Check ($lg -gt 0 -and $sm.Cells.Item($lg + 2, 2).Text -eq "Rough OD" -and $sm.Cells.Item($lg + 3, 2).Text -eq "Groove") "longest: Rough OD, Groove (a tie, sheet order) - '$($sm.Cells.Item($lg + 2, 2).Text)', '$($sm.Cells.Item($lg + 3, 2).Text)'"
    Check ($sm.Cells.Item($lg + 2, 6).Text -eq "0:05:00" -and $sm.Cells.Item($lg + 2, 7).Text -eq "41.7%") "time 0:05:00, share 41.7% ('$($sm.Cells.Item($lg + 2, 6).Text)', '$($sm.Cells.Item($lg + 2, 7).Text)')"
    Check ($sm.Cells.Item($lg + 5, 2).Text -eq "") "an op with no time is not listed"

    $cy = RowOf $sm "Cycle time (estimate)"
    $ws.Range("F5").Value2 = 600               # op 7: 300 s -> 600 s
    $xl.Calculate()
    Check ($sm.Cells.Item($lg + 2, 2).Text -eq "Groove" -and $sm.Cells.Item($lg + 2, 3).Text -eq "7") "an edit re-ranks: Groove (op 7) is first ('$($sm.Cells.Item($lg + 2, 2).Text)')"
    Check ($sm.Cells.Item($cy, 3).Text -eq "0:12:00" -and $sm.Cells.Item($cy, 4).Text -eq "0:17:00" -and $sm.Cells.Item($cy, 5).Text -eq "+0:05:00") "cycle time: as dumped 0:12:00, now 0:17:00, +0:05:00 ('$($sm.Cells.Item($cy, 4).Text)', '$($sm.Cells.Item($cy, 5).Text)')"

    # Sort the main sheet by comment: the lists follow the rows, not the positions.
    [void] $ws.Range("A3:J6").Sort($ws.Range("D3"), 1)
    $xl.Calculate()
    Check ($ws.Range("D3").Text -eq "Finish OD" -and $sm.Cells.Item($lg + 2, 2).Text -eq "Groove" -and $sm.Cells.Item($lg + 2, 3).Text -eq "7") "after a sort, still Groove op 7 first"
    $sm.Cells.Item($lg + 2, 3).Select() | Out-Null
    Check ($sm.Cells.Item($lg + 2, 3).Formula -like "*HYPERLINK(*") "the op number is a link ('$($sm.Cells.Item($lg + 2, 3).Formula)')"

    $tr = RowOf $sm "Insert flips by tool (per part, now)"
    Check ($sm.Cells.Item($tr + 2, 2).Text -eq "OD ROUGH" -and $sm.Cells.Item($tr + 2, 5).Text -eq "2" -and $sm.Cells.Item($tr + 3, 2).Text -eq "OD FINISH") "tools by flips: OD ROUGH 2, OD FINISH"
    $ws.Range("H3:H6").Value2 = 0              # no flips anywhere
    $xl.Calculate()
    Check ($sm.Cells.Item($tr + 2, 2).Text -eq "No insert flips counted on this part." -and $sm.Cells.Item($tr + 3, 2).Text -eq "") "no flips: says so ('$($sm.Cells.Item($tr + 2, 2).Text)')"
    $ws.Range("A3:J6").Sort($ws.Range("A3"), 1) | Out-Null
    $ws.Range("H3").Value2 = 2; $ws.Range("H4").Value2 = 0.5; $ws.Range("H5").Value2 = 1
    $xl.Calculate()

    $ar = RowOf $sm "Most air cutting (cut time with no stock in the way)"
    Check ($sm.Cells.Item($ar + 2, 2).Text -eq "Groove" -and $sm.Cells.Item($ar + 2, 6).Text -eq "0:01:40" -and $sm.Cells.Item($ar + 2, 7).Text -eq "50%") "air: Groove 0:01:40 50% ('$($sm.Cells.Item($ar + 2, 6).Text)')"

    # Insert cost: 4-edge CNMG at 12.50 (2 flips), 2-edge VNMG at 8.00 (1.5 flips).
    $tl.Range("G6").Value2 = 12.5
    $tl.Range("G7").Value2 = 8
    $xl.Calculate()
    Check ([math]::Abs($tl.Range("H6").Value2 - 6.25) -lt 1e-9 -and [math]::Abs($tl.Range("H7").Value2 - 6) -lt 1e-9) "cost per part per insert: 6.25, 6 ($($tl.Range('H6').Value2), $($tl.Range('H7').Value2))"
    Check ($tl.Range("G6").Text -match "12[.,]50") "cost shown as currency ('$($tl.Range('G6').Text)', $($tl.Range('G6').NumberFormat))"
    $qr = RowOf $sm "Batch quantity (type it in)"
    $pp = RowOf $sm "Insert cost per part (average)"
    $pb = RowOf $sm "Insert cost per batch (whole inserts)"
    Check ([math]::Abs($sm.Cells.Item($pp, 3).Value2 - 12.25) -lt 1e-9) "insert cost per part 12.25 ($($sm.Cells.Item($pp, 3).Value2))"
    $sm.Cells.Item($qr, 3).Value2 = 10
    $xl.Calculate()
    # 10 parts: CNMG 20 edges = 5 inserts x 12.50; VNMG 15 edges = 8 inserts x 8.00.
    Check ([math]::Abs($sm.Cells.Item($pb, 3).Value2 - 126.5) -lt 1e-9) "batch of 10: 62.50 + 64.00 = 126.50 ($($sm.Cells.Item($pb, 3).Value2))"
    $bt = RowOf $sm "Cycle time per batch (now)"
    Check ($sm.Cells.Item($bt, 3).Text -eq "2:50:00") "cycle time per batch of 10: 2:50:00 ('$($sm.Cells.Item($bt, 3).Text)')"

    # The filter on the main sheet still belongs to the main sheet.
    Check ($ws.AutoFilterMode) "the main sheet keeps its filter"

    $wb.Close($false)

    # ---- history_sample.xlsx (tests\history_test.exe): the Summary's "From the machine"
    # cells - typed times taken apart past 24 hours, the estimate against them - and the
    # History sheet after the Summary, printable; then what a dump keeps from it once
    # Excel has saved it.
    $hf = Join-Path $Dir "history_check.xlsx"
    Copy-Item -LiteralPath (Join-Path $Dir "history_sample.xlsx") -Destination $hf -Force
    $wb = $xl.Workbooks.Open($hf)
    $xl.Calculate()
    $sm = $wb.Worksheets.Item("Summary")
    $hs = $wb.Worksheets.Item("History")
    $ws = $wb.Worksheets.Item("Lathe params")
    $names = ($wb.Worksheets | ForEach-Object { $_.Name }) -join ', '
    Check ($wb.Worksheets.Item(1).Name -eq "Summary" -and $wb.Worksheets.Item(2).Name -eq "History" -and $wb.Worksheets.Item(3).Name -eq "Lathe params" -and $wb.ActiveSheet.Name -eq "Summary") "Summary, History, Lathe params - opens on the Summary ($names)"
    Check ($ws.AutoFilterMode) "the main sheet keeps its filter behind the History"
    $errors = @()
    foreach ($s in @($sm, $hs)) { foreach ($c in $s.UsedRange.Cells) { if ("$($c.Text)" -like "#*" -and "$($c.Text)" -ne "#") { $errors += $s.Name + "!" + $c.Address(0, 0) + "=" + $c.Text } } }
    Check ($errors.Count -eq 0) "no error values on the Summary or the History ($($errors -join ' '))"

    $cy = RowOf $sm "Actual cycle time per part (m:ss or h:mm:ss)"
    $dt = RowOf $sm "Date it was measured"
    $pr = RowOf $sm "Parts run"
    $ir = RowOf $sm "Inserts used for that run (total, if known)"
    $ag = RowOf $sm "Machine against the estimate"
    $ip = RowOf $sm "Inserts per part: machine, estimate"
    Check ($cy -gt 0 -and $sm.Cells.Item($cy, 3).Text -eq "0:13:10" -and $sm.Cells.Item($cy, 3).NumberFormat -eq "@" -and $sm.Cells.Item($cy, 9).Value2 -eq 790) "the last measurement shown again, as Text, 790 s in the working column ($($sm.Cells.Item($cy, 9).Value2))"
    Check ($sm.Cells.Item($dt, 3).Text -eq "2026-10-12" -and $sm.Cells.Item($dt, 3).Validation.Type -eq 4) "the date shown yyyy-mm-dd, checked as a date ('$($sm.Cells.Item($dt, 3).Text)')"
    Check ($sm.Cells.Item($ag, 3).Text -eq "+0:01:10" -and $sm.Cells.Item($ag, 4).Text -eq "+9.7%" -and $sm.Cells.Item($ag, 5).Text -eq "The machine is slower than the estimate.") "machine 0:13:10 against the estimate 0:12:00: +0:01:10, +9.7%, slower ('$($sm.Cells.Item($ag, 3).Text)' '$($sm.Cells.Item($ag, 4).Text)')"
    Check ($sm.Cells.Item($ip, 3).Text -eq "0.78" -and $sm.Cells.Item($ip, 4).Text -eq "1.25") "inserts a part: machine 31 / 40 = 0.78, estimate 2/4 + 1.5/2 = 1.25 ('$($sm.Cells.Item($ip, 3).Text)', '$($sm.Cells.Item($ip, 4).Text)')"
    $cases = @(@("4:30", 270, "-0:07:30", "-62.5%", "faster"), @("1:02:30", 3750, "+0:50:30", "+420.8%", "slower"),
               @("26:32:07", 95527, "+26:20:07", "+13167.6%", "slower"), @("75:30", 4530, "+1:03:30", "+529.2%", "slower"),
               @("720", 720, "+0:00:00", "0.0%", "matches"))
    foreach ($k in $cases) {
        $sm.Cells.Item($cy, 3).Value2 = $k[0]
        $xl.Calculate()
        Check ($sm.Cells.Item($cy, 9).Value2 -eq $k[1] -and $sm.Cells.Item($ag, 3).Text -eq $k[2] -and $sm.Cells.Item($ag, 4).Text -eq $k[3] -and $sm.Cells.Item($ag, 5).Text -like "*$($k[4])*" -and $sm.Cells.Item($cy, 3).Validation.Value) "typed $($k[0]): $($k[1]) s, $($k[2]) $($k[3]) ('$($sm.Cells.Item($cy, 9).Value2)' '$($sm.Cells.Item($ag, 3).Text)' '$($sm.Cells.Item($ag, 4).Text)')"
    }
    $sm.Cells.Item($cy, 3).Value2 = "soon"
    $xl.Calculate()
    Check (-not $sm.Cells.Item($cy, 3).Validation.Value -and $sm.Cells.Item($cy, 9).Text -eq "" -and $sm.Cells.Item($ag, 5).Text -eq "Type the actual time above.") "not a time: refused by the cell's rule, nothing worked out"
    $ws.Range("F5").Value2 = 600               # an edit: op 7 300 s -> 600 s, the estimate 0:17:00
    $sm.Cells.Item($cy, 3).Value2 = "16:00"
    $xl.Calculate()
    Check ($sm.Cells.Item($ag, 3).Text -eq "-0:01:00" -and $sm.Cells.Item($ag, 5).Text -eq "The machine is faster than the estimate.") "the estimate follows the edits: 16:00 against 0:17:00 is -0:01:00 ('$($sm.Cells.Item($ag, 3).Text)')"

    # The History sheet: titled, its records, printable.
    Check ($hs.Range("A1").Text -eq "History - SAMPLE.mcam") "the History page names the part ('$($hs.Range('A1').Text)')"
    $dr = 0; for ($q = 1; $q -le 40; $q++) { if ($hs.Cells.Item($q, 2).Text -like "Dumped 4 operations*") { $dr = $q; break } }
    Check ($dr -gt 0 -and $hs.Cells.Item($dr, 3).Text -eq "0:12:00" -and $hs.Cells.Item($dr, 8).Text -eq "12.25") "the dump's line: 0:12:00, insert cost 12.25"
    Check ($hs.PageSetup.Orientation -eq 2 -and $hs.PageSetup.FitToPagesWide -eq 1 -and $hs.PageSetup.PrintTitleRows -ne "") "printable: landscape, one page wide, the heading on every page ($($hs.PageSetup.Orientation), $($hs.PageSetup.FitToPagesWide), $($hs.PageSetup.PrintTitleRows))"

    # Saved by Excel, then read the way a dump reads it.
    $sm.Cells.Item($cy, 3).Value2 = "26:32:07"
    $sm.Cells.Item($dt, 3).Value2 = 46309
    $sm.Cells.Item($pr, 3).Value2 = 40
    $sm.Cells.Item($ir, 3).Value2 = 31
    $sm.Cells.Item($ir + 1, 3).Value2 = "typed in Excel"
    $saved = Join-Path $Dir "history_saved.xlsx"
    if (Test-Path -LiteralPath $saved) { Remove-Item -LiteralPath $saved -Force }
    $wb.SaveCopyAs($saved)
    $wb.Close($false)
    $harvest = @(& (Join-Path $Dir "history_test.exe") --harvest $saved)
    Check ($harvest -contains "2026-10-13 07:05 | actual | measured: 2026-10-14 | cycle: 26:32:07 | parts: 40 | inserts: 31 | estimate: 0:17:00 | note: typed in Excel") "a dump keeps what Excel saved ($($harvest -join '; '))"
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_summary: $failed FAILED"; exit 1 }
"check_summary: all checks passed"
