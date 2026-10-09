<#
    Real-Excel check of the Summary page: opens summary_sample.xlsx (tests\summary_test.exe)
    and checks that it is the first sheet and the one that opens, that nothing on it is an
    error, that its ranked lists follow edits on the main sheet (and survive a sort), and
    that the batch quantity and the Tools page's cost per insert give the insert cost.

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
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_summary: $failed FAILED"; exit 1 }
"check_summary: all checks passed"
