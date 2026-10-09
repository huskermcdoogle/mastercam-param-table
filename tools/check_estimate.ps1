<#
    Real-Excel check of the live estimate formula: opens estimate.xlsx (written by
    tests\estimate_test.exe), makes each scenario's cell edits, and compares the
    formula's seconds with what the C++ model expects (estimate_expect.txt).

    Usage:  powershell -File tools\check_estimate.ps1 <dir holding estimate.xlsx>
#>
param([string] $Dir = "")
$ErrorActionPreference = "Stop"
if (-not $Dir) { $Dir = if ($env:PT_TEST_OUT) { $env:PT_TEST_OUT } else { Join-Path $env:TEMP "ParamTableTests" } }
$file = Join-Path $Dir "estimate.xlsx"
$expect = Get-Content (Join-Path $Dir "estimate_expect.txt")
$failed = 0
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    foreach ($line in $expect) {
        $edits, $want = $line -split '\|'
        $wb = $xl.Workbooks.Open($file); $ws = $wb.Worksheets.Item("Lathe params")
        if ($edits) {
            foreach ($e in ($edits -split ';')) {
                $ref, $val = $e -split '=', 2
                $d = 0.0
                if ([double]::TryParse($val, [ref]$d)) { $ws.Range($ref).Value2 = $d } else { $ws.Range($ref).Formula = [string]$val }
            }
        }
        $xl.Calculate()
        $got = [double]$ws.Range('G3').Value2
        $ok = [math]::Abs($got - [double]$want) -lt 0.01
        if (-not $ok) { $failed++ }
        "{0,-28} excel {1,12:N3}  model {2,12:N3}  {3}" -f $(if ($edits) { $edits } else { '(as dumped)' }), $got, [double]$want, $(if ($ok) { 'ok' } else { 'FAIL' })
        $wb.Close($false)
    }
} finally {
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
if ($failed) { "check_estimate: $failed FAILED"; exit 1 }
"check_estimate: all scenarios match"
