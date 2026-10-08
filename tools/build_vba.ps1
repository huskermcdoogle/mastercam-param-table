<#
    Compile the workbook macros: vba\ParamTable.bas (a module) and vba\Sheet1.cls (the main
    sheet's events) into res\vbaProject.bin - the compiled block an .xlsm carries. The add-in
    embeds that file, and writes it into a dump when macros are switched on.

    Run after changing anything in vba\, then rebuild the add-in. Needs Excel, with "Trust
    access to the VBA project object model" enabled (Excel > Options > Trust Center).

    The sheet must keep the code name Sheet1 and the workbook ThisWorkbook - the add-in
    writes those names into every macro workbook so the compiled code finds its sheet.
#>
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = "Stop"
if (-not $Root) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$tmp = Join-Path $env:TEMP "ParamTableVba"
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

# The VBA editor imports ANSI text with CRLF line ends.
$bas = Join-Path $tmp "ParamTable.bas"
$text = (Get-Content -LiteralPath (Join-Path $Root "vba\ParamTable.bas") -Raw) -replace "`r?`n", "`r`n"
[IO.File]::WriteAllText($bas, $text, [Text.Encoding]::GetEncoding(1252))
$sheetCode = (Get-Content -LiteralPath (Join-Path $Root "vba\Sheet1.cls") -Raw) -replace "`r?`n", "`r`n"

$xlsm = Join-Path $tmp "vba.xlsm"
if (Test-Path -LiteralPath $xlsm) { Remove-Item -LiteralPath $xlsm -Force }

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Lathe params"

    # Code names read as empty until the project has been opened - open it first.
    $vbp = $wb.VBProject
    $names = @($vbp.VBComponents | ForEach-Object { $_.Name })
    if ($names -notcontains "Sheet1" -or $names -notcontains "ThisWorkbook") {
        throw "the VBA project has '$($names -join ', ')', expected Sheet1 and ThisWorkbook"
    }
    [void] $vbp.VBComponents.Import($bas)
    $cm = $vbp.VBComponents.Item("Sheet1").CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($sheetCode)

    $wb.SaveAs($xlsm, 52)        # xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($xlsm)
try {
    $entry = $zip.GetEntry("xl/vbaProject.bin")
    if ($null -eq $entry) { throw "the saved workbook has no xl/vbaProject.bin" }
    $out = Join-Path $Root "res\vbaProject.bin"
    $in = $entry.Open()
    $fs = [IO.File]::Create($out)
    try { $in.CopyTo($fs) } finally { $fs.Close(); $in.Close() }
} finally { $zip.Dispose() }

"compiled -> res\vbaProject.bin ({0:N0} bytes)" -f (Get-Item -LiteralPath $out).Length
