<#
    Compile the workbook macros: every vba\*.bas (modules), the forms
    (vba\*.vb) and vba\Sheet1.cls (the main sheet's events) into res\vbaProject.bin - the compiled block an .xlsm carries. The add-in
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

# PT_TEST_OUT (when set) keeps one checkout's build out of another's way.
$tmp = Join-Path $(if ($env:PT_TEST_OUT) { $env:PT_TEST_OUT } else { $env:TEMP }) "ParamTableVba"
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

# The VBA editor imports ANSI text with CRLF line ends.
$modules = foreach ($m in (Get-ChildItem -LiteralPath (Join-Path $Root "vba") -Filter *.bas | ForEach-Object BaseName)) {
    $bas = Join-Path $tmp "$m.bas"
    $text = (Get-Content -LiteralPath (Join-Path $Root "vba\$m.bas") -Raw) -replace "`r?`n", "`r`n"
    [IO.File]::WriteAllText($bas, $text, [Text.Encoding]::GetEncoding(1252))
    $bas
}
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
    foreach ($bas in $modules) { [void] $vbp.VBComponents.Import($bas) }
    # The forms: named controls made here; their look and behaviour are all in vba\<name>.vb.
    function Add-Form ($name, $controls) {
        $frm = $vbp.VBComponents.Add(3)                  # vbext_ct_MSForm
        $frm.Name = $name
        foreach ($c in $controls) { [void] $frm.Designer.Controls.Add($c[0], $c[1], $true) }
        $code = (Get-Content -LiteralPath (Join-Path $Root "vba\$name.vb") -Raw) -replace "`r?`n", "`r`n"
        $fm = $frm.CodeModule
        if ($fm.CountOfLines -gt 0) { $fm.DeleteLines(1, $fm.CountOfLines) }
        $fm.AddFromString($code)
    }
    Add-Form "TextEditor" @(@("Forms.Label.1", "lblOps"), @("Forms.CommandButton.1", "btnPrev"), @("Forms.CommandButton.1", "btnNext"),
                            @("Forms.CommandButton.1", "btnHelp"), @("Forms.MultiPage.1", "mpg"), @("Forms.Label.1", "lblStatus"),
                            @("Forms.CommandButton.1", "btnRevert"), @("Forms.CommandButton.1", "btnApply"), @("Forms.CommandButton.1", "btnClose"))
    # Its tabs are the pages of mpg, each holding its own boxes: (page, control, name).
    $pages = $vbp.VBComponents.Item("TextEditor").Designer.Controls.Item("mpg").Pages
    while ($pages.Count -lt 3) { [void] $pages.Add() }
    foreach ($c in @(@(0, "Forms.Label.1", "lblCmtHint"), @(0, "Forms.TextBox.1", "txtComment"), @(0, "Forms.Label.1", "lblCmtCount"),
                     @(0, "Forms.Label.1", "lblCmtNow"),
                     @(1, "Forms.Label.1", "lblInspHint"), @(1, "Forms.TextBox.1", "txtInsp"), @(1, "Forms.Label.1", "lblInspCount"),
                     @(1, "Forms.Label.1", "lblQuick"), @(1, "Forms.CommandButton.1", "btnQ1"), @(1, "Forms.CommandButton.1", "btnQ2"),
                     @(1, "Forms.CommandButton.1", "btnQ3"), @(1, "Forms.CommandButton.1", "btnQ4"), @(1, "Forms.Label.1", "lblQuickNote"),
                     @(1, "Forms.CheckBox.1", "chkInspOn"), @(1, "Forms.Label.1", "lblInspNow"),
                     @(2, "Forms.Label.1", "lblManOp"), @(2, "Forms.ComboBox.1", "cboOp"), @(2, "Forms.Label.1", "lblManHint"),
                     @(2, "Forms.TextBox.1", "txtManual"), @(2, "Forms.Label.1", "lblManCount"))) {
        [void] $pages.Item($c[0]).Controls.Add($c[1], $c[2], $true)
    }
    Add-Form "CoolantPicker" (@(@("Forms.Label.1", "lblOps"), @("Forms.CommandButton.1", "btnPrev"), @("Forms.CommandButton.1", "btnNext"),
                                @("Forms.CommandButton.1", "btnHelp"), @("Forms.Label.1", "lblV9Head"), @("Forms.Label.1", "lblV9Hint"),
                                @("Forms.Label.1", "lblV9Now"), @("Forms.Label.1", "linSep"), @("Forms.Label.1", "lblXHead"),
                                @("Forms.Label.1", "lblXHint"), @("Forms.Label.1", "lblXB"), @("Forms.Label.1", "lblXW"),
                                @("Forms.Label.1", "lblXA"), @("Forms.Label.1", "lblXNow"), @("Forms.Label.1", "lblXNote"),
                                @("Forms.Label.1", "lblNone"), @("Forms.Label.1", "lblStatus"), @("Forms.CommandButton.1", "btnRevert"),
                                @("Forms.CommandButton.1", "btnApply"), @("Forms.CommandButton.1", "btnClose")) +
                              @(1..4 | ForEach-Object { , @("Forms.OptionButton.1", "optV$_") }) +
                              @(1..12 | ForEach-Object { @("Forms.Label.1", "lblX$_"), @("Forms.CheckBox.1", "chkB$_"),
                                                         @("Forms.CheckBox.1", "chkW$_"), @("Forms.CheckBox.1", "chkA$_") }))
    Add-Form "EditBox" @(@("Forms.Label.1", "lblInfo"), @("Forms.Label.1", "lblPrompt"), @("Forms.ComboBox.1", "cbo"),
                         @("Forms.Label.1", "lblPreview"), @("Forms.ListBox.1", "lstDetail"),
                         @("Forms.CommandButton.1", "btnOK"), @("Forms.CommandButton.1", "btnCancel"))
    Add-Form "PlanBox" @(@("Forms.Label.1", "lblInfo"), @("Forms.Label.1", "lblScope"), @("Forms.ComboBox.1", "cboScope"),
                         @("Forms.Label.1", "lblTarget"), @("Forms.ComboBox.1", "cboTarget"), @("Forms.Label.1", "lblHow"),
                         @("Forms.ComboBox.1", "cboHow"), @("Forms.CheckBox.1", "chk1"), @("Forms.CheckBox.1", "chk2"),
                         @("Forms.ListBox.1", "lst"), @("Forms.Label.1", "lblSummary"),
                         @("Forms.CommandButton.1", "btnOK"), @("Forms.CommandButton.1", "btnCancel"))
    Add-Form "Calculator" (@(@("Forms.Label.1", "lblInfo"), @("Forms.Label.1", "lblNote"), @("Forms.CheckBox.1", "chkMetric"),
                             @("Forms.CommandButton.1", "btnClose"), @("Forms.Label.1", "lblChip"), @("Forms.Label.1", "lblThin"),
                             @("Forms.CommandButton.1", "btnUseFeed"), @("Forms.CommandButton.1", "btnSetRow")) +
                           @("Dia", "Surf", "Rpm", "Rev", "Min", "IC", "Ap", "Hex", "Fn" | ForEach-Object { @("Forms.Label.1", "lbl$_"), @("Forms.TextBox.1", "txt$_") }))

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

# The UserForm's MSForms reference records C:\Users\<you>\...\MSForms.exd - take the name out.
python (Join-Path $Root "tools\scrub_vba.py") $out $env:USERNAME
if ($LASTEXITCODE) { throw "scrub_vba failed" }

"compiled -> res\vbaProject.bin ({0:N0} bytes)" -f (Get-Item -LiteralPath $out).Length
