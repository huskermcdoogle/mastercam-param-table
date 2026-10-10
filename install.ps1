<#
    Install the Parameter Table Tool from a package folder - no build tools, no elevation.
    Either layout works:
      - the release zip's: ParamTable.ft beside a ParamTable\ folder (the DLL, help\,
        SHA256SUMS.txt, BUILD-INFO.txt) - run this from the folder holding ParamTable.ft;
      - a FLAT folder: the DLL, the function table, SHA256SUMS.txt (and help\) side by side.

    This is what goes to another machine. deploy.ps1 is the developer's copy step and
    expects the build tree; this one expects only what package.ps1 puts in the folder.

    It checks that the files arrived intact before it copies anything: a file that was
    truncated or altered in transit fails here, not inside Mastercam.
#>

[CmdletBinding()]
param([string] $Here = "")

$ErrorActionPreference = "Stop"

# Resolved in the BODY: $PSScriptRoot is empty while a param default is evaluated under -File.
if (-not $Here) { $Here = Split-Path -Parent $MyInvocation.MyCommand.Path }

if (Get-Process Mastercam -ErrorAction SilentlyContinue) {
    Write-Error "Mastercam is running - it holds the DLL open. Close it, then install."
    exit 1
}

# Where the files are: the release layout (ParamTable\ beside the .ft) or flat.
$pkg = if (Test-Path -LiteralPath (Join-Path $Here "ParamTable\ParamTable.dll")) { Join-Path $Here "ParamTable" } else { $Here }

# ---- 1. The files are what was packaged.
$sums = Join-Path $pkg "SHA256SUMS.txt"
if (-not (Test-Path -LiteralPath $sums)) { Write-Error "SHA256SUMS.txt is missing - cannot verify the files."; exit 1 }

foreach ($line in Get-Content -LiteralPath $sums) {
    if ($line -notmatch '^([0-9A-Fa-f]{64})\s+(.+)$') { continue }
    $want = $Matches[1].ToUpper(); $name = $Matches[2].Trim()
    # Names are relative to the folder holding ParamTable.ft; in a flat folder, by file name.
    $path = Join-Path $Here $name
    if (-not (Test-Path -LiteralPath $path)) { $path = Join-Path $pkg ($name -replace '^ParamTable\\', '') }
    if (-not (Test-Path -LiteralPath $path)) { Write-Error "$name is missing."; exit 1 }
    $got = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpper()
    if ($got -ne $want) { Write-Error "$name does not match its recorded hash - it was altered or damaged in transit. Not installing."; exit 1 }
    Write-Host ("  ok  {0}  {1}" -f $got.Substring(0, 12), $name)
}

# ---- 2. Where it goes.
$key = "HKCU:\SOFTWARE\CNC Software\Mastercam 2026"
$userDir = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).UserDir
if (-not $userDir -or -not (Test-Path -LiteralPath $userDir)) {
    Write-Error "Could not find Mastercam 2026's UserDir in the registry ($key). Is Mastercam 2026 installed for this user?"
    exit 1
}
$addins = Join-Path $userDir "Add-Ins"
$sub    = Join-Path $addins "ParamTable"
foreach ($d in @($addins, $sub)) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

$ft = if (Test-Path -LiteralPath (Join-Path $Here "ParamTable.ft")) { Join-Path $Here "ParamTable.ft" } else { Join-Path $pkg "ParamTable.ft" }
Copy-Item -LiteralPath $ft -Destination (Join-Path $addins "ParamTable.ft") -Force
Copy-Item -LiteralPath (Join-Path $pkg "ParamTable.dll") -Destination (Join-Path $sub "ParamTable.dll") -Force

# The user manual (help\), replacing an older copy.
$help = Join-Path $pkg "help"
if (Test-Path -LiteralPath (Join-Path $help "index.html")) {
    $helpDst = Join-Path $sub "help"
    if (Test-Path -LiteralPath $helpDst) { Remove-Item -LiteralPath $helpDst -Recurse -Force }
    Copy-Item -LiteralPath $help -Destination $helpDst -Recurse
}

Write-Host ""
Write-Host "installed:"
Write-Host ("  {0}" -f (Join-Path $addins "ParamTable.ft"))
Write-Host ("  {0}" -f (Join-Path $sub "ParamTable.dll"))
if (Test-Path -LiteralPath (Join-Path $sub "help\index.html")) { Write-Host ("  {0}" -f (Join-Path $sub "help\")) }

# ---- 3. Does this machine match what it was built against.
Write-Host ""
$mc  = 'C:\Program Files\Mastercam 2026\Mastercam.dll'
$mfc = 'C:\Windows\System32\mfc140u.dll'
if (Test-Path $mc)  { Write-Host ("this machine's Mastercam.dll : {0}" -f (Get-Item $mc).VersionInfo.FileVersion) }
if (Test-Path $mfc) { Write-Host ("this machine's mfc140u.dll   : {0}   (needs 14.40 or newer)" -f (Get-Item $mfc).VersionInfo.FileVersion) }
$info = Join-Path $pkg "BUILD-INFO.txt"
if (Test-Path -LiteralPath $info) { Write-Host ""; Get-Content -LiteralPath $info | ForEach-Object { Write-Host "  $_" } }

Write-Host ""
Write-Host "Start Mastercam. The four functions are under Customize > 'Parameter Table Tool'."
