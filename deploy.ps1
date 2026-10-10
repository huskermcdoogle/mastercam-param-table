<#
    Deploy the Parameter Table Tool into Mastercam 2026's per-user Add-Ins folder.

    NO ELEVATION. Unlike a C-Hook dropped in Program Files\...\chooks, the
    per-user Add-Ins folder is the user's own, which is the whole reason the
    2026 template installs there. The folder comes from the registry the way the
    template's own post-build does (HKCU\SOFTWARE\CNC Software\Mastercam 2026,
    value UserDir) - it is under OneDrive on this machine, so it is not a path
    that can be guessed.

    Mastercam holds the DLL open for the life of the process and loads it only at
    startup, so it must be CLOSED to deploy, and a fresh start is what picks the
    new one up.

    Layout, matching the FUNC_DLL line in ParamTable.ft:
        <UserDir>Add-Ins\ParamTable.ft
        <UserDir>Add-Ins\ParamTable\ParamTable.dll
        <UserDir>Add-Ins\ParamTable\help\        the user manual (index.html ...)
#>

[CmdletBinding()]
param(
    [string] $Root = "",
    [string] $Configuration = "Release"
)

$ErrorActionPreference = "Stop"

# Resolved in the BODY: $PSScriptRoot is empty while a param default is evaluated under -File.
if (-not $Root) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }

if (Get-Process Mastercam -ErrorAction SilentlyContinue) {
    Write-Error "Mastercam is running - it holds the DLL open. Close it, then deploy."
    exit 1
}

$dll = Join-Path $Root "build\x64\$Configuration\ParamTable.dll"
$ft  = Join-Path $Root "ParamTable.ft"
foreach ($f in @($dll, $ft)) {
    if (-not (Test-Path -LiteralPath $f)) { Write-Error "not found: $f - build first"; exit 1 }
}

$key = "HKCU:\SOFTWARE\CNC Software\Mastercam 2026"
$userDir = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).UserDir
if (-not $userDir -or -not (Test-Path -LiteralPath $userDir)) {
    Write-Error "Could not find Mastercam 2026's UserDir in the registry ($key)."
    exit 1
}

$addins = Join-Path $userDir "Add-Ins"
if (-not (Test-Path -LiteralPath $addins)) {
    New-Item -ItemType Directory -Path $addins | Out-Null
}
$sub = Join-Path $addins "ParamTable"
if (-not (Test-Path -LiteralPath $sub)) {
    New-Item -ItemType Directory -Path $sub | Out-Null
}

Copy-Item -LiteralPath $ft  -Destination (Join-Path $addins "ParamTable.ft") -Force
Copy-Item -LiteralPath $dll -Destination (Join-Path $sub "ParamTable.dll") -Force

# The user manual beside the DLL (ParamTable\help): Mastercam's "Parameter Table - user
# manual", the Summary's link, the ribbon's User manual and every window's ? open it.
# The old copy goes first, so a page that was removed does not linger.
$helpSrc = Join-Path $Root "help"
if (Test-Path -LiteralPath $helpSrc) {
    $helpDst = Join-Path $sub "help"
    if (Test-Path -LiteralPath $helpDst) { Remove-Item -LiteralPath $helpDst -Recurse -Force }
    Copy-Item -LiteralPath $helpSrc -Destination $helpDst -Recurse
    # The list of pictures still to make is for the repo, not for users.
    Get-ChildItem -LiteralPath $helpDst -Recurse -Filter "NEEDED.txt" | Remove-Item -Force
    Write-Host ("user manual          -> {0} ({1} files)" -f $helpDst, @(Get-ChildItem -LiteralPath $helpDst -Recurse -File).Count)
}

# The copy succeeding is not proof the right bytes arrived: two builds of the
# same size an hour apart look identical in a log line.
$built    = Get-Item $dll
$deployed = Get-Item (Join-Path $sub "ParamTable.dll")
if ($built.Length -ne $deployed.Length) {
    Write-Error "deployed DLL does not match the build"
    exit 1
}

Write-Host ("deployed {0:N0} bytes -> {1}" -f $deployed.Length, $sub)
Write-Host ("function table       -> {0}" -f (Join-Path $addins "ParamTable.ft"))
Write-Host ""
Write-Host "Start Mastercam. The Parameter Table tab is on the ribbon (ParamTableParamTable-ribbon.log says how the start went)."
