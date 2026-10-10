<#
    Package the current build into a folder (and a zip of it) that can be carried to
    another machine. Contents:

        ParamTable.ft                    the function table
        ParamTable\ParamTable.dll        the build
        ParamTable\help\                 the user manual (plain web pages; index.html)
        ParamTable\SHA256SUMS.txt        the hash of every file here, to check nothing changed in transit
        ParamTable\BUILD-INFO.txt        which source and toolset it came from

    A hash proves the files are unchanged since packaging. It says nothing about whether
    the code is safe - that is what review is for, and the source is in git.
#>

[CmdletBinding()]
param([string] $Root = "", [string] $Configuration = "Release")

$ErrorActionPreference = "Stop"

# Resolved in the BODY: $PSScriptRoot is empty while a param default is evaluated under -File.
if (-not $Root) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }

$dll = Join-Path $Root "build\x64\$Configuration\ParamTable.dll"
foreach ($f in @($dll, (Join-Path $Root "ParamTable.ft"))) {
    if (-not (Test-Path -LiteralPath $f)) { Write-Error "not found: $f - build first"; exit 1 }
}

$stamp = Get-Date -Format "yyyyMMdd-HHmm"
$out   = Join-Path $Root "dist\ParamTable-$stamp"
# LAID OUT AS IT SITS INSIDE Documents\My Mastercam 2026\Mastercam\Add-Ins: the .ft at the
# top and everything else in a ParamTable\ folder, so the contents of this folder are
# dropped straight into Add-Ins with no folder to make by hand.
$sub = Join-Path $out "ParamTable"
New-Item -ItemType Directory -Force -Path $sub | Out-Null

Copy-Item -LiteralPath $dll -Destination (Join-Path $sub "ParamTable.dll")
Copy-Item -LiteralPath (Join-Path $Root "ParamTable.ft") -Destination $out
Copy-Item -LiteralPath (Join-Path $Root "LICENSE") -Destination (Join-Path $sub "LICENSE.txt")

# The user manual: help\ beside the DLL, as Mastercam's "Parameter Table - user manual" and
# the workbook's links expect it. The list of pictures still to make stays in the repo.
$helpSrc = Join-Path $Root "help"
if (-not (Test-Path -LiteralPath (Join-Path $helpSrc "index.html"))) { Write-Error "help\index.html not found - the manual is missing"; exit 1 }
Copy-Item -LiteralPath $helpSrc -Destination (Join-Path $sub "help") -Recurse
Get-ChildItem -LiteralPath (Join-Path $sub "help") -Recurse -Filter "NEEDED.txt" | Remove-Item -Force

# ---- Build info: the commit, whether the tree was clean, and what it was built with.
$commit = (& git -C $Root rev-parse --short HEAD 2>$null)
$dirty  = (& git -C $Root status --porcelain 2>$null)
$mc     = 'C:\Program Files\Mastercam 2026\Mastercam.dll'
$toolset = (Get-ChildItem 'C:\Program Files\Microsoft Visual Studio\18\Community\VC\Tools\MSVC' -Directory -Filter '14.40*' -ErrorAction SilentlyContinue | Select-Object -First 1).Name

$lines = @(
    "Parameter Table Tool",
    ("packaged        : {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm")),
    ("source commit   : {0}{1}" -f $commit, $(if ($dirty) { "  (UNCOMMITTED CHANGES in the tree)" } else { "" })),
    ("MSVC toolset    : v143 {0}" -f $toolset),
    ("built against   : Mastercam 2026 SDK; Mastercam.dll {0}" -f $(if (Test-Path $mc) { (Get-Item $mc).VersionInfo.FileVersion } else { "unknown" })),
    "imports         : Mastercam.dll, MCCore.dll, MCTool.dll, CHookAPI.dll, mfc140u.dll, the C runtime, kernel32, user32 - no networking"
)
Set-Content -LiteralPath (Join-Path $sub "BUILD-INFO.txt") -Value $lines -Encoding UTF8

# ---- Hashes LAST, over the files as they will be shipped - the manual's pages too, so a
# page altered in transit is caught like the DLL would be.
$names = @('ParamTable.ft', 'ParamTable\ParamTable.dll', 'ParamTable\BUILD-INFO.txt', 'ParamTable\LICENSE.txt')
$names += Get-ChildItem -LiteralPath (Join-Path $sub "help") -Recurse -File | Sort-Object FullName |
    ForEach-Object { 'ParamTable\help\' + $_.FullName.Substring((Join-Path $sub "help").Length + 1) }
$sums = foreach ($n in $names) {
    "{0}  {1}" -f (Get-FileHash -LiteralPath (Join-Path $out $n) -Algorithm SHA256).Hash, $n
}
Set-Content -LiteralPath (Join-Path $sub "SHA256SUMS.txt") -Value $sums -Encoding ASCII

$zip = "$out.zip"

# A folder inside Dropbox is being synced the moment it is written, and Dropbox holds a
# file open while it does - which fails the zip with "in use by another process". A few
# short retries clear it.
for ($try = 1; $try -le 6; $try++) {
    try {
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
        Compress-Archive -Path (Join-Path $out '*') -DestinationPath $zip -ErrorAction Stop
        break
    } catch {
        if ($try -eq 6) { throw }
        Start-Sleep -Seconds 3
    }
}

Write-Host ""
Get-Content -LiteralPath (Join-Path $sub "SHA256SUMS.txt")
Write-Host ""
Write-Host ("folder : {0}" -f $out)
Write-Host ("zip    : {0}  ({1:N0} bytes)" -f $zip, (Get-Item $zip).Length)
