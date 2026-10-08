<#
    Draws the add-in's command icons (16 and 32 px PNGs, transparent) into res\Resources:
    a spreadsheet with an arrow going in (dump_*, blue) or coming back out (load_*, orange).
    16 px is placed pixel by pixel so it stays sharp; 32 px is drawn.

    Usage:  powershell -File tools\make_icons.ps1 [-Preview <png>]
#>
param([string] $Preview = "")
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$Root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $Root "res\Resources"

$C = @{
    E = [Drawing.Color]::FromArgb(255, 64, 72, 82)      # sheet outline
    G = [Drawing.Color]::FromArgb(255, 33, 115, 70)     # header row
    W = [Drawing.Color]::White
    L = [Drawing.Color]::FromArgb(255, 186, 194, 204)   # grid lines
    B = [Drawing.Color]::FromArgb(255, 24, 104, 214)    # dump arrow
    O = [Drawing.Color]::FromArgb(255, 234, 112, 8)     # load arrow
}

# --- 16 px, pixel by pixel ------------------------------------------------------------------
function Small ($load) {
    $bmp = New-Object Drawing.Bitmap 16, 16, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $px = @{}
    for ($y = 1; $y -le 14; $y++) {
        for ($x = 5; $x -le 14; $x++) {
            $k = "W"
            if ($x -eq 5 -or $x -eq 14 -or $y -eq 1 -or $y -eq 14) { $k = "E" }
            elseif ($y -le 3) { $k = "G" }
            elseif ($x -eq 10 -or $y -eq 7 -or $y -eq 10 -or $y -eq 12) { $k = "L" }
            $px["$x,$y"] = $k
        }
    }
    # Arrow on row 8: shaft rows 7..9, head 7 tall.
    $arrow = @()
    if ($load) {
        $tip = 0; $d = -1; $shaft = 4..11
    } else {
        $tip = 11; $d = 1; $shaft = 0..7
    }
    foreach ($x in $shaft) { foreach ($y in 7..9) { $arrow += , @($x, $y) } }
    for ($i = 0; $i -le 3; $i++) {
        foreach ($y in (8 - $i)..(8 + $i)) { $arrow += , @(($tip - $d * $i), $y) }
    }
    $ink = if ($load) { "O" } else { "B" }
    $set = @{}; foreach ($p in $arrow) { $set["$($p[0]),$($p[1])"] = 1 }
    # White ring where the arrow crosses the sheet, so it reads against the grid.
    foreach ($p in $arrow) {
        foreach ($n in @(@(1, 0), @(-1, 0), @(0, 1), @(0, -1))) {
            $q = "$($p[0] + $n[0]),$($p[1] + $n[1])"
            if (-not $set.ContainsKey($q) -and $px.ContainsKey($q) -and $px[$q] -ne "E") { $px[$q] = "W" }
        }
    }
    foreach ($q in $set.Keys) { $px[$q] = $ink }
    foreach ($q in $px.Keys) {
        $xy = $q.Split(","); $bmp.SetPixel([int]$xy[0], [int]$xy[1], $C[$px[$q]])
    }
    $bmp
}

# --- 32 px, drawn ---------------------------------------------------------------------------
function Large ($load) {
    $bmp = New-Object Drawing.Bitmap 32, 32, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.Clear([Drawing.Color]::Transparent)
    $brush = { param($k) New-Object Drawing.SolidBrush $C[$k] }
    $g.FillRectangle((& $brush E), 10, 2, 21, 28)
    $g.FillRectangle((& $brush W), 11, 3, 19, 26)
    $g.FillRectangle((& $brush G), 11, 3, 19, 5)
    foreach ($x in 17, 24) { $g.FillRectangle((& $brush L), $x, 8, 1, 21) }
    foreach ($y in 13, 19, 24) { $g.FillRectangle((& $brush L), 11, $y, 19, 1) }

    if ($load) { $tail = 24; $tip = 0 } else { $tail = 0; $tip = 23 }
    $d = [math]::Sign($tip - $tail); $head = $tip - $d * 9
    $pts = @(
        @($tail, 13), @($head, 13), @($head, 7), @($tip, 16), @($head, 25), @($head, 19), @($tail, 19)
    ) | ForEach-Object { New-Object Drawing.PointF $_[0], $_[1] }
    $g.SmoothingMode = "AntiAlias"; $g.PixelOffsetMode = "Half"
    $g.SetClip((New-Object Drawing.Rectangle 11, 3, 19, 26))
    $ring = New-Object Drawing.Pen $C.W, 4
    $ring.LineJoin = "Round"
    $g.DrawPolygon($ring, $pts)
    $g.ResetClip()
    $g.FillPolygon((& $brush $(if ($load) { "O" } else { "B" })), $pts)
    $g.Dispose()
    $bmp
}

$made = @()
foreach ($kind in "dump", "load") {
    $l = $kind -eq "load"
    foreach ($pair in @(@("small", (Small $l)), @("large", (Large $l)))) {
        $name = "{0}_{1}.png" -f $kind, $pair[0]
        $pair[1].Save((Join-Path $out $name), [Drawing.Imaging.ImageFormat]::Png)
        $made += , $pair[1]
        "wrote res\Resources\$name"
    }
}

# A blown-up sheet of all four, plus actual size, on white and on a dark ribbon.
if ($Preview) {
    $p = New-Object Drawing.Bitmap 1040, 560
    $g = [Drawing.Graphics]::FromImage($p)
    $g.InterpolationMode = "NearestNeighbor"; $g.PixelOffsetMode = "Half"
    $g.FillRectangle([Drawing.Brushes]::White, 0, 0, 1040, 280)
    $g.FillRectangle((New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(255, 45, 45, 48))), 0, 280, 1040, 280)
    for ($band = 0; $band -lt 2; $band++) {
        $x = 20
        foreach ($m in $made) {
            $z = if ($m.Width -eq 16) { 12 } else { 6 }
            $g.DrawImage($m, $x, 20 + 280 * $band, $m.Width * $z, $m.Width * $z)
            $g.DrawImage($m, $x + 200, 220 + 280 * $band, $m.Width, $m.Width)
            $x += 260
        }
    }
    $p.Save($Preview, [Drawing.Imaging.ImageFormat]::Png)
    "preview -> $Preview"
}
