<#
    Make the user manual's pictures (help\images\*.png) WITHOUT screenshots.

    Sheet pictures: a COPY of a sample workbook is opened in a hidden Excel of our own;
    a range goes through Range.CopyPicture into a chart object and Chart.Export to a PNG.
    Window pictures: the macro window is opened in that hidden Excel and only THAT window
    is captured, with the Win32 PrintWindow call on its handle - found by its class
    (ThunderDFrame, a VBA UserForm) AND by belonging to our Excel's process. Nothing else
    on the screen is ever read.

    Then tools\make_help.py puts each picture into its page (a page names the picture it
    wants; while the file is missing the page shows a marked placeholder instead).

    Sources (copied to a temp folder first; the originals are never opened):
      - the newest tests\Oil Spool Head_2026_lathe_params_*.xlsm (a public Mastercam Tech
        Exchange sample part) - in this checkout's tests\ or the main checkout's;
      - else macro_sample.xlsm from tests\run_tests.bat (%PT_TEST_OUT%, else
        %TEMP%\ParamTableTests).
    The copy gets the CURRENT macros (res\vbaProject.bin) so the windows are the ones in
    this tree. Any cell holding a folder path is cut to the file name before a picture.

    Run it from a PowerShell prompt in the repo (PowerShell -File may be blocked):
        & .\tools\make_help_images.ps1                 # every picture
        & .\tools\make_help_images.ps1 -Only win-set   # just some (names: -List)
        & .\tools\make_help_images.ps1 -List
    Re-runnable: it overwrites its own PNGs. Needs Excel, and "Trust access to the VBA
    project object model" for the window pictures (as tools\build_vba.ps1 does).

    NOTE: Range.CopyPicture goes through the Windows clipboard. The clipboard's text is
    put back at the end; anything else on it (a picture) is lost.
    A window may show on the screen for a moment while it is captured.
    Never commit the workbooks themselves - only the PNGs in help\images.
#>
[CmdletBinding()]
param(
    [string[]] $Only = @(),
    [string] $Workbook = "",
    [string] $SampleDir = "",
    [string] $Out = "",
    [switch] $List,
    [switch] $KeepTemp,
    [double] $Scale = 1.5
)
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Out) { $Out = Join-Path $Root "help\images" }

# ============================================================ what to make
#
# SHEET pictures: Sheet and Range. In a range or a callout, {Some text} is the row of the
# first cell in columns A-B holding exactly that text, {Some text}+2 two rows below.
#   Show      only these columns (by their name in the heading row); the rest hidden for
#             the picture - a wide sheet made readable.
#   Edits     take the picture after the example edits ($Edits below).
#   Folded    the main sheet's column groups folded to their +/- handles.
#   Macro     run first (a sheet the macros make); "ribbon:<button id>" runs that button's
#             macro the watched way (as for windows, below).
#   Range     "used:14" = the first 14 rows of what the sheet holds.
#   Callouts  @( @(number, cell, where) ) - cell: "D10", "B{text}", "row2:feed" (row 2 of
#             the column named feed), "op7:feed" (op 7's row), "head:Insert" (that heading),
#             "inserts:G" (column G of the first insert row). where: l / r / c (left, right,
#             middle of the cell); lb / rb at the bottom of a tall cell.
#
# WINDOW pictures: Window = how to open it -
#   "ribbon:<button id>"  runs that ribbon button's own macro (vba\ribbon.xml), as a click does;
#   "vba:<code>"          VBA that sets f to a window, filled in; it is shown modeless.
#                         Reach the workbook's macros ONLY through Application.Run and f As
#                         Object: a macro renamed later is then a plain error, caught - never a
#                         compile error (which would open the VBA editor on the screen).
#   Select    @{ Ops = op numbers; Cols = column names } - the cells selected first.
#   Callouts  @( @(number, control name, l / r) ) - only for "vba:" windows; a control
#             that is not there is skipped.
#   Pending   the window is being rebuilt: the picture is made - check it before use.

$Shots = @(
    @{ Name = "sheet-summary-top"; Sheet = "Summary"; Range = "A1:F{Insert cost per batch (whole inserts)}+1"; Edits = $true
       Callouts = @(@(1, "B{Cycle time (estimate)}", "r"), @(2, "B{Batch quantity (type it in)}", "r"), @(3, "B{Insert cost per batch (whole inserts)}", "r")) },
    @{ Name = "sheet-summary-lists"; Sheet = "Summary"; Range = "A{Longest operations (estimated time, now)}:G{Longest operations (estimated time, now)}+7"; Edits = $true
       Callouts = @(@(1, "C{Longest operations (estimated time, now)}+2", "r"), @(2, "G{Longest operations (estimated time, now)}+2", "l")) },
    @{ Name = "sheet-main"; Sheet = "Lathe params"; Range = "A1:A8"; Edits = $true
       Show = @("op_idn", "type", "tool", "comment", "changes", "feed", "feed_mode", "speed", "speed_mode", "coolant")
       Callouts = @(@(1, "row1:speed_mode", "r"), @(2, "row2:speed", "r"), @(3, "op7:feed", "l"), @(4, "op7:changes", "l")) },
    @{ Name = "sheet-estimate"; Sheet = "Lathe params"; Range = "A1:A8"; Edits = $true
       Show = @("op_idn", "comment", "feed", "speed", "cycle_time", "est_cycle_time", "time_change")
       Callouts = @(@(1, "op7:feed", "l"), @(2, "op7:cycle_time", "r"), @(3, "op7:est_cycle_time", "r")) },
    @{ Name = "sheet-tools"; Sheet = "Tools"; Range = "A1:A3"
       Show = @("Tool", "Name", "Used by", "Insert", "Flips / part", "Cut time / part", "Inspection", "Edge life (fallback)", "Insert size (IC)", "Entering angle", "Picture")
       Callouts = @(@(1, "row2:Insert", "rb"), @(2, "row2:Flips / part", "lb"), @(3, "row2:Edge life (fallback)", "rb")) },
    @{ Name = "sheet-inserts"; Sheet = "Tools"; Range = "inserts"
       Callouts = @(@(1, "inserts:D", "l"), @(2, "inserts:G", "l"), @(3, "inserts:I", "l"), @(4, "inserts:J", "l")) },
    @{ Name = "sheet-scenarios"; Sheet = "Scenarios"; Range = "A1:H{Parameters changed}+4"; Macro = "scenarios" },
    @{ Name = "sheet-changes"; Sheet = "Changes"; Range = "A1:F4"; Edits = $true; Macro = "ParamTable.ListChanges"; Widen = 3 },

    @{ Name = "win-set"; Window = 'vba:Set f = Application.Run("ParamTable.EditWindow", "set", Application.Run("ParamTable.DataCells", Selection)): f.TypeIn "0.012"'
       Select = @{ Ops = @(1, 7, 8); Cols = @("feed") }
       Callouts = @(@(1, "cbo", "r"), @(2, "lblPreview", "r"), @(3, "lstDetail", "l"), @(4, "btnOK", "l")) },
    @{ Name = "win-scenario-save"; Window = 'vba:Set f = Application.Run("ParamTable.EditWindow", "scensave", Nothing): f.TypeIn "Faster bore"'; Edits = $true
       Callouts = @(@(1, "cbo", "r"), @(2, "btnOK", "l")) },
    @{ Name = "win-scale"; Window = 'vba:Set f = Application.Run("ParamTable.EditWindow", "scale", Application.Run("ParamTable.DataCells", Selection)): f.TypeIn "+10%"'
       Select = @{ Ops = @(1, 7, 8); Cols = @("feed") }
       Callouts = @(@(1, "cbo", "r"), @(2, "lblPreview", "r"), @(3, "btnOK", "l")) },
    # The op windows that stay open beside the sheet: opened by their ribbon button, then filled in.
    @{ Name = "win-text"; Window = 'vba:Application.Run "Panel.RbText", Nothing: Set f = PtForm("TextEditor"): f.Act "tab=insp": f.Act "!btnQ1"'
       Select = @{ Ops = @(7, 8); Cols = @("comment") }
       Callouts = @(@(1, "lblOps", "r"), @(2, "mpg", "l"), @(3, "btnApply", "l")) },
    @{ Name = "win-coolant"; Window = 'vba:Application.Run "Panel.RbCoolant", Nothing: Set f = PtForm("CoolantPicker"): f.Act "v9=Flood"'
       Select = @{ Ops = @(1, 2, 7); Cols = @("coolant") }
       Callouts = @(@(1, "lblOps", "r"), @(2, "btnApply", "l")) },
    @{ Name = "win-calculator"; Window = 'vba:Application.Run "Panel.RbSpeed", Nothing: Set f = PtForm("Calculator")'
       Select = @{ Ops = @(7); Cols = @("feed") }
       Callouts = @(@(1, "cboOp", "r"), @(2, "mpg", "l")) },
    @{ Name = "win-calculator-chip"; Window = 'vba:Application.Run "Panel.RbSpeed", Nothing: Set f = PtForm("Calculator"): f.Controls("mpg").Value = 1: f.TypeIn "txtKr", "45"'
       Select = @{ Ops = @(1); Cols = @("feed") }
       Callouts = @(, @(1, "mpg", "l")) },
    @{ Name = "win-inspection"; Window = 'vba:Application.Run "Panel.RbInspect", Nothing: Set f = PtForm("InspectBox"): f.SetInputs 0, 1, ""'
       Select = @{ Ops = @(1, 2, 3, 7); Cols = @("feed") }
       Callouts = @(@(1, "cboGoal", "r"), @(2, "lst", "l"), @(3, "btnApply", "l")) },
    @{ Name = "win-target"; Window = 'vba:Application.Run "Panel.RbTarget", Nothing: Set f = PtForm("PlanBox"): f.TypeTarget "-10%": f.Press "preview"'
       Select = @{ Ops = @(7, 8); Cols = @("feed") }
       Callouts = @(@(1, "cboTarget", "r"), @(2, "chkMain", "l"), @(3, "lst", "l"), @(4, "btnApply", "l")) },
    # Copy from op: op 7 tuned first (feed 0.012, 220 SFM), then copied into op 8. Last: it changes op 7.
    @{ Name = "win-copy"; Window = 'vba:Dim ws As Object, r7 As Long: Set ws = Application.Run("ParamTable.MainSheet"): r7 = Application.Run("Panel.RowOfOp", 7): ws.Cells(r7, Application.Run("ParamTable.ColOf", "feed")).Value = 0.012: ws.Cells(r7, Application.Run("ParamTable.ColOf", "speed")).Value = 220: Set f = Application.Run("ParamTable.EditWindow", "copy", Application.Run("ParamTable.DataCells", Selection))'
       Select = @{ Ops = @(7, 8); Cols = @("feed", "speed") }
       Callouts = @(@(1, "cbo", "r"), @(2, "cboInto", "r"), @(3, "lblPreview", "r"), @(4, "btnOK", "l")) },

    # Sheets the new commands make (Macro "ribbon:<id>" runs the button's macro first, watched).
    @{ Name = "sheet-check"; Sheet = "Program check"; Range = "used:14"; Macro = "ribbon:ptCheck" },
    @{ Name = "sheet-report"; Sheet = "Change report"; Range = "used:30"; Edits = $true; Macro = "ribbon:ptReport" },
    @{ Name = "sheet-slowest"; Sheet = "Slowest ops"; Range = "used:10"; Macro = "ribbon:ptSlowest" }
)

# The edits the pictures show, typed as a person would (through the sheet's own events):
# the bore feeds (the manual's running example) and - in a scenario only - the facing
# speed. Op numbers are the Oil Spool Head's; in another sample the first rows are used.
$Edits = @(
    @{ Op = 7; Col = "feed"; Value = 0.012; Scenario = "Bore feed 0.012" },
    @{ Op = 8; Col = "feed"; Value = 0.012; Scenario = "Bore feed 0.012" },
    @{ Op = 1; Col = "speed"; Value = 250; Scenario = "Face 250 SFM"; ScenarioOnly = $true }
)

if ($List) {
    $Shots | ForEach-Object { "{0,-22} {1}{2}" -f $_.Name, $(if ($_.Window) { $_.Window } else { "sheet '" + $_.Sheet + "'" }), $(if ($_.Pending) { "   (window being rebuilt)" } else { "" }) }
    return
}
if ($Only.Count -gt 0) { $Shots = @($Shots | Where-Object { $Only -contains $_.Name }) }
if ($Shots.Count -eq 0) { throw "no picture by that name - see -List" }

# ============================================================ Win32: our windows only

$winSource = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class PtWin
{
    public delegate bool EnumProc (IntPtr h, IntPtr l);
    [StructLayout (LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout (LayoutKind.Sequential)] public struct POINT { public int X, Y; }

    [DllImport ("user32.dll")] static extern bool EnumWindows (EnumProc cb, IntPtr l);
    [DllImport ("user32.dll")] static extern uint GetWindowThreadProcessId (IntPtr h, out uint pid);
    [DllImport ("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName (IntPtr h, StringBuilder s, int n);
    [DllImport ("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText (IntPtr h, StringBuilder s, int n);
    [DllImport ("user32.dll")] static extern bool IsWindowVisible (IntPtr h);
    [DllImport ("user32.dll")] public static extern bool IsWindow (IntPtr h);
    [DllImport ("user32.dll")] static extern bool GetWindowRect (IntPtr h, out RECT r);
    [DllImport ("user32.dll")] static extern bool GetClientRect (IntPtr h, out RECT r);
    [DllImport ("user32.dll")] static extern bool ClientToScreen (IntPtr h, ref POINT p);
    [DllImport ("dwmapi.dll")] static extern int DwmGetWindowAttribute (IntPtr h, int a, out RECT r, int size);
    [DllImport ("user32.dll")] static extern bool PrintWindow (IntPtr h, IntPtr hdc, uint flags);
    [DllImport ("user32.dll")] static extern bool PostMessage (IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport ("user32.dll")] static extern bool SetProcessDPIAware ();

    public static void DpiAware () { try { SetProcessDPIAware (); } catch { } }

    public static uint ProcessOf (IntPtr h) { uint pid; GetWindowThreadProcessId (h, out pid); return pid; }
    public static string ClassOf (IntPtr h) { var s = new StringBuilder (256); GetClassName (h, s, 256); return s.ToString (); }
    public static string TextOf (IntPtr h) { var s = new StringBuilder (512); GetWindowText (h, s, 512); return s.ToString (); }

    /// The visible top-level windows of ONE process, of one class ("" = any).
    public static IntPtr[] Find (uint pid, string cls)
    {
        var found = new List<IntPtr> ();
        EnumWindows ((h, l) =>
        {
            if (IsWindowVisible (h) && ProcessOf (h) == pid && (cls == "" || ClassOf (h) == cls))
                found.Add (h);
            return true;
        }, IntPtr.Zero);
        return found.ToArray ();
    }

    public static void Close (IntPtr h) { PostMessage (h, 0x0010, IntPtr.Zero, IntPtr.Zero); }   // WM_CLOSE

    /// The window drawn by itself into a bitmap (PrintWindow, full content), cut to its
    /// visible frame (no invisible resize border).
    public static Bitmap Capture (IntPtr h)
    {
        RECT w, f;
        GetWindowRect (h, out w);
        if (DwmGetWindowAttribute (h, 9, out f, Marshal.SizeOf (typeof (RECT))) != 0) f = w;   // DWMWA_EXTENDED_FRAME_BOUNDS
        int ww = w.R - w.L, wh = w.B - w.T;
        var full = new Bitmap (ww, wh, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage (full))
        {
            IntPtr dc = g.GetHdc ();
            PrintWindow (h, dc, 2);                      // PW_RENDERFULLCONTENT
            g.ReleaseHdc (dc);
        }
        var crop = new Rectangle (f.L - w.L, f.T - w.T, f.R - f.L, f.B - f.T);
        if (crop.Width <= 0 || crop.Height <= 0 || crop.Right > ww || crop.Bottom > wh) return full;
        var cut = full.Clone (crop, PixelFormat.Format32bppArgb);
        full.Dispose ();
        return cut;
    }

    /// Where the client area starts inside the captured frame, and its size - to put
    /// callouts on controls (whose places are in points inside the client area).
    public static int[] ClientPlace (IntPtr h)
    {
        RECT w, f, c;
        GetWindowRect (h, out w);
        if (DwmGetWindowAttribute (h, 9, out f, Marshal.SizeOf (typeof (RECT))) != 0) f = w;
        GetClientRect (h, out c);
        var p = new POINT ();
        ClientToScreen (h, ref p);
        return new int[] { p.X - f.L, p.Y - f.T, c.R - c.L, c.B - c.T };
    }

    /// A numbered callout: a filled circle with a white number, on a picture.
    public static void Callout (Bitmap b, int n, float x, float y, float d)
    {
        using (var g = Graphics.FromImage (b))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            var r = new RectangleF (x - d / 2, y - d / 2, d, d);
            using (var fill = new SolidBrush (Color.FromArgb (255, 194, 24, 91)))
            using (var edge = new Pen (Color.White, Math.Max (2f, d / 12)))
            using (var font = new Font ("Segoe UI", d * 0.55f, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                g.FillEllipse (fill, r);
                g.DrawEllipse (edge, r);
                g.DrawString (n.ToString (), font, Brushes.White, new RectangleF (r.X, r.Y + d * 0.03f, r.Width, r.Height), sf);
            }
        }
    }
}
'@
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition $winSource
[PtWin]::DpiAware()

# ============================================================ the sources

function Find-Sample {
    if ($Workbook) { return (Resolve-Path -LiteralPath $Workbook).Path }
    $dirs = @(Join-Path $Root "tests")
    try {
        # In a git worktree the test parts sit in the main checkout's tests\.
        $common = (& git -C $Root rev-parse --git-common-dir 2>$null)
        if ($common) {
            if (-not [IO.Path]::IsPathRooted($common)) { $common = Join-Path $Root $common }
            $main = Split-Path -Parent ([IO.Path]::GetFullPath($common))
            $dirs += Join-Path $main "tests"
        }
    } catch { }
    $newest = $dirs | Where-Object { Test-Path -LiteralPath $_ } |
        ForEach-Object { Get-ChildItem -LiteralPath $_ -Filter "Oil Spool Head_2026_lathe_params_*.xlsm" -ErrorAction SilentlyContinue } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) { return $newest.FullName }
    if (-not $SampleDir) { $SampleDir = if ($env:PT_TEST_OUT) { $env:PT_TEST_OUT } else { Join-Path $env:TEMP "ParamTableTests" } }
    $m = Join-Path $SampleDir "macro_sample.xlsm"
    if (Test-Path -LiteralPath $m) { return $m }
    throw "No sample workbook: no tests\Oil Spool Head_2026_lathe_params_*.xlsm and no $m (run tests\run_tests.bat)."
}

$source = Find-Sample
$tmp = Join-Path $env:TEMP ("PT_help_images_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
New-Item -ItemType Directory -Force -Path $Out | Out-Null
$copy = Join-Path $tmp "sample.xlsm"
Copy-Item -LiteralPath $source -Destination $copy
"source : {0}" -f (Split-Path -Leaf $source)

# The copy carries the macros of THIS tree: its xl/vbaProject.bin replaced by res\vbaProject.bin.
$vba = Join-Path $Root "res\vbaProject.bin"
if (Test-Path -LiteralPath $vba) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($copy, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $e = $zip.GetEntry("xl/vbaProject.bin")
        if ($e) {
            $e.Delete()
            [void] [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $vba, "xl/vbaProject.bin")
            "macros : this tree's res\vbaProject.bin"
        }
    } finally { $zip.Dispose() }
}

# The clipboard's text, to put back (CopyPicture goes through the clipboard).
$clipText = $null
try { $clipText = Get-Clipboard -Raw -ErrorAction Stop } catch { }

# ============================================================ our own hidden Excel

$script:xl = $null; $script:xlPid = 0; $script:wb = $null; $script:edited = $false; $script:dog = $null

# THE WATCHDOG. A VBA error in a hidden Excel still puts up a window you can see (a compile
# error opens the VBA editor and a "Microsoft Visual Basic for Applications" box) and then
# waits for ever. A background job watches OUR Excel's process only: a VBA error box or the
# VBA editor, or any box left up for 20 seconds, and it stops that process (and only that).
function Start-Watchdog([int] $xlPid) {
    Stop-Watchdog
    $script:dog = Start-Job -ArgumentList $xlPid, $winSource -ScriptBlock {
        param([int] $p, [string] $src)
        Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition $src
        $seen = @{}
        while (Get-Process -Id $p -ErrorAction SilentlyContinue) {
            $vbe = @([PtWin]::Find([uint32] $p, "wndclass_desked_gsk"))
            $boxes = @([PtWin]::Find([uint32] $p, "#32770"))
            $vbaBox = @($boxes | Where-Object { [PtWin]::TextOf($_) -like "Microsoft Visual Basic*" })
            $stuck = @($boxes | Where-Object { $k = "$_"; if (-not $seen[$k]) { $seen[$k] = Get-Date }; ((Get-Date) - $seen[$k]).TotalSeconds -gt 20 })
            if ($vbe.Count -or $vbaBox.Count -or $stuck.Count) {
                Stop-Process -Id $p -Force
                "stopped our Excel ($p): " + $(if ($vbe.Count -or $vbaBox.Count) { "a VBA error / the VBA editor came up" } else { "a box was left waiting" })
                break
            }
            Start-Sleep -Milliseconds 300
        }
    }
}

function Stop-Watchdog {
    if ($script:dog) {
        $said = Receive-Job $script:dog -ErrorAction SilentlyContinue
        if ($said) { "  watchdog: $said" }
        Stop-Job $script:dog -ErrorAction SilentlyContinue
        Remove-Job $script:dog -Force -ErrorAction SilentlyContinue
        $script:dog = $null
    }
}

function Start-Excel {
    $script:xl = New-Object -ComObject Excel.Application       # always a NEW instance of our own
    $script:xl.Visible = $false
    $script:xl.DisplayAlerts = $false
    $script:xl.AutomationSecurity = 1                           # macros run (in our copy)
    $script:xlPid = [PtWin]::ProcessOf([IntPtr] [long] $script:xl.Hwnd)
    if ($script:xlPid -eq 0) { throw "could not tell which process our Excel is" }
    Start-Watchdog $script:xlPid
    # Park our (hidden) Excel's window off the screen: the windows it opens centre on it.
    try { $script:xl.WindowState = -4143; $script:xl.Left = -4000; $script:xl.Top = -4000 } catch { }
    $script:wb = $script:xl.Workbooks.Open($copy)
    $script:edited = $false
    Clean-Paths
    # Every column group open. (The macros find a column with Range.Find on its values,
    # which skips HIDDEN cells - with "Working numbers" folded, as a dump starts, the
    # scenarios' totals read 0:00. Opening the groups keeps the pictures right.)
    try { [void] (Main-Sheet).Outline.ShowLevels(0, 2) } catch { }
    # The manual's folder, as a dump records it, for the "? How this works" links the macros
    # write (this checkout's help\ - the copy only).
    try { $script:wb.Names.Item("PT_Help").Delete() } catch { }
    [void] $script:wb.Names.Add("PT_Help", '="' + (Join-Path $Root "help") + '"', $false)
    # A very hidden sheet of our own, where the window pictures leave their notes.
    $ev = $script:xl.EnableEvents
    $script:xl.EnableEvents = $false
    $data = $script:wb.Worksheets.Add([Type]::Missing, $script:wb.Worksheets.Item($script:wb.Worksheets.Count))
    $data.Name = "PtShotData"
    $data.Visible = 2                                            # xlSheetVeryHidden
    (Main-Sheet).Activate()
    $script:xl.EnableEvents = $ev
}

# Our Excel still there? (The watchdog may have stopped it.) If not, a fresh one.
function Ensure-Excel {
    if (-not $script:xlPid -or -not (Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue)) {
        "        (our Excel was gone - starting a fresh one)"
        Stop-Excel -Kill
        Start-Excel
    }
}

function Stop-Excel([switch] $Kill) {
    # Every window of ours closed first, so nothing is left on the screen.
    if ($script:xlPid) {
        foreach ($h in [PtWin]::Find([uint32] $script:xlPid, "ThunderDFrame")) { [PtWin]::Close($h) }
        Start-Sleep -Milliseconds 300
    }
    if ($script:xl -and -not $Kill) {
        try { $script:wb.Close($false) } catch { }
        try { $script:xl.Quit() } catch { }
    }
    if ($script:xl) { try { [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($script:xl) } catch { } }
    $script:xl = $null; $script:wb = $null
    if ($script:xlPid) {
        # Only OUR Excel - the process this script started. No other Excel is touched.
        $p = Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq "EXCEL") {
            if (-not $Kill) { [void] $p.WaitForExit(15000) }
            if (-not $p.HasExited) { Stop-Process -Id $script:xlPid -Force }
        }
    }
    $script:xlPid = 0
    Stop-Watchdog
}

# No folder paths in a picture: a cell holding one keeps only the file name.
function Clean-Paths {
    $ev = $script:xl.EnableEvents
    $script:xl.EnableEvents = $false
    foreach ($ws in $script:wb.Worksheets) {
        $used = $ws.UsedRange
        for ($guard = 0; $guard -lt 200; $guard++) {
            $f = $used.Find(":\", [Type]::Missing, -4163, 2)        # in values, part of the cell
            if (-not $f) { break }
            $t = [string] $f.Text
            $clean = [regex]::Replace($t, '[A-Za-z]:\\(?:[^\\]*\\)*', '')
            if ($clean -eq $t) { break }
            Set-Value $f $clean
        }
    }
    $script:xl.EnableEvents = $ev
}

# The main sheet. Excel can be busy for a moment after a window closes (a COM call then
# comes back empty): asked again, a few times.
function Main-Sheet {
    for ($i = 0; $i -lt 20; $i++) {
        $sheets = $null
        try { $sheets = $script:wb.Worksheets } catch { }
        if ($sheets) { return $sheets.Item("Lathe params") }
        Start-Sleep -Milliseconds 300
    }
    $n = try { $script:xl.Workbooks.Count } catch { "?" }
    throw "Excel does not answer (workbook object $(if ($null -eq $script:wb) { 'gone' } else { 'there' }), $n workbook(s) open)"
}

# A column by its heading (row 2 on the main sheet, row 1 elsewhere), exact case. 0 = none.
function Col-Of([string] $name, $ws = $null) {
    if (-not $ws) { $ws = Main-Sheet }
    $row = if ($ws.Name -eq "Lathe params") { 2 } else { 1 }
    $last = $ws.Cells($row, $ws.Columns.Count).End(-4159).Column
    $vals = $ws.Range($ws.Cells($row, 1), $ws.Cells($row, $last)).Value2
    for ($c = 1; $c -le $last; $c++) { if ([string] $vals[1, $c] -ceq $name) { return $c } }
    return 0
}

function Row-Of($op) {
    $ws = Main-Sheet
    $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    for ($r = 3; $r -le $last; $r++) { if ("$($ws.Cells($r, 1).Value2)" -eq "$op") { return $r } }
    return 0
}

function Set-Edit($e, [switch] $Back) {
    $ws = Main-Sheet
    $r = Row-Of $e.Op
    if ($r -eq 0) { $r = 3 + [array]::IndexOf($Edits, $e) }
    $c = Col-Of $e.Col
    if ($c -eq 0) { return }
    $cell = $ws.Cells($r, $c)
    if ($Back) { Set-Value $cell $e.Old; return }
    if (-not $e.ContainsKey("Old")) { $e.Old = $cell.Value2 }
    Set-Value $cell $e.Value
}

# A cell's value. Not "$cell.Value2 = x": PowerShell keeps the COM call it bound first
# (to a text value) and then fails a number with "cannot cast Double to String".
function Set-Value($cell, $v) {
    [void] [System.__ComObject].InvokeMember("Value2", [Reflection.BindingFlags]::SetProperty, $null, $cell, @($v))
}

function Make-Edits {
    if ($script:edited) { return }
    $script:xl.EnableEvents = $true
    foreach ($e in $Edits) { if (-not $e.ScenarioOnly) { Set-Edit $e } }
    $script:xl.CalculateFull()
    $script:edited = $true
}

# Two scenarios saved (one per idea), then every edit on the sheet, then compared.
function Make-Scenarios {
    $script:xl.EnableEvents = $true
    foreach ($e in $Edits) { if ($script:edited -and $e.ContainsKey("Old")) { Set-Edit $e -Back } }
    $script:edited = $false
    foreach ($name in @($Edits | ForEach-Object { $_.Scenario } | Select-Object -Unique)) {
        $mine = @($Edits | Where-Object { $_.Scenario -eq $name })
        foreach ($e in $mine) { Set-Edit $e }
        $script:xl.CalculateFull()
        [void] $script:xl.Run("Planner.SaveScenario", $name)
        foreach ($e in $mine) { Set-Edit $e -Back }
    }
    Make-Edits
    [void] $script:xl.Run("Planner.CompareScenarios")
}

# ============================================================ sheet pictures

function Save-Bitmap($bmp, $file) {
    $bmp.Save($file, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

# "{Some text}" -> the row of the first cell in A:B holding exactly that text; "+n" after it adds.
function Resolve-Rows($ws, [string] $spec) {
    return [regex]::Replace($spec, '\{([^}]+)\}(\+(\d+))?', {
        param($m)
        $text = $m.Groups[1].Value
        $add = if ($m.Groups[3].Success) { [int] $m.Groups[3].Value } else { 0 }
        $last = [math]::Max($ws.Cells($ws.Rows.Count, 1).End(-4162).Row, $ws.Cells($ws.Rows.Count, 2).End(-4162).Row)
        for ($r = 1; $r -le $last; $r++) {
            if ([string] $ws.Cells($r, 1).Text -eq $text -or [string] $ws.Cells($r, 2).Text -eq $text) { return [string] ($r + $add) }
        }
        throw "no row reads '$text'"
    })
}

function Cell-For($ws, [string] $spec, $insertsRow) {
    if ($spec -match '^row(\d+):(.+)$') { $c = Col-Of $Matches[2] $ws; if ($c) { return $ws.Cells([int] $Matches[1], $c) }; return $null }
    if ($spec -match '^op(\d+):(.+)$') { $r = Row-Of $Matches[1]; $c = Col-Of $Matches[2] $ws; if ($r -and $c) { return $ws.Cells($r, $c) }; return $null }
    if ($spec -match '^head:(.+)$') { $c = Col-Of $Matches[1] $ws; if ($c) { return $ws.Cells(1, $c) }; return $null }
    if ($spec -match '^inserts:([A-Z]+)$') { if ($insertsRow) { return $ws.Range("$($Matches[1])$($insertsRow + 1)") }; return $null }
    try { return $ws.Range((Resolve-Rows $ws $spec)) } catch { return $null }
}

function Inserts-Row($ws) {
    $last = $ws.Cells($ws.Rows.Count, 2).End(-4162).Row
    for ($r = 2; $r -le $last + 5; $r++) { if ([string] $ws.Cells($r, 1).Value2 -eq "Inserts") { return $r } }
    return 0
}

function Shoot-Range($shot) {
    $ws = $null
    try { $ws = $script:wb.Worksheets.Item($shot.Sheet) } catch { }
    if (-not $ws) { "  skip  {0}: no sheet '{1}' in the sample" -f $shot.Name, $shot.Sheet; return }
    $ws.Activate()
    $insertsRow = if ($ws.Name -eq "Tools") { Inserts-Row $ws } else { 0 }
    try {
        if ($shot.Range -eq "inserts") {
            if (-not $insertsRow) { throw "no inserts table" }
            $end = $insertsRow
            while ([string] $ws.Cells($end + 1, 2).Value2 -ne "") { $end++ }
            # Out to the table's last heading (Parts per edge, Usual edge time ...).
            $lastHead = $ws.Cells($insertsRow, $ws.Columns.Count).End(-4159).Column
            $rangeText = $ws.Range($ws.Cells($insertsRow, 1), $ws.Cells($end, [math]::Max(9, $lastHead))).Address($false, $false)
        } elseif ($shot.Range -match '^used:(\d+)$') {
            # The first rows of what the sheet holds.
            $u = $ws.UsedRange
            $rangeText = $ws.Range($ws.Cells(1, 1), $ws.Cells([math]::Min([int] $Matches[1], $u.Row + $u.Rows.Count - 1), $u.Column + $u.Columns.Count - 1)).Address($false, $false)
        } else {
            $rangeText = Resolve-Rows $ws $shot.Range
        }
    } catch { "  skip  {0}: {1}" -f $shot.Name, $_.Exception.Message; return }
    $hidden = @()
    if ($shot.Show) {
        # Only the named columns: the others hidden for the picture (put back after).
        $want = @{}
        foreach ($n in $shot.Show) { $c = Col-Of $n $ws; if ($c) { $want[$c] = $true } }
        $headRow = if ($ws.Name -eq "Lathe params") { 2 } else { 1 }
        $last = $ws.Cells($headRow, $ws.Columns.Count).End(-4159).Column
        for ($c = 1; $c -le $last; $c++) {
            if (-not $want.ContainsKey($c) -and -not $ws.Columns($c).Hidden) { $ws.Columns($c).Hidden = $true; $hidden += $c }
            elseif ($want.ContainsKey($c) -and $ws.Columns($c).Hidden) { $ws.Columns($c).Hidden = $false }
        }
        $lastWanted = ($want.Keys | Measure-Object -Maximum).Maximum
        $rows = $ws.Range($rangeText).Rows.Count
        $rangeText = $ws.Range($ws.Cells(1, 1), $ws.Cells($rows, $lastWanted)).Address($false, $false)
    }
    if ($shot.Folded) { try { [void] $ws.Outline.ShowLevels(0, 1) } catch { } }
    $r = $ws.Range($rangeText)
    # Widen: a few characters more in each column (a sheet the macros autofit tightly).
    if ($shot.Widen) { foreach ($col in $r.Columns) { $col.ColumnWidth = $col.ColumnWidth + $shot.Widen } }
    $file = Join-Path $Out ($shot.Name + ".png")
    $ok = $false; $why = ""
    for ($try = 1; $try -le 6 -and -not $ok; $try++) {
        $co = $null
        try {
            [void] $r.CopyPicture(1, -4147)                            # as on screen, as a (vector) picture
            $w = $r.Width * $Scale; $h = $r.Height * $Scale
            $co = $ws.ChartObjects().Add(0, 0, $w, $h)
            $co.Chart.ChartArea.Format.Line.Visible = 0
            [void] $co.Select()
            $co.Chart.Paste()
            if ($co.Chart.Shapes.Count -lt 1) { throw "nothing pasted" }
            $pic = $co.Chart.Shapes.Item(1)
            $pic.LockAspectRatio = 0
            $pic.Left = 0; $pic.Top = 0; $pic.Width = $co.Chart.ChartArea.Width; $pic.Height = $co.Chart.ChartArea.Height
            [void] $co.Chart.Export($file, "PNG")
            $co.Delete()
            $ok = $true
        } catch {
            $why = $_.Exception.Message
            try { if ($co) { $co.Delete() } } catch { }
            Start-Sleep -Milliseconds (300 * $try)
        }
    }
    if (-not $ok) {
        foreach ($c in $hidden) { $ws.Columns($c).Hidden = $false }
        "  FAIL  {0}: CopyPicture / Export did not work - {1}" -f $shot.Name, $why; $script:failed++; return
    }

    # The callouts, while the columns are still as in the picture.
    if ($shot.Callouts) {
        $loaded = New-Object Drawing.Bitmap $file
        $bmp = New-Object Drawing.Bitmap $loaded
        $loaded.Dispose()
        $k = $bmp.Width / $r.Width
        $d = [math]::Round(15 * $k)
        foreach ($c in $shot.Callouts) {
            $cell = Cell-For $ws $c[1] $insertsRow
            if (-not $cell) { "        callout {0}: '{1}' not found - left out" -f $c[0], $c[1]; continue }
            if ($ws.Columns($cell.Column).Hidden -or $cell.Row -lt $r.Row -or $cell.Row -gt $r.Row + $r.Rows.Count - 1) { "        callout {0}: '{1}' is not in the picture - left out" -f $c[0], $c[1]; continue }
            $x0 = ($cell.Left - $r.Left) * $k; $y0 = ($cell.Top - $r.Top) * $k
            $cw = $cell.Width * $k; $ch = $cell.Height * $k
            # A "b" after l / r: at the bottom of a tall cell (where its value sits).
            $y = if ($c[2] -like "?b") { $y0 + $ch - [math]::Min($ch, 20 * $k) / 2 } else { $y0 + $ch / 2 }
            switch -wildcard ($c[2]) {
                "l*" { $x = $x0 + $d / 2 + 2 }
                "r*" { $x = $x0 + $cw - $d / 2 - 2 }
                default { $x = $x0 + $cw / 2 }
            }
            [PtWin]::Callout($bmp, [int] $c[0], [float] $x, [float] $y, [float] $d)
        }
        Save-Bitmap $bmp $file
    }
    foreach ($c in $hidden) { $ws.Columns($c).Hidden = $false }
    if ($shot.Folded) { try { [void] $ws.Outline.ShowLevels(0, 2) } catch { } }
    "  ok    {0}.png  ({1}!{2})" -f $shot.Name, $ws.Name, $rangeText
}

# ============================================================ window pictures

# "ribbon:<id>" -> "Module.Macro Nothing": the ribbon button's own macro (ribbon.xml onAction).
function Ribbon-Macro([string] $id) {
    $x = [xml] (Get-Content -LiteralPath (Join-Path $Root "vba\ribbon.xml") -Raw)
    $node = $x.SelectSingleNode("//*[@id='$id']")
    if (-not $node -or -not $node.onAction) { return $null }
    $proc = $node.onAction
    foreach ($bas in Get-ChildItem (Join-Path $Root "vba\*.bas")) {
        $text = Get-Content -LiteralPath $bas.FullName -Raw
        if ($text -match ("Public Sub " + [regex]::Escape($proc) + "\(control As IRibbonControl\)")) {
            $mod = if ($text -match 'Attribute VB_Name = "([^"]+)"') { $Matches[1] } else { $bas.BaseName }
            return "Application.Run `"$mod.$proc`", Nothing"
        }
    }
    return $null
}

# The module a window picture runs from - added to OUR copy only, one picture at a time.
function Put-Module([string] $body) {
    $vbp = $script:wb.VBProject
    try { $old = $vbp.VBComponents.Item("PtShot"); $vbp.VBComponents.Remove($old) } catch { }
    $m = $vbp.VBComponents.Add(1)
    $m.Name = "PtShot"
    $m.CodeModule.AddFromString($body)
}

$moduleTemplate = @'
Option Explicit
' Added by tools\make_help_images.ps1 to a COPY of a sample workbook - never shipped.
Public Sub PtGo()
    Dim f As Object
    On Error GoTo Bad
    ThisWorkbook.Names.Add Name:="PtShotErr", RefersTo:="=""""", Visible:=False
    %CODE%
    If Not f Is Nothing Then
        f.Show vbModeless
        PtPlaces f
    End If
    Exit Sub
Bad:
    ThisWorkbook.Names.Add Name:="PtShotErr", RefersTo:="=""" & Replace(Err.Description, """", "'") & """", Visible:=False
End Sub

' Every window closed the way its own code would (Unload), not by a message to the window:
' a hidden Excel can take a closed resizable window for the last one and shut itself.
Public Sub PtShut()
    Dim i As Long
    On Error Resume Next
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next
End Sub

' A window that is open now, by its form's name ("TextEditor"), or Nothing.
Private Function PtForm(ByVal n As String) As Object
    Dim u As Object
    For Each u In VBA.UserForms
        If TypeName(u) = n Then Set PtForm = u
    Next
End Function

' Every control's place in the window, in points ("@inside width;name:left,top,width,height;...").
Private Sub PtPlaces(ByVal f As Object)
    Dim c As Object, s As String
    On Error Resume Next
    s = "@" & f.InsideWidth & ";"
    For Each c In f.Controls
        ' Only controls placed on the window itself (one on a tab page is placed on the page).
        If c.Visible And c.Parent Is f Then s = s & c.Name & ":" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height & ";"
    Next
    ' On a hidden sheet of our own, not in a name: a name's text stops at 255 characters.
    ThisWorkbook.Worksheets("PtShotData").Range("A1").Value = s
End Sub
'@

$runTemplate = @'
Option Explicit
' Added by tools\make_help_images.ps1 to a COPY of a sample workbook - never shipped.
Public Sub PtGo()
    On Error GoTo Bad
    ThisWorkbook.Names.Add Name:="PtShotDone", RefersTo:="=0", Visible:=False
    %CODE%
    ThisWorkbook.Names.Add Name:="PtShotDone", RefersTo:="=1", Visible:=False
    Exit Sub
Bad:
    ThisWorkbook.Names.Add Name:="PtShotErr", RefersTo:="=""" & Replace(Err.Description, """", "'") & """", Visible:=False
    ThisWorkbook.Names.Add Name:="PtShotDone", RefersTo:="=2", Visible:=False
End Sub
'@

# Run a ribbon button's macro that makes a sheet (no window to wait in): from Excel's own
# loop, watched. Sets $script:lateOk when it finished; a message box instead is closed
# (our Excel restarted) and reported. (Says things as it goes - it returns nothing.)
function Run-Late([string] $id, [string] $name) {
    $script:lateOk = $false
    $code = Ribbon-Macro $id
    if (-not $code) { "  skip  {0}: the ribbon has no button '{1}' with a macro" -f $name, $id; return }
    Put-Module ($runTemplate.Replace("%CODE%", $code))
    try { $script:wb.Names.Item("PtShotDone").Delete() } catch { }
    $script:xl.OnTime([DateTime]::Now.AddSeconds(1), "'" + $script:wb.Name + "'!PtShot.PtGo")
    $deadline = [DateTime]::Now.AddSeconds(90)
    Start-Sleep -Milliseconds 1500
    while ([DateTime]::Now -lt $deadline) {
        $dlg = [PtWin]::Find([uint32] $script:xlPid, "#32770")
        if ($dlg.Count -gt 0) {
            "  skip  {0}: a message box came up instead ('{1}')" -f $name, [PtWin]::TextOf($dlg[0])
            Stop-Excel -Kill; Start-Excel
            return
        }
        foreach ($h in [PtWin]::Find([uint32] $script:xlPid, "ThunderDFrame")) { [PtWin]::Close($h) }
        $done = Name-Value "PtShotDone"
        if ($done -eq "1") {
            # Finished - unless a box is still on its way up.
            Start-Sleep -Milliseconds 500
            if ([PtWin]::Find([uint32] $script:xlPid, "#32770").Count -eq 0) { $script:lateOk = $true; return }
            continue
        }
        if ($done -eq "2") { "  skip  {0}: {1}" -f $name, (Name-Value "PtShotErr"); return }
        if (-not (Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue)) { "  skip  {0}: our Excel stopped (see the watchdog line)" -f $name; return }
        Start-Sleep -Milliseconds 300
    }
    "  skip  {0}: the macro did not finish in 90 seconds" -f $name
    Stop-Excel -Kill; Start-Excel
}

function Name-Value([string] $n) {
    try { return [string] $script:xl.Evaluate($script:wb.Names.Item($n).RefersTo) } catch { return "" }
}

function Shoot-Window($shot) {
    if ($shot.Edits) { Make-Edits }
    elseif ($script:edited) {
        # This window shows the sheet as dumped: the example edits taken back first.
        foreach ($e in $Edits) { if ($e.ContainsKey("Old")) { Set-Edit $e -Back } }
        $script:edited = $false
    }
    $ws = Main-Sheet
    $ws.Activate()
    $cells = @()
    if ($shot.Select) {
        foreach ($op in $shot.Select.Ops) {
            $r = Row-Of $op
            if (-not $r) { continue }
            foreach ($n in $shot.Select.Cols) {
                $c = Col-Of $n
                if ($c) { $cells += $ws.Cells($r, $c).Address($false, $false) }
            }
        }
    }
    if ($cells.Count -eq 0) { $cells = @("A3") }
    [void] $ws.Range($cells -join ",").Select()

    $how = $shot.Window
    if ($how -like "ribbon:*") {
        $code = Ribbon-Macro $how.Substring(7)
        if (-not $code) { "  skip  {0}: the ribbon has no button '{1}' with a macro" -f $shot.Name, $how.Substring(7); return }
    } else {
        $code = $how.Substring(4)
    }
    try { Put-Module ($moduleTemplate.Replace("%CODE%", $code)) }
    catch { "  skip  {0}: cannot add a module (is 'Trust access to the VBA project object model' on?) - {1}" -f $shot.Name, $_.Exception.Message; $script:failed++; return }
    try { Set-Value $script:wb.Worksheets.Item("PtShotData").Range("A1") "" } catch { }

    # Run it from Excel's own loop, a moment later - a window that waits (modal) never blocks us.
    $script:xl.OnTime([DateTime]::Now.AddSeconds(1), "'" + $script:wb.Name + "'!PtShot.PtGo")
    $deadline = [DateTime]::Now.AddSeconds(25)
    $form = [IntPtr]::Zero; $box = [IntPtr]::Zero
    while ([DateTime]::Now -lt $deadline) {
        Start-Sleep -Milliseconds 150
        $forms = [PtWin]::Find([uint32] $script:xlPid, "ThunderDFrame")
        if ($forms.Count -gt 0) { $form = $forms[0]; break }
        $dlg = [PtWin]::Find([uint32] $script:xlPid, "#32770")
        if ($dlg.Count -gt 0) { $box = $dlg[0]; break }
    }
    if ($form -ne [IntPtr]::Zero) {
        Start-Sleep -Milliseconds 800                                  # let it finish drawing
        $bmp = [PtWin]::Capture($form)
        $places = ""
        if ($shot.Callouts -and $how -like "vba:*") {
            for ($i = 0; $i -lt 30 -and -not $places; $i++) {
                Start-Sleep -Milliseconds 100
                try { $places = [string] $script:wb.Worksheets.Item("PtShotData").Range("A1").Value2 } catch { }
            }
        }
        if ($places -match '^@([\d.,]+);') {
            $inside = [double] ($Matches[1] -replace ',', '.')
            $cp = [PtWin]::ClientPlace($form)
            $k = $cp[2] / $inside
            $d = [math]::Round(17 * $k)
            $map = @{}
            foreach ($p in $places.Split(';')) { if ($p -match '^(\w+):(.+)$') { $map[$Matches[1]] = @($Matches[2].Split(',') | ForEach-Object { [double] ($_ -replace ',', '.') }) } }
            foreach ($c in $shot.Callouts) {
                $v = $map[$c[1]]
                if (-not $v) { "        callout {0}: no control '{1}' - left out" -f $c[0], $c[1]; continue }
                $y = $cp[1] + ($v[1] + [math]::Min($v[3], 24) / 2) * $k
                if ($c[2] -eq "r") { $x = $cp[0] + ($v[0] + $v[2]) * $k - $d / 2 - 2 } else { $x = $cp[0] + $v[0] * $k - $d / 2 - 2; if ($x -lt $d / 2) { $x = $d / 2 } }
                [PtWin]::Callout($bmp, [int] $c[0], [float] $x, [float] $y, [float] $d)
            }
        }
        Save-Bitmap $bmp (Join-Path $Out ($shot.Name + ".png"))
        "  ok    {0}.png  ('{1}'){2}" -f $shot.Name, [PtWin]::TextOf($form), $(if ($shot.Pending) { "  - a window being rebuilt: check the picture" } else { "" })
        # Closed by its own code (Unload, from Excel's loop) - a modal window by a click on its
        # close box instead, as it is still waiting.
        try { $script:xl.OnTime([DateTime]::Now, "'" + $script:wb.Name + "'!PtShot.PtShut") } catch { }
        for ($i = 0; $i -lt 30 -and [PtWin]::IsWindow($form); $i++) { Start-Sleep -Milliseconds 100 }
        if ([PtWin]::IsWindow($form)) { [PtWin]::Close($form) }
        for ($i = 0; $i -lt 40 -and [PtWin]::IsWindow($form); $i++) { Start-Sleep -Milliseconds 100 }
        if ([PtWin]::IsWindow($form)) { "        the window did not close - restarting our Excel"; Stop-Excel -Kill; Start-Excel }
    } elseif ($box -ne [IntPtr]::Zero) {
        # A message box (or a VBA error) instead of a window: keep a picture of it to read, then
        # stop our Excel (a VBA error box would otherwise wait for ever).
        Save-Bitmap ([PtWin]::Capture($box)) (Join-Path $tmp ($shot.Name + "-message.png"))
        "  skip  {0}: a message box came up instead ('{1}') - picture kept in the temp folder (use -KeepTemp)" -f $shot.Name, [PtWin]::TextOf($box)
        $script:failed++
        Stop-Excel -Kill; Start-Excel
    } else {
        $err = Name-Value "PtShotErr"
        "  skip  {0}: no window came up{1}" -f $shot.Name, $(if ($err) { " - $err" } else { "" })
        $script:failed++
    }
}

# ============================================================ run

$script:failed = 0
try {
    Start-Excel
    "excel  : our own, hidden (process $script:xlPid)"
    $sheets = @($Shots | Where-Object { -not $_.Window })
    # Pictures of the sheet as dumped first; the scenarios (which make the edits); the rest
    # with the edits; last the sheets the ribbon's commands make.
    $order = @($sheets | Where-Object { -not $_.Edits -and -not $_.Macro }) +
             @($sheets | Where-Object { $_.Macro -eq "scenarios" }) +
             @($sheets | Where-Object { $_.Edits -and $_.Macro -notlike "ribbon:*" }) +
             @($sheets | Where-Object { $_.Macro -like "ribbon:*" })
    foreach ($shot in $order) {
        Ensure-Excel
        try {
            if ($shot.Macro -eq "scenarios") { Make-Scenarios }
            elseif ($shot.Edits) { Make-Edits }
            if ($shot.Macro -like "ribbon:*") {
                Run-Late $shot.Macro.Substring(7) $shot.Name
                if (-not $script:lateOk) { $script:failed++; continue }
            } elseif ($shot.Macro -and $shot.Macro -ne "scenarios") { [void] $script:xl.Run($shot.Macro) }
            Shoot-Range $shot
            # A command that filtered the main sheet: every row back for the next picture.
            $ms = Main-Sheet
            if ($ms.FilterMode) { $ms.ShowAllData() }
        } catch { "  skip  {0}: {1}`n{2}" -f $shot.Name, $_.Exception.Message, $_.ScriptStackTrace; $script:failed++ }
    }
    foreach ($shot in @($Shots | Where-Object { $_.Window })) {
        Ensure-Excel
        try { Shoot-Window $shot } catch { "  skip  {0}: {1} ({2})" -f $shot.Name, $_.Exception.Message, ($_.ScriptStackTrace -split "`n")[0]; $script:failed++ }
    }
} finally {
    Stop-Excel
    if ($null -ne $clipText) { try { Set-Clipboard -Value $clipText } catch { } }
    if (-not $KeepTemp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue } else { "temp   : $tmp" }
}
""
"pictures in {0}" -f $Out
if ($script:failed) { "{0} picture(s) not made - see above" -f $script:failed }
"Next: python tools\make_help.py  (puts the pictures into the pages)"
