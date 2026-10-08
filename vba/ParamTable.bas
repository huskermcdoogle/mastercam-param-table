Attribute VB_Name = "ParamTable"
' Parameter Table Tool - the workbook's commands: the "Parameter Table" ribbon tab,
' and what the sheet does when a cell changes (two-way links, cross-checks).
'
' Plain text in the repo (vba\ParamTable.bas); tools\build_vba.ps1 compiles it into
' res\vbaProject.bin, which the add-in writes into a dump when macros are switched on.
'
' NOTHING HERE IS NEEDED TO LOAD A SHEET BACK. The load reads values; these only make
' editing quicker. Every write goes through the cell's own rule (the limits the load
' enforces), so a command cannot put in a value a person could not type.
Option Explicit

Private Const MAIN_SHEET As String = "Lathe params"
Private Const DUMPED_SHEET As String = "Dumped"
Private Const HEADER_ROW As Long = 2
Private Const FIRST_ROW As Long = 3
Private Const TITLE As String = "Parameter Table"

' ============================================================ lookups

Public Function MainSheet() As Worksheet
    Set MainSheet = ThisWorkbook.Worksheets(MAIN_SHEET)
End Function

' A column by its name in the column-name row, or 0.
Public Function ColOf(ByVal name As String) As Long
    Dim f As Range
    Set f = MainSheet.Rows(HEADER_ROW).Find(What:=name, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If f Is Nothing Then ColOf = 0 Else ColOf = f.Column
End Function

Public Function LastRow() As Long
    LastRow = MainSheet.Cells(MainSheet.Rows.Count, 1).End(xlUp).Row
End Function

Public Function LastCol() As Long
    LastCol = MainSheet.Cells(HEADER_ROW, MainSheet.Columns.Count).End(xlToLeft).Column
End Function

' A cell of one row by column name, or Empty.
Private Function ValueAt(ByVal r As Long, ByVal name As String) As Variant
    Dim c As Long
    c = ColOf(name)
    If c > 0 Then ValueAt = MainSheet.Cells(r, c).Value
End Function

' The Dumped row holding an operation's values as dumped, or 0.
Public Function DumpedRow(ByVal opId As Variant) As Long
    Dim ws As Worksheet, f As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Or IsEmpty(opId) Then Exit Function
    Set f = ws.Columns(1).Find(What:=opId, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        If f.Row >= FIRST_ROW Then DumpedRow = f.Row
    End If
End Function

' Calculated columns - their value follows edits elsewhere and is not an edit itself.
Private Function Untracked(ByVal header As String) As Boolean
    Select Case header
        Case "changes", "est_seconds", "est_cycle_time", "time_change": Untracked = True
    End Select
End Function

Private Function OnMain() As Boolean
    If TypeName(Selection) <> "Range" Then Exit Function
    If Selection.Worksheet.Name <> MAIN_SHEET Then
        MsgBox "Select cells on the '" & MAIN_SHEET & "' sheet first.", vbInformation, TITLE
        Exit Function
    End If
    OnMain = True
End Function

' The data cells of a range: rows from FIRST_ROW, visible ones only (a filter is respected).
Private Function DataCells(ByVal r As Range) As Range
    Dim body As Range
    If r Is Nothing Then Exit Function
    If r.Worksheet.Name <> MAIN_SHEET Then Exit Function
    Set body = Intersect(r, MainSheet.Range(MainSheet.Cells(FIRST_ROW, 1), MainSheet.Cells(LastRow, LastCol)))
    If body Is Nothing Then Exit Function
    On Error Resume Next
    Set DataCells = body.SpecialCells(xlCellTypeVisible)
End Function

' Two cell values the same, as the load would see it (numbers within a hair).
Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) And IsEmpty(b) Then SameValue = True: Exit Function
    If IsNumeric(a) And IsNumeric(b) And Not IsEmpty(a) And Not IsEmpty(b) And VarType(a) <> vbString Then
        SameValue = Abs(CDbl(a) - CDbl(b)) <= 0.000000001 * Application.Max(1, Abs(CDbl(a)))
    Else
        SameValue = (LCase$(Trim$(CStr(a))) = LCase$(Trim$(CStr(b))))
    End If
End Function

' ============================================================ writing

' Whether a cell's value passes its own rule. A cell with no rule takes anything.
Private Function CellOk(ByVal c As Range) As Boolean
    CellOk = True
    On Error Resume Next
    CellOk = c.Validation.Value
End Function

' Write a value the way typing it would be judged; undone if the cell's rule refuses it.
Public Function TryWrite(ByVal c As Range, ByVal v As Variant) As Boolean
    Dim had As Boolean, f As String, old As Variant
    had = c.HasFormula
    If had Then f = c.Formula Else old = c.Value
    If VarType(v) = vbString Then
        If IsNumeric(v) And Trim$(v) <> "" Then v = CDbl(v)
    End If
    c.Value = v
    If CellOk(c) Then
        TryWrite = True
    ElseIf had Then
        c.Formula = f
    Else
        c.Value = old
    End If
End Function

Private Sub Report(ByVal what As String, ByVal done As Long, ByVal refused As Long)
    Dim s As String
    s = done & " cell(s) " & what & "."
    If refused > 0 Then s = s & vbCrLf & refused & " refused - read-only, not for that kind of operation, or outside its limits."
    MsgBox s, vbInformation, TITLE
End Sub

' ============================================================ bulk edits

' Set every selected cell to one value.
Public Sub SetSelected()
    Dim cells As Range, v As Variant
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    v = Application.InputBox("Value for the " & cells.Count & " selected cell(s):", TITLE & " - set selected")
    If VarType(v) = vbBoolean Then Exit Sub
    ReportPair "set", SetCells(cells, v)
End Sub

' "done refused" - so a test can call it without a message box in the way.
Public Function SetCells(ByVal cells As Range, ByVal v As Variant) As String
    Dim c As Range, ok As Long, bad As Long
    For Each c In cells.Cells
        If TryWrite(c, v) Then ok = ok + 1 Else bad = bad + 1
    Next
    SetCells = ok & " " & bad
End Function

Private Sub ReportPair(ByVal what As String, ByVal pair As String)
    Dim p() As String
    p = Split(pair, " ")
    Report what, CLng(p(0)), CLng(p(1))
End Sub

' Scale every selected number by a percentage (110 = 10% more).
Public Sub ScaleSelected()
    Dim cells As Range, p As Variant
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    p = Application.InputBox("Scale the selected numbers to what percent?" & vbCrLf & _
                             "(110 = 10% more, 90 = 10% less)", TITLE & " - scale", 100, Type:=1)
    If VarType(p) = vbBoolean Then Exit Sub
    ReportPair "scaled", ScaleCells(cells, CDbl(p))
End Sub

Public Function ScaleCells(ByVal cells As Range, ByVal percent As Double) As String
    Dim c As Range, ok As Long, bad As Long, v As Variant
    For Each c In cells.Cells
        v = c.Value
        If Not c.HasFormula And Not IsEmpty(v) And IsNumeric(v) And VarType(v) <> vbString Then
            If TryWrite(c, Round(CDbl(v) * percent / 100, 10)) Then ok = ok + 1 Else bad = bad + 1
        End If
    Next
    ScaleCells = ok & " " & bad
End Function

' Copy one operation's values into the selected rows, for the selected columns.
Public Sub CopyFromOp()
    Dim cells As Range, op As Variant
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    op = Application.InputBox("Copy the selected columns FROM which operation (op_idn)?", TITLE & " - copy", Type:=1)
    If VarType(op) = vbBoolean Then Exit Sub
    If MainSheet.Columns(1).Find(What:=op, LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        MsgBox "No operation " & op & " on this sheet.", vbExclamation, TITLE
        Exit Sub
    End If
    ReportPair "copied from op " & op, CopyCells(cells, op)
End Sub

Public Function CopyCells(ByVal cells As Range, ByVal op As Variant) As String
    Dim f As Range, c As Range, ok As Long, bad As Long, src As Range
    Set f = MainSheet.Columns(1).Find(What:=op, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then CopyCells = "0 0": Exit Function
    For Each c In cells.Cells
        If c.Row <> f.Row Then
            Set src = MainSheet.Cells(f.Row, c.Column)
            If Not src.HasFormula Then
                If TryWrite(c, src.Value) Then ok = ok + 1 Else bad = bad + 1
            End If
        End If
    Next
    CopyCells = ok & " " & bad
End Function

' ============================================================ revert

Public Sub RevertSelected()
    If Not OnMain() Then Exit Sub
    RevertCells DataCells(Selection), True
End Sub

Public Sub RevertRows()
    If Not OnMain() Then Exit Sub
    RevertCells DataCells(Intersect(Selection.EntireRow, MainSheet.UsedRange)), True
End Sub

Public Sub RevertAll()
    If MsgBox("Put EVERY cell back to its dumped value?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub
    RevertCells DataCells(MainSheet.UsedRange), True
End Sub

Public Function RevertCells(ByVal cells As Range, ByVal tell As Boolean) As Long
    Dim ws As Worksheet, c As Range, dr As Long, n As Long
    If cells Is Nothing Then Exit Function
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    Application.EnableEvents = False
    For Each c In cells.Cells
        If Not Untracked(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) And Not c.HasFormula Then
            dr = DumpedRow(MainSheet.Cells(c.Row, 1).Value)
            If dr > 0 Then
                If Not SameValue(c.Value, ws.Cells(dr, c.Column).Value) Then
                    c.Value = ws.Cells(dr, c.Column).Value
                    n = n + 1
                End If
            End If
        End If
    Next
    Application.EnableEvents = True
    RevertCells = n
    If tell Then MsgBox n & " cell(s) put back to their dumped value.", vbInformation, TITLE
End Function

' ============================================================ review

' A "Changes" sheet: every cell that differs from the dump, old and new.
Public Function ListChanges() As Long
    Dim ws As Worksheet, out As Worksheet, r As Long, c As Long, dr As Long, n As Long
    Dim hdr As String, lc As Long, cmt As Long
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Changes").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set out = ThisWorkbook.Worksheets.Add(After:=MainSheet)
    out.Name = "Changes"
    out.Range("A1:F1").Value = Array("op_idn", "type", "comment", "parameter", "dumped", "now")
    out.Range("A1:F1").Font.Bold = True
    lc = LastCol
    cmt = ColOf("comment")
    For r = FIRST_ROW To LastRow
        dr = DumpedRow(MainSheet.Cells(r, 1).Value)
        If dr > 0 Then
            For c = 1 To lc
                hdr = CStr(MainSheet.Cells(HEADER_ROW, c).Value)
                If hdr <> "" And Not Untracked(hdr) Then
                    If Not SameValue(MainSheet.Cells(r, c).Value, ws.Cells(dr, c).Value) Then
                        n = n + 1
                        out.Cells(n + 1, 1).Value = MainSheet.Cells(r, 1).Value
                        out.Cells(n + 1, 2).Value = MainSheet.Cells(r, 2).Value
                        If cmt > 0 Then out.Cells(n + 1, 3).Value = MainSheet.Cells(r, cmt).Value
                        out.Cells(n + 1, 4).Value = hdr
                        out.Cells(n + 1, 5).Value = ws.Cells(dr, c).Value
                        out.Cells(n + 1, 6).Value = MainSheet.Cells(r, c).Value
                    End If
                End If
            Next
        End If
    Next
    If n = 0 Then out.Cells(2, 1).Value = "No changes from the dump."
    out.Columns("A:F").AutoFit
    ListChanges = n
End Function

Public Sub ShowChanges()
    Dim n As Long
    n = ListChanges()
    ThisWorkbook.Worksheets("Changes").Activate
End Sub

' ============================================================ calculators

' SFM <-> RPM at a diameter.
Public Sub CalcSpeed()
    Dim s As Variant, p() As String, d As Double, v As Double
    s = Application.InputBox("Enter a DIAMETER and a SURFACE SPEED to get RPM, e.g.  14 200" & vbCrLf & _
                             "or a diameter and RPM followed by 'rpm' to get SFM, e.g.  14 55 rpm" & vbCrLf & vbCrLf & _
                             "(inches and SFM; in a metric part: mm and m/min)", TITLE & " - speed")
    If VarType(s) = vbBoolean Then Exit Sub
    p = Split(Application.WorksheetFunction.Trim(CStr(s)), " ")
    If UBound(p) < 1 Then Exit Sub
    d = Val(p(0)): v = Val(p(1))
    If d <= 0 Then Exit Sub
    If UBound(p) >= 2 Then
        MsgBox Format(v, "0") & " RPM at " & d & " dia  =  " & Format(SurfaceFromRpm(d, v, False), "0.0") & " SFM" & _
               "  (" & Format(SurfaceFromRpm(d, v, True), "0.0") & " m/min if mm)", vbInformation, TITLE
    Else
        MsgBox Format(v, "0") & " SFM at " & d & " dia  =  " & Format(RpmFromSurface(d, v, False), "0") & " RPM" & _
               "  (" & Format(RpmFromSurface(d, v, True), "0") & " RPM if mm and m/min)", vbInformation, TITLE
    End If
End Sub

Public Function RpmFromSurface(ByVal dia As Double, ByVal surface As Double, ByVal metric As Boolean) As Double
    If dia > 0 Then RpmFromSurface = IIf(metric, 1000, 12) * surface / (Application.Pi() * dia)
End Function

Public Function SurfaceFromRpm(ByVal dia As Double, ByVal rpm As Double, ByVal metric As Boolean) As Double
    SurfaceFromRpm = rpm * Application.Pi() * dia / IIf(metric, 1000, 12)
End Function

' Per rev <-> per minute at an RPM.
Public Sub CalcFeed()
    Dim s As Variant, p() As String, f As Double, rpm As Double
    s = Application.InputBox("Enter a feed PER REV and an RPM to get per minute, e.g.  0.01 550" & vbCrLf & _
                             "or a feed per minute and an RPM followed by 'min' to get per rev, e.g.  5.5 550 min", _
                             TITLE & " - feed")
    If VarType(s) = vbBoolean Then Exit Sub
    p = Split(Application.WorksheetFunction.Trim(CStr(s)), " ")
    If UBound(p) < 1 Then Exit Sub
    f = Val(p(0)): rpm = Val(p(1))
    If rpm <= 0 Then Exit Sub
    If UBound(p) >= 2 Then
        MsgBox f & " per min at " & rpm & " RPM  =  " & Format(f / rpm, "0.00000") & " per rev", vbInformation, TITLE
    Else
        MsgBox f & " per rev at " & rpm & " RPM  =  " & Format(f * rpm, "0.000") & " per min", vbInformation, TITLE
    End If
End Sub

' ============================================================ pickers

' Several coolants at one timing (a dropdown takes one).
Public Sub PickCoolant()
    Dim c As Range, hdr As String, list As String, opts() As String, i As Long, s As Variant
    Dim pick() As String, out As String, k As Long
    If Not OnMain() Then Exit Sub
    Set c = ActiveCell
    hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
    If Left$(hdr, 8) <> "coolant_" Or c.Row < FIRST_ROW Then
        MsgBox "Click a cell in coolant_before, coolant_with or coolant_after first.", vbInformation, TITLE
        Exit Sub
    End If
    On Error Resume Next
    list = c.Validation.Formula1
    On Error GoTo 0
    list = Replace(list, """", "")
    If list = "" Then MsgBox "This operation has no coolant choices.", vbInformation, TITLE: Exit Sub
    opts = Split(list, ",")
    For i = 1 To UBound(opts)             ' 0 is "none"
        s = s & i & "  " & opts(i) & vbCrLf
    Next
    s = Application.InputBox("Coolants for " & hdr & " - numbers joined by +, e.g.  1+3" & vbCrLf & _
                             "(0 = none)" & vbCrLf & vbCrLf & s, TITLE & " - coolant")
    If VarType(s) = vbBoolean Then Exit Sub
    pick = Split(Replace(CStr(s), " ", ""), "+")
    For i = 0 To UBound(pick)
        k = Val(pick(i))
        If k >= 1 And k <= UBound(opts) Then out = out & IIf(out = "", "", " + ") & opts(k)
    Next
    If out = "" Then out = "none"
    c.Value = out          ' two coolants are not a dropdown choice; the load reads "A + B"
End Sub

' ============================================================ as cells change

' Called by the sheet whenever cells change: keep linked pairs in step, and say so in
' the status bar when a row looks wrong. Nothing is refused here - that is the load's job.
Public Sub OnCellsChanged(ByVal Target As Range)
    Dim c As Range, hdr As String, msgs As String, seen As String
    If Target.Cells.Count > 5000 Then Exit Sub
    For Each c In Target.Cells
        If c.Row >= FIRST_ROW Then
            hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
            Select Case hdr
                Case "stepover": LinkAmount c, "stepover_percent"
                Case "radius": LinkAmount c, "radius_percent"
                Case "stepover_percent": LinkPercent c, "stepover"
                Case "radius_percent": LinkPercent c, "radius"
            End Select
            If InStr(seen, "|" & c.Row & "|") = 0 Then
                seen = seen & "|" & c.Row & "|"
                msgs = msgs & CheckRow(c.Row)
            End If
        End If
    Next
    If msgs = "" Then Application.StatusBar = False Else Application.StatusBar = TITLE & ":  " & msgs
End Sub

' An amount changed: its percent follows (a formula already does; a typed value is updated).
Private Sub LinkAmount(ByVal c As Range, ByVal percentName As String)
    Dim pc As Long, r As Variant
    pc = ColOf(percentName)
    r = ValueAt(c.Row, "tool_radius")
    If pc = 0 Or IsEmpty(r) Or Not IsNumeric(r) Then Exit Sub
    If CDbl(r) = 0 Or MainSheet.Cells(c.Row, pc).HasFormula Then Exit Sub
    If Not IsEmpty(c.Value) And IsNumeric(c.Value) Then
        MainSheet.Cells(c.Row, pc).Value = Round(CDbl(c.Value) / CDbl(r) * 100, 6)
    End If
End Sub

' A percent was typed: its amount follows.
Private Sub LinkPercent(ByVal c As Range, ByVal amountName As String)
    Dim ac As Long, r As Variant
    ac = ColOf(amountName)
    r = ValueAt(c.Row, "tool_radius")
    If ac = 0 Or IsEmpty(r) Or Not IsNumeric(r) Then Exit Sub
    If Not IsEmpty(c.Value) And IsNumeric(c.Value) And Not c.HasFormula Then
        MainSheet.Cells(c.Row, ac).Value = Round(CDbl(c.Value) * CDbl(r) / 100, 10)
    End If
End Sub

' Gentle warnings about one row - things that are legal but usually a slip.
Public Function CheckRow(ByVal r As Long) As String
    Dim op As Variant, sm As Variant, mx As Variant, fm As Variant, f As Variant, mm As Boolean, s As String
    op = MainSheet.Cells(r, 1).Value
    sm = ValueAt(r, "speed_mode"): mx = ValueAt(r, "max_ss")
    fm = ValueAt(r, "feed_mode"): f = ValueAt(r, "feed")
    mm = (CStr(ValueAt(r, "units")) = "mm")
    If CStr(sm) = "CSS" And (IsEmpty(mx) Or Val(CStr(mx)) = 0) Then s = s & "op " & op & ": CSS with no max_ss.  "
    If IsNumeric(f) And Not IsEmpty(f) Then
        If CStr(fm) = "per min" And CDbl(f) > 0 And CDbl(f) < IIf(mm, 25, 1) Then _
            s = s & "op " & op & ": feed " & f & " per min is very slow - meant per rev?  "
        If CStr(fm) = "per rev" And CDbl(f) > IIf(mm, 2.5, 0.1) Then _
            s = s & "op " & op & ": feed " & f & " per rev is very high - meant per min?  "
    End If
    CheckRow = s
End Function

' ============================================================ ribbon

Public Sub RbSet(control As IRibbonControl): SetSelected: End Sub
Public Sub RbScale(control As IRibbonControl): ScaleSelected: End Sub
Public Sub RbCopy(control As IRibbonControl): CopyFromOp: End Sub
Public Sub RbRevertSel(control As IRibbonControl): RevertSelected: End Sub
Public Sub RbRevertRows(control As IRibbonControl): RevertRows: End Sub
Public Sub RbRevertAll(control As IRibbonControl): RevertAll: End Sub
Public Sub RbChanges(control As IRibbonControl): ShowChanges: End Sub
Public Sub RbSpeed(control As IRibbonControl): CalcSpeed: End Sub
Public Sub RbFeed(control As IRibbonControl): CalcFeed: End Sub
Public Sub RbCoolant(control As IRibbonControl): PickCoolant: End Sub
