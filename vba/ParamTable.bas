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
Public Function ValueAt(ByVal r As Long, ByVal name As String) As Variant
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
Public Function Untracked(ByVal header As String) As Boolean
    Select Case header
        Case "changes", "est_seconds", "est_cycle_time", "time_change", "flips_est", "cut_seconds_est", _
             "flips_uncommented", "flips_part", "edge_limit", "edge_after", "mrr", "mrr_avg", "mrr_engaged": Untracked = True
    End Select
End Function

Public Function OnMain() As Boolean
    If TypeName(Selection) <> "Range" Then Exit Function
    If Selection.Worksheet.Name <> MAIN_SHEET Then
        MsgBox "Select cells on the '" & MAIN_SHEET & "' sheet first.", vbInformation, TITLE
        Exit Function
    End If
    OnMain = True
End Function

' The data cells of a range: rows from FIRST_ROW, visible ones only (a filter is respected).
Public Function DataCells(ByVal r As Range) As Range
    Dim body As Range
    If r Is Nothing Then Exit Function
    If r.Worksheet.Name <> MAIN_SHEET Then Exit Function
    Set body = Intersect(r, MainSheet.Range(MainSheet.Cells(FIRST_ROW, 1), MainSheet.Cells(LastRow, LastCol)))
    If body Is Nothing Then Exit Function
    On Error Resume Next
    Set DataCells = body.SpecialCells(xlCellTypeVisible)
End Function

' Every data cell, hidden rows too (what "everything" means, whatever the filter shows).
Public Function AllData() As Range
    Set AllData = MainSheet.Range(MainSheet.Cells(FIRST_ROW, 1), MainSheet.Cells(Application.Max(FIRST_ROW, LastRow), LastCol))
End Function

' Two cell values the same, as the load would see it (numbers within a hair).
Public Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
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
    Panel.Journal c                         ' for undo
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

' What a command did, as a short line in the status bar - the window already showed what
' would happen, so no box to click away. The rows' own warnings (CSS with no max_ss ...)
' stay on the line after it: a result never hides them.
Private Sub Report(ByVal what As String, ByVal done As Long, ByVal refused As Long, Optional ByVal cells As Range)
    Dim s As String, w As String
    s = done & " cell(s) " & what & "."
    If refused > 0 Then s = s & "  " & refused & " refused - read-only, not for that kind of operation, outside its limits, or in other units."
    If done > 0 Then s = s & "  Ctrl+Z to undo."
    w = RowWarnings(cells)
    Application.StatusBar = TITLE & ":  " & s & IIf(w <> "", "     " & w, "")
End Sub

' The warnings (CheckRow) of the rows some cells are on, each row once.
Private Function RowWarnings(ByVal cells As Range) As String
    Dim c As Range, seen As String
    If cells Is Nothing Then Exit Function
    For Each c In cells.Cells
        If c.Row >= FIRST_ROW And InStr(seen, "|" & c.Row & "|") = 0 Then
            seen = seen & "|" & c.Row & "|"
            RowWarnings = RowWarnings & CheckRow(c.Row)
            If Len(RowWarnings) > 400 Then Exit Function
        End If
    Next
End Function

' ============================================================ bulk edits

' Set every selected cell to one value. Coolant cells get the coolant window instead -
' several coolants at once, each row checked against its own machine's coolants.
Public Sub SetSelected()
    Dim cells As Range, names As String, f As EditBox
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    If AllCoolant(cells) Then
        names = AskCoolants(cells, cells.Count & " selected coolant cell(s)")
        If names <> "" Then ReportPair "set to " & names, SetCoolantCells(cells, names), cells
        Exit Sub
    End If
    Set f = EditWindow("set", cells)
    f.Show
    If f.Accepted Then ReportPair "set to " & f.Value(), SetCells(cells, f.Value()), cells
    Unload f
End Sub

' "done refused" - so a test can call it without a message box in the way.
Public Function SetCells(ByVal cells As Range, ByVal v As Variant) As String
    Dim c As Range, ok As Long, bad As Long
    For Each c In cells.Cells
        If TryWrite(c, v) Then ok = ok + 1 Else bad = bad + 1
    Next
    SetCells = ok & " " & bad
End Function

Private Sub ReportPair(ByVal what As String, ByVal pair As String, Optional ByVal cells As Range)
    Dim p() As String
    p = Split(pair, " ")
    Report what, CLng(p(0)), CLng(p(1)), cells
End Sub

' Make every selected number bigger or smaller: +10% = 10% more, -10% = 10% less.
Public Sub ScaleSelected()
    Dim cells As Range, f As EditBox, pc As Double
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    Set f = EditWindow("scale", cells)
    f.Show
    If f.Accepted Then
        If Planner.PercentChange(f.Value(), pc) = "" Then ReportPair "made " & Planner.PercentWords(pc), ScaleCells(cells, pc), cells
    End If
    Unload f
End Sub

' Scale numbers by a change in percent (10 = 10% more, -10 = 10% less) - "done refused".
Public Function ScaleCells(ByVal cells As Range, ByVal change As Double) As String
    Dim c As Range, ok As Long, bad As Long, v As Variant
    For Each c In cells.Cells
        v = c.Value
        If Not c.HasFormula And Not IsEmpty(v) And IsNumeric(v) And VarType(v) <> vbString Then
            If TryWrite(c, Round(CDbl(v) * (1 + change / 100), 10)) Then ok = ok + 1 Else bad = bad + 1
        End If
    Next
    ScaleCells = ok & " " & bad
End Function

' Copy one op's values, for the selected columns, into the other selected rows or into
' every other op of its tool. The op the cursor is on is the one copied from, to start.
Public Sub CopyFromOp()
    Dim cells As Range, f As EditBox, t As Range
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    Set f = EditWindow("copy", cells)
    f.Show
    If f.Accepted Then
        Set t = CopyTargets(cells, f.Value(), f.IntoIdx())
        If Not t Is Nothing Then ReportPair "copied from op " & f.Value(), CopyCells(t, f.Value()), t
    End If
    Unload f
End Sub

' The cells a copy writes: the selection (into = 0) or the selected columns on every
' other op of the source op's tool (1). Nothing when there are none.
Public Function CopyTargets(ByVal cells As Range, ByVal op As Variant, ByVal into As Long) As Range
    If into = 1 Then Set CopyTargets = Planner.ToolTargets(cells, op) Else Set CopyTargets = cells
End Function

' Copy an op's values into cells (their columns, from its row) - "done refused". A value
' whose units differ between the two rows (per rev / per min, CSS / RPM) and are not copied
' with it is refused: it would mean something else there.
Public Function CopyCells(ByVal cells As Range, ByVal op As Variant) As String
    Dim f As Range, c As Range, ok As Long, bad As Long, src As Range, copied As String
    Set f = OpCell(op)
    If f Is Nothing Then CopyCells = "0 0": Exit Function
    copied = CopiedColumns(cells)
    For Each c In cells.Cells
        If c.Row <> f.Row Then
            Set src = MainSheet.Cells(f.Row, c.Column)
            If Not src.HasFormula Then
                If UnitClash(c, f.Row, copied) <> "" Then
                    bad = bad + 1
                ElseIf TryWrite(c, src.Value) Then
                    ok = ok + 1
                Else
                    bad = bad + 1
                End If
            End If
        End If
    Next
    CopyCells = ok & " " & bad
End Function

' An operation's op_idn cell, or Nothing.
Public Function OpCell(ByVal op As Variant) As Range
    If Trim$(CStr(op)) = "" Then Exit Function
    If IsNumeric(op) Then op = CDbl(op)
    Set OpCell = MainSheet.Range(MainSheet.Cells(FIRST_ROW, 1), MainSheet.Cells(LastRow, 1)).Find( _
                     What:=op, LookIn:=xlValues, LookAt:=xlWhole)
End Function

' ---------------------------------------------------------------- units go with the value

' The column saying what units a value is in, or "".
Private Function UnitColumn(ByVal hdr As String) As String
    Select Case hdr
    Case "feed", "plunge", "retract": UnitColumn = hdr & "_mode"
    Case "speed": UnitColumn = "speed_mode"
    Case "finish_ss": UnitColumn = "finish_ss_css"
    Case "pt_rough_speed": UnitColumn = "pt_rough_css"
    Case "pt_fin_speed": UnitColumn = "pt_fin_css"
    Case "pt_rough_feed_axial": UnitColumn = "pt_rough_axial_type"
    Case "pt_rough_feed_radial": UnitColumn = "pt_rough_radial_type"
    Case "pt_fin_feed_axial": UnitColumn = "pt_fin_axial_type"
    Case "pt_fin_feed_radial": UnitColumn = "pt_fin_radial_type"
    End Select
End Function

' And back: the value a units column belongs to, or "".
Private Function ValueOfUnit(ByVal hdr As String) As String
    Dim v As Variant
    For Each v In Array("feed", "plunge", "retract", "speed", "finish_ss", "pt_rough_speed", "pt_fin_speed", _
                        "pt_rough_feed_axial", "pt_rough_feed_radial", "pt_fin_feed_axial", "pt_fin_feed_radial")
        If UnitColumn(CStr(v)) = hdr Then ValueOfUnit = v: Exit Function
    Next
End Function

' "|feed|feed_mode|" - the columns some cells are in.
Private Function CopiedColumns(ByVal cells As Range) As String
    Dim c As Range, h As String
    CopiedColumns = "|"
    If cells Is Nothing Then Exit Function
    For Each c In cells.Cells
        h = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
        If InStr(CopiedColumns, "|" & h & "|") = 0 Then CopiedColumns = CopiedColumns & h & "|"
    Next
End Function

' Why copying a cell from the source row would change what it means - its units differ
' between the two rows and the units are not copied with it (or units are copied without
' their value) - in words; "" when it is fine.
Private Function UnitClash(ByVal c As Range, ByVal srcRow As Long, ByVal copied As String) As String
    Dim hdr As String, other As String, uc As Long, a As String, b As String, isUnit As Boolean
    hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
    other = UnitColumn(hdr)
    If other = "" Then
        other = ValueOfUnit(hdr)
        If other = "" Then Exit Function
        isUnit = True
    End If
    If InStr(copied, "|" & other & "|") > 0 Then Exit Function      ' copied together: fine
    uc = IIf(isUnit, c.Column, ColOf(other))
    If uc = 0 Then Exit Function
    a = Trim$(CStr(MainSheet.Cells(srcRow, uc).Value))
    b = Trim$(CStr(MainSheet.Cells(c.Row, uc).Value))
    If a = "" Or b = "" Or LCase$(a) = LCase$(b) Then Exit Function
    If isUnit Then
        UnitClash = "op " & MainSheet.Cells(c.Row, 1).Value & "'s " & other & " is a " & b & " value - select the " & other & " too"
    Else
        UnitClash = "op " & MainSheet.Cells(srcRow, 1).Value & " is " & a & ", op " & MainSheet.Cells(c.Row, 1).Value & _
                    " is " & b & " - select the " & a & " / " & b & " column too"
    End If
End Function

' ---------------------------------------------------------------- the edit window

' "4 cells on ops 2, 7" - what is selected, in words.
Private Function Describe(ByVal cells As Range) As String
    Dim c As Range, ops As String, o As String, no As Long
    For Each c In cells.Cells
        o = CStr(MainSheet.Cells(c.Row, 1).Value)
        If InStr("|" & ops & "|", "|" & o & "|") = 0 Then
            no = no + 1
            If no <= 8 Then ops = ops & IIf(ops = "", "", "|") & o
        End If
    Next
    Describe = cells.Count & IIf(cells.Count = 1, " cell", " cells") & " on op" & IIf(no = 1, " ", "s ") & _
               Replace(ops, "|", ", ") & IIf(no > 8, ", ...", "")
End Function

' The window for one of the edits, set up on the cells but not shown. Modes: set, scale,
' copy (from an op), scensave / scenload / scendel.
Public Function EditWindow(ByVal mode As String, ByVal cells As Range) As EditBox
    Dim f As EditBox, c As Range, lst As String, sameList As Boolean, items() As String, k() As String
    Dim r As Long, n As Long, cmt As Long, v As String, seen As String, its As Variant, hint As String
    Set f = New EditBox
    If Not cells Is Nothing Then hint = "columns: " & ColumnsOf(cells)
    Select Case mode
    Case "set"
        ' Every cell has the same dropdown: offer just that list. Otherwise the values there now.
        sameList = True
        For Each c In cells.Cells
            v = ListOf(c)
            If c.Address = cells.Cells(1, 1).Address Then lst = v
            If v = "" Or v <> lst Then sameList = False
        Next
        If sameList Then
            f.Setup "set", cells, TITLE & " - set", Describe(cells), hint, "Pick the value for every selected cell:", _
                    Split(lst, ","), Empty, True, "", "set"
        Else
            ReDim items(0 To 0)
            For Each c In cells.Cells
                v = CStr(c.Value)
                If v <> "" And InStr(vbLf & seen & vbLf, vbLf & v & vbLf) = 0 And n < 12 Then
                    ReDim Preserve items(0 To n): items(n) = v: n = n + 1
                    seen = seen & vbLf & v
                End If
            Next
            If n > 0 Then
                f.Setup "set", cells, TITLE & " - set", Describe(cells), hint, _
                        "New value for every selected cell (the list has what is there now):", items, Empty, False, _
                        CStr(cells.Cells(1, 1).Value), "set"
            Else
                f.Setup "set", cells, TITLE & " - set", Describe(cells), hint, "New value for every selected cell:", _
                        Empty, Empty, False, "", "set"
            End If
        End If
    Case "scale"
        f.Setup "scale", cells, TITLE & " - scale", Describe(cells), hint, _
                "How much bigger or smaller?   +10% = 10% more,   -10% = 10% less", _
                Array("+5%", "+10%", "+20%", "-5%", "-10%", "-20%"), Empty, False, "", "scale"
    Case "copy"
        cmt = ColOf("comment")
        For r = FIRST_ROW To LastRow
            ReDim Preserve items(0 To n): ReDim Preserve k(0 To n)
            k(n) = CStr(MainSheet.Cells(r, 1).Value)
            items(n) = "op " & k(n) & "    " & MainSheet.Cells(r, 2).Value & _
                       IIf(cmt > 0, "    " & MainSheet.Cells(r, IIf(cmt > 0, cmt, 1)).Value, "")
            n = n + 1
        Next
        f.Setup "copy", cells, TITLE & " - copy from an op", "Copies " & Describe(cells) & _
                " - the same columns - from one op into others.", hint, "Copy from:", items, k, True, _
                CopySource(cells), "copy"
    Case "scensave", "scenload", "scendel"
        its = Planner.ScenarioNames()
        Select Case mode
        Case "scensave"
            f.Setup mode, Nothing, TITLE & " - save scenario", _
                    "Keeps every change now on the sheet under a name, to restore or compare later (saved in this workbook).", _
                    "", "Name (pick one to replace it):", its, Empty, False, "", "scenarios"
        Case "scenload"
            v = ""
            If UBound(its) >= 0 Then v = CStr(its(UBound(its)))      ' the latest
            f.Setup mode, Nothing, TITLE & " - restore scenario", _
                    "Puts the sheet back to the dump, then sets the scenario's values - each through its cell's rule.", _
                    "", "Restore which scenario?", its, Empty, True, v, "scenarios"
        Case Else
            f.Setup mode, Nothing, TITLE & " - delete scenario", "Deletes a saved scenario; the sheet is not changed.", _
                    "", "Delete which scenario?", its, Empty, True, "", "scenarios"
        End Select
    End Select
    Set EditWindow = f
End Function

' The op a copy starts from: the one the cursor is on, when it is an op's row - else the
' first selected one.
Private Function CopySource(ByVal cells As Range) As String
    Dim r As Long
    On Error Resume Next
    If ActiveSheet.Name = MAIN_SHEET Then r = ActiveCell.Row
    On Error GoTo 0
    If r < FIRST_ROW Or r > LastRow Then r = cells.Cells(1, 1).Row
    CopySource = CStr(MainSheet.Cells(r, 1).Value)
End Function

' Where a copy from an op can go, in words: [0] the selected rows (less the op copied
' from), [1] every other op of that op's tool.
Public Function CopyIntoChoices(ByVal cells As Range, ByVal op As Variant) As Variant
    Dim src As Range, c As Range, ops As String, o As String, n As Long, t As Range, tool As String, a As String, b As String
    Set src = OpCell(op)
    If src Is Nothing Then
        CopyIntoChoices = Array("The selected rows", "Every other op of its tool")
        Exit Function
    End If
    For Each c In cells.Cells
        If c.Row <> src.Row Then
            o = CStr(MainSheet.Cells(c.Row, 1).Value)
            If InStr("|" & ops & "|", "|" & o & "|") = 0 Then
                n = n + 1
                If n <= 8 Then ops = ops & IIf(ops = "", "", "|") & o
            End If
        End If
    Next
    If n = 0 Then
        a = "The selected rows - none besides op " & op & " (select the rows to copy into)"
    Else
        a = "The selected rows: op" & IIf(n = 1, " ", "s ") & Replace(ops, "|", ", ") & IIf(n > 8, ", ...", "")
    End If
    tool = Planner.ToolOfRow(src.Row)
    Set t = Planner.ToolTargets(cells, op)
    If tool = "" Then
        b = "Every other op of its tool - op " & op & " has no tool number"
    ElseIf t Is Nothing Then
        b = "Every other op of tool " & tool & " - no other op uses it"
    Else
        ops = "": n = 0
        For Each c In Intersect(t.EntireRow, MainSheet.Columns(1)).Cells
            n = n + 1
            If n <= 8 Then ops = ops & IIf(ops = "", "", ", ") & MainSheet.Cells(c.Row, 1).Value
        Next
        b = "Every other op of tool " & tool & ": op" & IIf(n = 1, " ", "s ") & ops & IIf(n > 8, ", ...", "")
    End If
    CopyIntoChoices = Array(a, b)
End Function

' Where a copy goes to start with: the other selected rows when there are some, else the
' rest of the tool.
Public Function CopyIntoDefault(ByVal cells As Range, ByVal op As Variant) As Long
    Dim src As Range, c As Range
    Set src = OpCell(op)
    If src Is Nothing Then Exit Function
    For Each c In cells.Cells
        If c.Row <> src.Row Then Exit Function
    Next
    CopyIntoDefault = 1
End Function

' "feed, speed" - the selected columns, in words.
Private Function ColumnsOf(ByVal cells As Range) As String
    Dim c As Range, cols As String, h As String, nc As Long
    For Each c In cells.Cells
        h = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
        If InStr("|" & cols & "|", "|" & h & "|") = 0 Then
            nc = nc + 1
            If nc <= 5 Then cols = cols & IIf(cols = "", "", "|") & h
        End If
    Next
    ColumnsOf = Replace(cols, "|", ", ") & IIf(nc > 5, " and " & nc - 5 & " more", "")
End Function

' A cell's dropdown list ("a,b,c"), or "" when it has none.
Public Function ListOf(ByVal c As Range) As String
    On Error Resume Next
    If c.Validation.Type = xlValidateList Then ListOf = Replace(c.Validation.Formula1, """", "")
End Function

' Whether a cell's rule would take a value - judged without writing it, for the preview.
' (TryWrite still has the last word when OK writes.)
Public Function WouldTake(ByVal c As Range, ByVal v As Variant) As Boolean
    Dim t As Long, op As Long, f1 As String, f2 As String, a As Double, b As Double, x As Double, s As String, it As Variant
    WouldTake = True
    On Error Resume Next
    t = -1
    t = c.Validation.Type
    If t = -1 Then Exit Function                      ' no rule
    op = c.Validation.Operator
    f1 = Replace(c.Validation.Formula1, "=", "")
    f2 = Replace(c.Validation.Formula2, "=", "")
    On Error GoTo 0
    s = CStr(v)
    Select Case t
    Case xlValidateCustom
        If UCase$(f1) = "FALSE" Then WouldTake = False   ' read-only / does not apply
        Exit Function
    Case xlValidateList
        If Trim$(s) = "" Then Exit Function
        WouldTake = False
        For Each it In Split(Replace(f1, """", ""), ",")
            If LCase$(Trim$(it)) = LCase$(Trim$(s)) Then WouldTake = True
        Next
        Exit Function
    Case xlValidateTextLength
        x = Len(s)
    Case xlValidateWholeNumber, xlValidateDecimal
        If Trim$(s) = "" Then Exit Function
        If Not IsNumeric(s) Then WouldTake = False: Exit Function
        x = CDbl(s)
        If t = xlValidateWholeNumber And x <> Int(x) Then WouldTake = False: Exit Function
    Case Else
        Exit Function
    End Select
    If Not IsNumeric(f1) Then Exit Function           ' a limit from a formula: leave it to TryWrite
    a = CDbl(f1)
    If IsNumeric(f2) Then b = CDbl(f2)
    Select Case op
        Case xlBetween: WouldTake = (x >= a And x <= b)
        Case xlNotBetween: WouldTake = (x < a Or x > b)
        Case xlEqual: WouldTake = (x = a)
        Case xlNotEqual: WouldTake = (x <> a)
        Case xlGreater: WouldTake = (x > a)
        Case xlLess: WouldTake = (x < a)
        Case xlGreaterEqual: WouldTake = (x >= a)
        Case xlLessEqual: WouldTake = (x <= a)
    End Select
End Function

' What an edit does to one cell: 0 = leaves it, 1 = changes it, 2 = already that,
' 3 = refused by its rule, 4 = not copied: its units differ (why says how). nv is the new
' value; copied the columns a copy takes ("|feed|speed|").
Private Function CellPlan(ByVal mode As String, ByVal c As Range, ByVal v As String, ByVal src As Range, ByVal copied As String, _
                          ByRef nv As Variant, ByRef why As String) As Long
    Dim pc As Double
    why = ""
    Select Case mode
    Case "set"
        If IsNumeric(v) Then nv = CDbl(v) Else nv = v
        If Not WouldTake(c, v) Then
            CellPlan = 3
        ElseIf SameValue(c.Value, nv) Then
            CellPlan = 2
        Else
            CellPlan = 1
        End If
    Case "scale"
        If c.HasFormula Or IsEmpty(c.Value) Or Not IsNumeric(c.Value) Then Exit Function
        If VarType(c.Value) = vbString Then Exit Function
        If Planner.PercentChange(v, pc) <> "" Then Exit Function
        nv = Round(CDbl(c.Value) * (1 + pc / 100), 10)
        If Not WouldTake(c, nv) Then
            CellPlan = 3
        ElseIf nv = c.Value Then
            CellPlan = 2
        Else
            CellPlan = 1
        End If
    Case "copy"
        If c.Row = src.Row Then Exit Function
        If MainSheet.Cells(src.Row, c.Column).HasFormula Then Exit Function
        nv = MainSheet.Cells(src.Row, c.Column).Value
        why = UnitClash(c, src.Row, copied)
        If why <> "" Then
            CellPlan = 4
        ElseIf Not WouldTake(c, nv) Then
            CellPlan = 3
        ElseIf SameValue(c.Value, nv) Then
            CellPlan = 2
        Else
            CellPlan = 1
        End If
    End Select
End Function

' The edit window's live line: "1|..." when OK would change something, "0|..." when not.
' into: a copy's 0 (the selected rows) or 1 (every other op of the tool).
Public Function EditPreview(ByVal mode As String, ByVal cells As Range, ByVal v As String, Optional ByVal into As Long = 0) As String
    Dim c As Range, take As Long, bad As Long, same As Long, units As Long, nv As Variant, eg As String, src As Range, shown As Long
    Dim targets As Range, ops As String, nops As Long, o As String, why As String, pc As Double, copied As String, lead As String
    Select Case mode
    Case "scensave", "scenload", "scendel"
        EditPreview = Planner.ScenarioPreview(mode, v)
        Exit Function
    Case "set"
        If Trim$(v) = "" Then EditPreview = "0|Type or pick a value.": Exit Function
    Case "scale"
        why = Planner.PercentChange(v, pc)
        If why <> "" Then EditPreview = "0|" & why: Exit Function
        lead = Planner.PercentWords(pc) & ":  "
    Case "copy"
        Set src = OpCell(v)
        If src Is Nothing Then EditPreview = "0|Pick the op to copy from.": Exit Function
    End Select
    Set targets = EditTargets(mode, cells, v, into)
    If targets Is Nothing Then EditPreview = "0|No other op uses this tool - pick the selected rows, or select some.": Exit Function
    copied = CopiedColumns(targets)
    For Each c In targets.Cells
        Select Case CellPlan(mode, c, v, src, copied, nv, why)
        Case 1
            take = take + 1
            If mode = "scale" And shown < 3 Then eg = eg & IIf(eg = "", "", ",   ") & c.Value & " -> " & nv: shown = shown + 1
            o = CStr(MainSheet.Cells(c.Row, 1).Value)
            If InStr("|" & ops & "|", "|" & o & "|") = 0 Then
                nops = nops + 1
                If nops <= 8 Then ops = ops & IIf(ops = "", "", "|") & o
            End If
        Case 2: same = same + 1
        Case 3: bad = bad + 1
        Case 4: units = units + 1
        End Select
    Next
    If eg <> "" Then eg = vbCrLf & "e.g.  " & eg
    If mode = "copy" And take > 0 Then eg = vbCrLf & "Into op" & IIf(nops = 1, " ", "s ") & Replace(ops, "|", ", ") & IIf(nops > 8, ", ...", "") & "."
    If mode = "copy" And take = 0 And same = 0 And bad = 0 And units = 0 Then
        EditPreview = "0|Nothing to copy into - select the rows to copy into, or pick every other op of the tool.": Exit Function
    End If
    EditPreview = IIf(take > 0, "1|", "0|") & lead & take & " cell(s) will change" & _
                  IIf(same > 0, ",  " & same & " already that", "") & _
                  IIf(bad > 0, ",  " & bad & " refused (read-only, not for that operation, or outside its limits)", "") & _
                  IIf(units > 0, ",  " & units & " not copied - the units differ (per rev / per min, CSS / RPM); listed below", "") & _
                  "." & eg
End Function

' The edit window's list: one line per cell it would change or refuse.
Public Function EditDetail(ByVal mode As String, ByVal cells As Range, ByVal v As String, Optional ByVal into As Long = 0) As Variant
    Dim c As Range, nv As Variant, s As String, n As Long, src As Range, targets As Range, st As Long, hdr As String
    Dim why As String, pc As Double, copied As String
    EditDetail = Array()
    Select Case mode
    Case "scensave", "scenload", "scendel"
        EditDetail = Planner.ScenarioDetail(IIf(mode = "scensave", "", v))
        Exit Function
    Case "set"
        If Trim$(v) = "" Then Exit Function
    Case "scale"
        If Planner.PercentChange(v, pc) <> "" Then Exit Function
    Case "copy"
        Set src = OpCell(v)
        If src Is Nothing Then Exit Function
    End Select
    Set targets = EditTargets(mode, cells, v, into)
    If targets Is Nothing Then Exit Function
    copied = CopiedColumns(targets)
    For Each c In targets.Cells
        st = CellPlan(mode, c, v, src, copied, nv, why)
        If st = 1 Or st = 3 Or st = 4 Then
            n = n + 1
            If n <= 300 Then
                hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
                s = s & IIf(s = "", "", vbLf) & "op " & MainSheet.Cells(c.Row, 1).Value & "    " & hdr & "    "
                Select Case st
                Case 1: s = s & IIf(CStr(c.Value) = "", "(blank)", CStr(c.Value)) & "  ->  " & CStr(nv)
                Case 3: s = s & "refused  (" & CStr(nv) & ")"
                Case 4: s = s & "not copied - " & why
                End Select
            End If
        End If
    Next
    If n > 300 Then s = s & vbLf & "... and " & n - 300 & " more"
    If s <> "" Then EditDetail = Split(s, vbLf)
End Function

' The cells an edit writes: the selection - for a copy into the tool, the rest of its ops.
Private Function EditTargets(ByVal mode As String, ByVal cells As Range, ByVal v As String, ByVal into As Long) As Range
    If mode = "copy" Then Set EditTargets = CopyTargets(cells, v, into) Else Set EditTargets = cells
End Function

' Save, restore or delete a scenario through the edit window.
Public Sub ScenarioWindow(ByVal mode As String)
    Dim f As EditBox, r As String, p() As String, n As Long
    If mode <> "scensave" And UBound(Planner.ScenarioNames()) < 0 Then
        MsgBox "No scenarios saved yet - edit the sheet, then Save scenario.", vbInformation, TITLE
        Exit Sub
    End If
    Set f = EditWindow(mode, Nothing)
    f.Show
    If f.Accepted Then
        Select Case mode
        Case "scensave"
            n = Planner.SaveScenario(f.Value())
            Application.StatusBar = TITLE & ":  " & n & " edited cell(s) saved as '" & Trim$(f.Value()) & "'."
        Case "scenload"
            p = Split(Planner.RestoreScenario(f.Value()), " ")
            Application.StatusBar = TITLE & ":  '" & f.Value() & "' restored - " & p(0) & " cell(s) set" & _
                                    IIf(p(1) <> "0", ", " & p(1) & " refused (the op or column is gone, or the value is outside its limits)", "") & "." & _
                                    IIf(Planner.AutoSaved() <> "", "  Your changes before it are kept as '" & Planner.AutoSaved() & "'.", "") & _
                                    "  Ctrl+Z to undo."
        Case "scendel"
            Planner.DeleteScenario f.Value()
            Application.StatusBar = TITLE & ":  '" & f.Value() & "' deleted."
        End Select
    End If
    Unload f
End Sub

' For tools\check_macros.ps1 (no window): what the window says and whether OK is on,
' after typing a value ("#3" picks list row 3) and, for a copy, where into (0 the selected
' rows, 1 the tool's other ops) - "preview|ok|value".
Public Function EditWindowSelfTest(ByVal mode As String, ByVal cells As Range, ByVal typed As String, Optional ByVal into As Long = -1) As String
    Dim f As EditBox
    Set f = EditWindow(mode, cells)
    If Left$(typed, 1) = "#" Then f.Pick CLng(Mid$(typed, 2)) Else f.TypeIn typed
    If into >= 0 Then f.PickInto into
    EditWindowSelfTest = f.Preview() & "|" & f.CanAccept() & "|" & f.Value()
    Unload f
End Function

' The same, reading the window's list instead: "lines#first line".
Public Function EditListSelfTest(ByVal mode As String, ByVal cells As Range, ByVal typed As String, Optional ByVal into As Long = -1) As String
    Dim f As EditBox
    Set f = EditWindow(mode, cells)
    If typed <> "" Then
        If Left$(typed, 1) = "#" Then f.Pick CLng(Mid$(typed, 2)) Else f.TypeIn typed
    End If
    If into >= 0 Then f.PickInto into
    EditListSelfTest = f.DetailCount() & "#" & f.DetailLine(0)
    Unload f
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

' Everything - rows a filter hides too. Says how many first; nothing to do, no question.
Public Sub RevertAll()
    Dim n As Long
    n = Planner.EditCount()
    If n = 0 Then Application.StatusBar = TITLE & ":  nothing to revert - the sheet is as dumped.": Exit Sub
    If MsgBox("Put all " & n & " edited cell(s) back to their dumped values?" & vbCrLf & vbCrLf & _
              "(Save scenario first to keep these edits to come back to.)", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub
    RevertCells AllData(), True
End Sub

' Where each column of the main sheet is on the Dumped sheet, found by its NAME (row 2 of
' both): a column inserted on the main sheet cannot make a cell compare with - or go back
' to - another column's dumped value. Index: the main sheet's column; 0 = not dumped.
Public Function DumpedColumns() As Long()
    Dim ws As Worksheet, out() As Long, c As Long, lc As Long, f As Range, hdr As String
    lc = LastCol
    ReDim out(1 To Application.Max(1, lc))
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then DumpedColumns = out: Exit Function
    For c = 1 To lc
        hdr = CStr(MainSheet.Cells(HEADER_ROW, c).Value)
        If hdr <> "" Then
            If CStr(ws.Cells(HEADER_ROW, c).Value) = hdr Then
                out(c) = c                                  ' where it was dumped - the usual case
            Else
                Set f = ws.Rows(HEADER_ROW).Find(What:=hdr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
                If Not f Is Nothing Then out(c) = f.Column
            End If
        End If
    Next
    DumpedColumns = out
End Function

Public Function RevertCells(ByVal cells As Range, ByVal tell As Boolean) As Long
    Dim ws As Worksheet, c As Range, dr As Long, n As Long, dcol() As Long, dc As Long
    If cells Is Nothing Then Exit Function
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    dcol = DumpedColumns()
    Application.EnableEvents = False
    On Error GoTo Done                      ' events always come back on
    For Each c In cells.Cells
        dc = 0
        If c.Column <= UBound(dcol) Then dc = dcol(c.Column)
        If dc > 0 And Not Untracked(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) And Not c.HasFormula Then
            dr = DumpedRow(MainSheet.Cells(c.Row, 1).Value)
            If dr > 0 Then
                If Not SameValue(c.Value, ws.Cells(dr, dc).Value) Then
                    Panel.Journal c
                    c.Value = ws.Cells(dr, dc).Value
                    n = n + 1
                End If
            End If
        End If
    Next
Done:
    Application.EnableEvents = True
    RevertCells = n
    If tell Then Application.StatusBar = TITLE & ":  " & n & " cell(s) put back to their dumped value." & IIf(n > 0, "  Ctrl+Z to undo.", "")
End Function

' ============================================================ review

' A "Changes" sheet: every cell that differs from the dump, old and new.
Public Function ListChanges() As Long
    Dim ws As Worksheet, out As Worksheet, r As Long, c As Long, dr As Long, n As Long
    Dim hdr As String, lc As Long, cmt As Long, dcol() As Long
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
    dcol = DumpedColumns()
    For r = FIRST_ROW To LastRow
        dr = DumpedRow(MainSheet.Cells(r, 1).Value)
        If dr > 0 Then
            For c = 1 To lc
                hdr = CStr(MainSheet.Cells(HEADER_ROW, c).Value)
                If hdr <> "" And Not Untracked(hdr) And dcol(c) > 0 Then
                    If Not SameValue(MainSheet.Cells(r, c).Value, ws.Cells(dr, dcol(c)).Value) Then
                        n = n + 1
                        out.Cells(n + 1, 1).Value = MainSheet.Cells(r, 1).Value
                        out.Cells(n + 1, 2).Value = MainSheet.Cells(r, 2).Value
                        If cmt > 0 Then out.Cells(n + 1, 3).Value = MainSheet.Cells(r, cmt).Value
                        out.Cells(n + 1, 4).Value = hdr
                        out.Cells(n + 1, 5).Value = ws.Cells(dr, dcol(c)).Value
                        out.Cells(n + 1, 6).Value = MainSheet.Cells(r, c).Value
                        ' The parameter links to the cell, to go and look.
                        out.Hyperlinks.Add Anchor:=out.Cells(n + 1, 4), Address:="", _
                                           SubAddress:="'" & MAIN_SHEET & "'!" & MainSheet.Cells(r, c).Address(False, False), _
                                           ScreenTip:="Go to op " & MainSheet.Cells(r, 1).Value & " " & hdr
                    End If
                End If
            Next
        End If
    Next
    If n = 0 Then out.Cells(2, 1).Value = "No changes from the dump."
    out.Columns("A:F").AutoFit
    If n > 0 Then out.Range("A1:F1").AutoFilter
    ListChanges = n
End Function

Public Sub ShowChanges()
    Dim n As Long
    n = ListChanges()
    ThisWorkbook.Worksheets("Changes").Activate
End Sub

' ============================================================ calculators

' The speed and feed calculator, filled from the row the active cell is on.
Public Sub CalcSpeed()
    Dim f As Calculator
    Set f = CalcWindow(ActiveCell)
    f.Show
    Unload f
End Sub

Public Sub CalcFeed()
    CalcSpeed
End Sub

' The calculator set up from a cell's row (speed, feed, units, max_ss), not shown.
Public Function CalcWindow(ByVal c As Range) As Calculator
    Dim f As Calculator, r As Long, sm As String, sp As Variant, fm As String, fd As Variant, mx As Variant
    Dim surf As String, rpm As String, rev As String, pm As String, info As String, metric As Boolean
    Set f = New Calculator
    info = "Type in any box - the others follow."
    If Not c Is Nothing Then
        If c.Worksheet.Name = MAIN_SHEET And c.Row >= FIRST_ROW And c.Row <= LastRow Then
            r = c.Row
            sp = ValueAt(r, "speed"): sm = CStr(ValueAt(r, "speed_mode"))
            fd = ValueAt(r, "feed"): fm = CStr(ValueAt(r, "feed_mode"))
            mx = ValueAt(r, "max_ss")
            metric = (CStr(ValueAt(r, "units")) = "mm")
            If IsNumeric(sp) And Not IsEmpty(sp) Then If sm = "CSS" Then surf = CStr(sp) Else rpm = CStr(sp)
            If IsNumeric(fd) And Not IsEmpty(fd) Then If fm = "per min" Then pm = CStr(fd) Else rev = CStr(fd)
            info = "From op " & MainSheet.Cells(r, 1).Value & " (" & MainSheet.Cells(r, 2).Value & ")" & _
                   IIf(surf <> "", " - CSS: type a diameter for the RPM there.", ".") & "  Type in any box - the others follow."
        End If
    End If
    f.Setup "", surf, rpm, rev, pm, metric, IIf(IsNumeric(mx) And Not IsEmpty(mx), Val(CStr(mx)), 0), info
    If r > 0 And ColOf("feed") > 0 Then f.SetRow MainSheet.Cells(r, ColOf("feed")), (fm = "per min"), RoundInsertDia(r, metric), DepthOf(r)
    Set CalcWindow = f
End Function

' A round insert's diameter from the row's tool radius - only when the radius is a round
' insert's (over 0.07 in / 1.8 mm; a nose radius is smaller), else "".
Private Function RoundInsertDia(ByVal r As Long, ByVal metric As Boolean) As String
    Dim rad As Variant
    rad = ValueAt(r, "tool_radius")
    If IsEmpty(rad) Or Not IsNumeric(rad) Then Exit Function
    If CDbl(rad) > IIf(metric, 1.8, 0.07) Then RoundInsertDia = CStr(Round(2 * CDbl(rad), 4))
End Function

' The row's depth per pass: stepover (dynamic), else rough_step, else depth.
Private Function DepthOf(ByVal r As Long) As String
    Dim v As Variant, nm As Variant
    For Each nm In Array("stepover", "rough_step", "depth")
        v = ValueAt(r, CStr(nm))
        If Not IsEmpty(v) And IsNumeric(v) Then
            If CDbl(v) > 0 Then DepthOf = CStr(v): Exit Function
        End If
    Next
End Function

' For tools\check_macros.ps1 (no window): the chip part of the calculator - fill from a
' cell, type "box=value;...", optionally press "!btnUseFeed" / "!btnSetRow";
' read back "insert dia|depth|chip|feed to program|thinning line|set-row button".
Public Function ChipSelfTest(ByVal c As Range, ByVal typed As String) As String
    Dim f As Calculator, t As Variant, kv() As String
    Set f = CalcWindow(c)
    For Each t In Split(typed, ";")
        If Left$(t, 1) = "!" Then
            f.Press Mid$(t, 2)
        ElseIf InStr(t, "=") > 0 Then
            kv = Split(t, "=")
            f.TypeIn kv(0), kv(1)
        End If
    Next
    ChipSelfTest = f.Field("txtIC") & "|" & f.Field("txtAp") & "|" & f.Field("txtHex") & "|" & f.Field("txtFn") & "|" & _
                   f.Field("lblThin") & "|" & f.Field("btnSetRow")
    Unload f
End Function

' For tools\check_macros.ps1 (no window): fill from a cell, type "box=value;box=value",
' read back "rpm|surface|per rev|per min|note".
Public Function CalcSelfTest(ByVal c As Range, ByVal typed As String) As String
    Dim f As Calculator, t As Variant, kv() As String
    Set f = CalcWindow(c)
    For Each t In Split(typed, ";")
        If InStr(t, "=") > 0 Then
            kv = Split(t, "=")
            f.TypeIn kv(0), kv(1)
        End If
    Next
    CalcSelfTest = f.Field("txtRpm") & "|" & f.Field("txtSurf") & "|" & f.Field("txtRev") & "|" & _
                   f.Field("txtMin") & "|" & f.Field("lblNote")
    Unload f
End Function

Public Function RpmFromSurface(ByVal dia As Double, ByVal surface As Double, ByVal metric As Boolean) As Double
    If dia > 0 Then RpmFromSurface = IIf(metric, 1000, 12) * surface / (Application.Pi() * dia)
End Function

Public Function SurfaceFromRpm(ByVal dia As Double, ByVal rpm As Double, ByVal metric As Boolean) As Double
    SurfaceFromRpm = rpm * Application.Pi() * dia / IIf(metric, 1000, 12)
End Function

' ============================================================ pickers

Private Function IsCoolantHeader(ByVal hdr As String) As Boolean
    IsCoolantHeader = (hdr = "coolant_before" Or hdr = "coolant_with" Or hdr = "coolant_after")
End Function

Private Function AllCoolant(ByVal cells As Range) As Boolean
    Dim c As Range
    For Each c In cells.Cells
        If Not IsCoolantHeader(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) Then Exit Function
    Next
    AllCoolant = True
End Function

' A coolant cell's choices: its dropdown list ("none" first), or an empty array.
Private Function CoolantChoices(ByVal c As Range) As Variant
    Dim list As String
    On Error Resume Next
    list = c.Validation.Formula1
    On Error GoTo 0
    list = Replace(list, """", "")
    If list = "" Or UCase$(list) = "FALSE" Or UCase$(list) = "=FALSE" Then
        CoolantChoices = Array()
    Else
        CoolantChoices = Split(list, ",")
    End If
End Function

' The coolant window for some cells: one tick box per coolant any of their machines has
' (marked * when not every row's machine has it), ticked as the cells are now when they
' all agree. Not shown - the caller shows it.
Private Function CoolantForm(ByVal cells As Range, ByVal what As String) As CoolantPicker
    Dim c As Range, opts As Variant, i As Long, k As Long, offer() As String, hits() As Long, n As Long
    Dim partial() As Boolean, current As String, same As Boolean, f As CoolantPicker
    ReDim offer(0 To 0): ReDim hits(0 To 0)
    same = True
    For Each c In cells.Cells
        opts = CoolantChoices(c)
        For i = 1 To UBound(opts)                         ' 0 is "none"
            For k = 0 To n - 1
                If LCase$(offer(k)) = LCase$(Trim$(opts(i))) Then Exit For
            Next
            If k = n Then
                ReDim Preserve offer(0 To n): ReDim Preserve hits(0 To n)
                offer(n) = Trim$(opts(i)): n = n + 1
            End If
            hits(k) = hits(k) + 1
        Next
        If c.Address <> cells.Cells(1, 1).Address Then
            If LCase$(CStr(c.Value)) <> LCase$(CStr(cells.Cells(1, 1).Value)) Then same = False
        End If
    Next
    If n = 0 Then Exit Function
    ReDim partial(0 To n - 1)
    For k = 0 To n - 1
        partial(k) = (hits(k) < cells.Count)
    Next
    If same Then current = CStr(cells.Cells(1, 1).Value)
    Set f = New CoolantPicker
    f.Setup offer, partial, current, what & ":  tick the coolants to turn on.  None ticked = none."
    Set CoolantForm = f
End Function

' Ask with the coolant window; the names joined by " + ", "none", or "" if cancelled.
Public Function AskCoolants(ByVal cells As Range, ByVal what As String) As String
    Dim f As CoolantPicker
    Set f = CoolantForm(cells, what)
    If f Is Nothing Then
        MsgBox "No coolant choices here - is this a row with coolant?", vbInformation, TITLE
        Exit Function
    End If
    f.Show
    If f.Accepted Then AskCoolants = f.Result()
    Unload f
End Function

' For tools\check_macros.ps1 (no window): the boxes offered, then what ticking boxes
' "1,3" (positions) would set - "captions|...#result".
Public Function CoolantPickerSelfTest(ByVal cells As Range, ByVal ticks As String) As String
    Dim f As CoolantPicker, t As Variant, i As Long
    Set f = CoolantForm(cells, "test")
    If f Is Nothing Then Exit Function
    If ticks <> "" Then
        For i = 1 To 12
            f.Tick i, False
        Next
        For Each t In Split(ticks, ",")
            f.Tick CLng(t), True
        Next
    End If
    CoolantPickerSelfTest = f.Captions() & "#" & f.Result()
    Unload f
End Function

' Coolant names into cells, each row checked against ITS OWN choices (machines differ).
' "A + B" is not a dropdown choice but is what the load reads - so the check is by name.
Public Function SetCoolantCells(ByVal cells As Range, ByVal names As String) As String
    Dim c As Range, ok As Long, bad As Long, opts As Variant, want() As String, i As Long, j As Long, found As Boolean, good As Boolean
    want = Split(names, " + ")
    For Each c In cells.Cells
        opts = CoolantChoices(c)
        good = UBound(opts) >= 0
        For i = 0 To UBound(want)
            found = False
            For j = 0 To UBound(opts)
                If LCase$(Trim$(opts(j))) = LCase$(Trim$(want(i))) Then found = True
            Next
            If Not found Then good = False
        Next
        If good Then
            Panel.Journal c
            c.Value = names
            ok = ok + 1
        Else
            bad = bad + 1
        End If
    Next
    SetCoolantCells = ok & " " & bad
End Function

' Several coolants at one timing, for the active cell.
Public Sub PickCoolant()
    Dim c As Range, names As String
    If Not OnMain() Then Exit Sub
    Set c = ActiveCell
    If Not IsCoolantHeader(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) Or c.Row < FIRST_ROW Then
        MsgBox "Click a cell in coolant_before, coolant_with or coolant_after first.", vbInformation, TITLE
        Exit Sub
    End If
    names = AskCoolants(c, "op " & MainSheet.Cells(c.Row, 1).Value & "  " & MainSheet.Cells(HEADER_ROW, c.Column).Value)
    If names <> "" Then SetCoolantCells c, names
End Sub

' ============================================================ manual entry text

' Whether a cell takes input at all (read-only and does-not-apply cells refuse everything).
Public Function TakesInput(ByVal c As Range) As Boolean
    Dim f As String
    TakesInput = True
    On Error Resume Next
    f = UCase$(Replace(c.Validation.Formula1, "=", ""))
    On Error GoTo 0
    If f = "FALSE" Then TakesInput = False
End Function

' The columns the text editor opens on (double-click, or the ribbon's Edit text).
Public Function IsTextColumn(ByVal hdr As String) As Boolean
    IsTextColumn = (hdr = "manual_text" Or hdr = "comment" Or hdr = "insp_comment")
End Function

Public Sub EditManualText()
    Dim c As Range
    If Not OnMain() Then Exit Sub
    Set c = ActiveCell
    If Not IsTextColumn(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) Or c.Row < FIRST_ROW Then
        MsgBox "Click a comment, inspection comment or manual entry text cell first.", vbInformation, TITLE
        Exit Sub
    End If
    EditText c
End Sub

' The most characters a cell takes, from its own rule (0 = none known).
Private Function TextLimit(ByVal c As Range) As Long
    On Error Resume Next
    If c.Validation.Type = xlValidateTextLength Then TextLimit = CLng(c.Validation.Formula1)
End Function

' The editor window on one text cell. True when the text was changed.
Public Function EditText(ByVal c As Range) As Boolean
    Dim f As TextEditor, info As String, g As Variant, hdr As String, what As String
    hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
    If Not TakesInput(c) Then
        If hdr = "manual_text" Then
            MsgBox "This row is not a manual entry.", vbInformation, TITLE
        Else
            MsgBox "This operation has no " & IIf(hdr = "comment", "comment", "tool inspection comment") & " to edit.", vbInformation, TITLE
        End If
        Exit Function
    End If
    info = "op " & MainSheet.Cells(c.Row, 1).Value
    Set f = New TextEditor
    If hdr = "manual_text" Then
        g = ValueAt(c.Row, "manual_gcode")
        If CStr(g) = "1006" Then info = info & "  -  output as CODE" Else If CStr(g) = "1005" Then info = info & "  -  output as a COMMENT"
        info = info & "        (Ctrl+Enter = OK,  Esc = Cancel)"
        f.LoadText CStr(c.Value), info
    Else
        what = IIf(hdr = "comment", "operation comment", "tool inspection comment")
        If hdr = "insp_comment" Then info = info & "  -  shown at each inspection stop; ROTATE / FLIP / CHANGE / INDEX makes it an insert flip"
        info = info & "      (Enter = OK,  Esc = Cancel)"
        f.LoadText CStr(c.Value), info, TextLimit(c), what, True
    End If
    f.Show
    If f.Accepted Then
        If f.EditedText() <> CStr(c.Value) Then
            Panel.Journal c
            c.Value = f.EditedText()
            EditText = True
        End If
    End If
    Unload f
End Function

' For the checks (no window): the counter and the limit, on given text.
Public Function EditorSelfTest(ByVal text As String, Optional ByVal maxLen As Long = 0, Optional ByVal oneLine As Boolean = False) As String
    Dim f As TextEditor
    Set f = New TextEditor
    f.LoadText text, "test", maxLen, IIf(oneLine, "comment", ""), oneLine
    EditorSelfTest = f.CountText() & "|" & f.CanAccept() & "|" & Len(f.EditedText())
    Unload f
End Function

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
        Panel.Journal MainSheet.Cells(c.Row, pc)
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
        Panel.Journal MainSheet.Cells(c.Row, ac)
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

' The edit commands' ribbon buttons are in Panel.bas, with undo around each.
Public Sub RbChanges(control As IRibbonControl): ShowChanges: End Sub
