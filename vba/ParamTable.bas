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

' The op windows that stay open beside the sheet - one of each (ShowTextWindow,
' ShowCoolantWindow); Nothing when closed.
Private textWin As TextEditor
Private coolWin As CoolantPicker

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

' What a command did, in the status bar - the window already showed what would happen,
' so no box to click away.
Private Sub Report(ByVal what As String, ByVal done As Long, ByVal refused As Long)
    Dim s As String
    s = done & " cell(s) " & what & "."
    If refused > 0 Then s = s & "  " & refused & " refused - read-only, not for that kind of operation, or outside its limits."
    Application.StatusBar = TITLE & ":  " & s
End Sub

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
        If names <> "" Then ReportPair "set to " & names, SetCoolantCells(cells, names)
        Exit Sub
    End If
    Set f = EditWindow("set", cells)
    f.Show
    If f.Accepted Then ReportPair "set to " & f.Value(), SetCells(cells, f.Value())
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

Private Sub ReportPair(ByVal what As String, ByVal pair As String)
    Dim p() As String
    p = Split(pair, " ")
    Report what, CLng(p(0)), CLng(p(1))
End Sub

' Scale every selected number by a percentage (110 = 10% more).
Public Sub ScaleSelected()
    Dim cells As Range, f As EditBox
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    Set f = EditWindow("scale", cells)
    f.Show
    If f.Accepted Then ReportPair "scaled to " & f.Value() & "%", ScaleCells(cells, CDbl(f.Value()))
    Unload f
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
    Dim cells As Range, f As EditBox
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    Set f = EditWindow("copy", cells)
    f.Show
    If f.Accepted Then ReportPair "copied from op " & f.Value(), CopyCells(cells, f.Value())
    Unload f
End Sub

Public Function CopyCells(ByVal cells As Range, ByVal op As Variant) As String
    Dim f As Range, c As Range, ok As Long, bad As Long, src As Range
    Set f = OpCell(op)
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

' An operation's op_idn cell, or Nothing.
Public Function OpCell(ByVal op As Variant) As Range
    If Trim$(CStr(op)) = "" Then Exit Function
    If IsNumeric(op) Then op = CDbl(op)
    Set OpCell = MainSheet.Range(MainSheet.Cells(FIRST_ROW, 1), MainSheet.Cells(LastRow, 1)).Find( _
                     What:=op, LookIn:=xlValues, LookAt:=xlWhole)
End Function

' ---------------------------------------------------------------- the edit window

' "N cell(s) in feed, speed  -  ops 2, 7" - what is selected, in words.
Private Function Describe(ByVal cells As Range) As String
    Dim c As Range, cols As String, ops As String, h As String, o As String, nc As Long, no As Long
    For Each c In cells.Cells
        h = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
        If InStr("|" & cols & "|", "|" & h & "|") = 0 Then
            nc = nc + 1
            If nc <= 4 Then cols = cols & IIf(cols = "", "", "|") & h
        End If
        o = CStr(MainSheet.Cells(c.Row, 1).Value)
        If InStr("|" & ops & "|", "|" & o & "|") = 0 Then
            no = no + 1
            If no <= 8 Then ops = ops & IIf(ops = "", "", "|") & o
        End If
    Next
    Describe = cells.Count & " cell(s) in " & Replace(cols, "|", ", ") & IIf(nc > 4, " and " & nc - 4 & " more", "") & _
               "  -  op" & IIf(no = 1, " ", "s ") & Replace(ops, "|", ", ") & IIf(no > 8, ", ...", "")
End Function

' The window for one of the edits, set up on the cells but not shown. Modes: set, scale,
' copy (from an op), tool (to every op of a tool), scensave / scenload / scendel.
Public Function EditWindow(ByVal mode As String, ByVal cells As Range) As EditBox
    Dim f As EditBox, c As Range, lst As String, sameList As Boolean, items() As String, k() As String
    Dim r As Long, n As Long, cmt As Long, v As String, seen As String, ks As Variant, its As Variant, nm As Variant
    Set f = New EditBox
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
            f.Setup "set", cells, TITLE & " - set selected", Describe(cells), "Pick the value for every selected cell:", _
                    Split(lst, ","), Empty, True, ""
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
                f.Setup "set", cells, TITLE & " - set selected", Describe(cells), _
                        "Value for every selected cell (the list has what is there now):", items, Empty, False, CStr(cells.Cells(1, 1).Value)
            Else
                f.Setup "set", cells, TITLE & " - set selected", Describe(cells), _
                        "Value for every selected cell:", Empty, Empty, False, ""
            End If
        End If
    Case "scale"
        f.Setup "scale", cells, TITLE & " - scale", Describe(cells), _
                "Scale the selected numbers to what percent?  (110 = 10% more, 90 = 10% less)", _
                Array("50", "80", "90", "95", "105", "110", "120", "150", "200"), Empty, False, "100"
    Case "copy"
        cmt = ColOf("comment")
        For r = FIRST_ROW To LastRow
            ReDim Preserve items(0 To n): ReDim Preserve k(0 To n)
            k(n) = CStr(MainSheet.Cells(r, 1).Value)
            items(n) = "op " & k(n) & "    " & MainSheet.Cells(r, 2).Value & _
                       IIf(cmt > 0, "    " & MainSheet.Cells(r, IIf(cmt > 0, cmt, 1)).Value, "")
            n = n + 1
        Next
        f.Setup "copy", cells, TITLE & " - copy from an operation", Describe(cells), _
                "Copy the selected columns FROM which operation?", items, k, True, ""
    Case "tool"
        ' From the first selected row (the one being copied), into the rest of its tool.
        r = cells.Cells(1, 1).Row
        its = Planner.ToolOps(r, ks)
        f.Setup "tool", cells, TITLE & " - apply to every op of the tool", _
                "Copy " & ColumnsOf(cells) & " to every other op of tool " & Planner.ToolOfRow(r) & " (hidden rows too).", _
                "Copy FROM which op of the tool?", its, ks, True, CStr(MainSheet.Cells(r, 1).Value)
    Case "scensave", "scenload", "scendel"
        its = Planner.ScenarioNames()
        Select Case mode
        Case "scensave"
            f.Setup mode, Nothing, TITLE & " - save scenario", _
                    "Keeps every edit now on the sheet under a name, to restore or compare later (saved in this workbook).", _
                    "Name (pick one to replace it):", its, Empty, False, ""
        Case "scenload"
            v = ""
            If UBound(its) >= 0 Then v = CStr(its(UBound(its)))      ' the latest
            f.Setup mode, Nothing, TITLE & " - restore scenario", _
                    "Puts the sheet back to the dump, then sets the scenario's values - each through its cell's rule.", _
                    "Restore which scenario?", its, Empty, True, v
        Case Else
            f.Setup mode, Nothing, TITLE & " - delete scenario", "Deletes a saved scenario; the sheet is not changed.", _
                    "Delete which scenario?", its, Empty, True, ""
        End Select
    End Select
    Set EditWindow = f
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

' The cells an edit writes: the selection, or for "tool" the rest of the tool's ops.
Private Function EditTargets(ByVal mode As String, ByVal cells As Range, ByVal v As String) As Range
    If mode = "tool" Then Set EditTargets = Planner.ToolTargets(cells, v) Else Set EditTargets = cells
End Function

' What an edit does to one cell: 0 = leaves it, 1 = changes it, 2 = already that,
' 3 = refused by its rule. nv is the new value.
Private Function CellPlan(ByVal mode As String, ByVal c As Range, ByVal v As String, ByVal src As Range, ByRef nv As Variant) As Long
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
        nv = Round(CDbl(c.Value) * CDbl(v) / 100, 10)
        If Not WouldTake(c, nv) Then
            CellPlan = 3
        ElseIf nv = c.Value Then
            CellPlan = 2
        Else
            CellPlan = 1
        End If
    Case "copy", "tool"
        If c.Row = src.Row Then Exit Function
        If MainSheet.Cells(src.Row, c.Column).HasFormula Then Exit Function
        nv = MainSheet.Cells(src.Row, c.Column).Value
        If Not WouldTake(c, nv) Then
            CellPlan = 3
        ElseIf SameValue(c.Value, nv) Then
            CellPlan = 2
        Else
            CellPlan = 1
        End If
    End Select
End Function

' The edit window's live line: "1|..." when OK would change something, "0|..." when not.
Public Function EditPreview(ByVal mode As String, ByVal cells As Range, ByVal v As String) As String
    Dim c As Range, take As Long, bad As Long, same As Long, nv As Variant, eg As String, src As Range, shown As Long
    Dim targets As Range, ops As String, nops As Long, o As String
    Select Case mode
    Case "scensave", "scenload", "scendel"
        EditPreview = Planner.ScenarioPreview(mode, v)
        Exit Function
    Case "set"
        If Trim$(v) = "" Then EditPreview = "0|Type or pick a value.": Exit Function
    Case "scale"
        If Trim$(v) = "" Or Not IsNumeric(v) Then EditPreview = "0|Type a percent, e.g. 110.": Exit Function
        If CDbl(v) <= 0 Then EditPreview = "0|The percent has to be more than 0.": Exit Function
    Case "copy", "tool"
        Set src = OpCell(v)
        If src Is Nothing Then EditPreview = "0|Pick the operation to copy from.": Exit Function
    End Select
    Set targets = EditTargets(mode, cells, v)
    If targets Is Nothing Then EditPreview = "0|No other operation uses this tool.": Exit Function
    For Each c In targets.Cells
        Select Case CellPlan(mode, c, v, src, nv)
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
        End Select
    Next
    If eg <> "" Then eg = vbCrLf & "e.g.  " & eg
    If mode = "tool" And take > 0 Then eg = vbCrLf & "Into op" & IIf(nops = 1, " ", "s ") & Replace(ops, "|", ", ") & IIf(nops > 8, ", ..." , "") & "."
    EditPreview = IIf(take > 0, "1|", "0|") & take & " cell(s) will change" & _
                  IIf(same > 0, ",  " & same & " already that", "") & _
                  IIf(bad > 0, ",  " & bad & " refused (read-only, not for that operation, or outside its limits)", "") & _
                  "." & eg
End Function

' The edit window's list: one line per cell it would change or refuse.
Public Function EditDetail(ByVal mode As String, ByVal cells As Range, ByVal v As String) As Variant
    Dim c As Range, nv As Variant, s As String, n As Long, src As Range, targets As Range, st As Long, hdr As String
    EditDetail = Array()
    Select Case mode
    Case "scensave", "scenload", "scendel"
        EditDetail = Planner.ScenarioDetail(IIf(mode = "scensave", "", v))
        Exit Function
    Case "set"
        If Trim$(v) = "" Then Exit Function
    Case "scale"
        If Trim$(v) = "" Or Not IsNumeric(v) Then Exit Function
        If CDbl(v) <= 0 Then Exit Function
    Case "copy", "tool"
        Set src = OpCell(v)
        If src Is Nothing Then Exit Function
    End Select
    Set targets = EditTargets(mode, cells, v)
    If targets Is Nothing Then Exit Function
    For Each c In targets.Cells
        st = CellPlan(mode, c, v, src, nv)
        If st = 1 Or st = 3 Then
            n = n + 1
            If n <= 300 Then
                hdr = CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)
                s = s & IIf(s = "", "", vbLf) & "op " & MainSheet.Cells(c.Row, 1).Value & "    " & hdr & "    " & _
                    IIf(st = 1, IIf(CStr(c.Value) = "", "(blank)", CStr(c.Value)) & "  ->  " & CStr(nv), "refused  (" & CStr(nv) & ")")
            End If
        End If
    Next
    If n > 300 Then s = s & vbLf & "... and " & n - 300 & " more"
    If s <> "" Then EditDetail = Split(s, vbLf)
End Function

' Copy the selected columns of the active op to every other op of its tool.
Public Sub ApplyToTool()
    Dim cells As Range, f As EditBox
    If Not OnMain() Then Exit Sub
    Set cells = DataCells(Selection)
    If cells Is Nothing Then Exit Sub
    If ColOf("tool") = 0 Or Planner.ToolOfRow(cells.Cells(1, 1).Row) = "" Then
        MsgBox "Click the cells to copy (feed, speed ...) in an operation's row that has a tool number.", vbInformation, TITLE
        Exit Sub
    End If
    Set f = EditWindow("tool", cells)
    f.Show
    If f.Accepted Then ReportPair "copied from op " & f.Value() & " to the other ops of its tool", Planner.ToolCopyCells(cells, f.Value())
    Unload f
End Sub

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
                                    IIf(p(1) <> "0", ", " & p(1) & " refused (the op or column is gone, or the value is outside its limits)", "") & "."
        Case "scendel"
            Planner.DeleteScenario f.Value()
            Application.StatusBar = TITLE & ":  '" & f.Value() & "' deleted."
        End Select
    End If
    Unload f
End Sub

' For tools\check_macros.ps1 (no window): what the window says and whether OK is on,
' after typing a value ("#3" picks list row 3) - "preview|ok|value".
Public Function EditWindowSelfTest(ByVal mode As String, ByVal cells As Range, ByVal typed As String) As String
    Dim f As EditBox
    Set f = EditWindow(mode, cells)
    If Left$(typed, 1) = "#" Then f.Pick CLng(Mid$(typed, 2)) Else f.TypeIn typed
    EditWindowSelfTest = f.Preview() & "|" & f.CanAccept() & "|" & f.Value()
    Unload f
End Function

' The same, reading the window's list instead: "lines#first line".
Public Function EditListSelfTest(ByVal mode As String, ByVal cells As Range, ByVal typed As String) As String
    Dim f As EditBox
    Set f = EditWindow(mode, cells)
    If typed <> "" Then
        If Left$(typed, 1) = "#" Then f.Pick CLng(Mid$(typed, 2)) Else f.TypeIn typed
    End If
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

Public Function RevertCells(ByVal cells As Range, ByVal tell As Boolean) As Long
    Dim ws As Worksheet, c As Range, dr As Long, n As Long
    If cells Is Nothing Then Exit Function
    Set ws = ThisWorkbook.Worksheets(DUMPED_SHEET)
    Application.EnableEvents = False
    On Error GoTo Done                      ' events always come back on
    For Each c In cells.Cells
        If Not Untracked(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value)) And Not c.HasFormula Then
            dr = DumpedRow(MainSheet.Cells(c.Row, 1).Value)
            If dr > 0 Then
                If Not SameValue(c.Value, ws.Cells(dr, c.Column).Value) Then
                    Panel.Journal c
                    c.Value = ws.Cells(dr, c.Column).Value
                    n = n + 1
                End If
            End If
        End If
    Next
Done:
    Application.EnableEvents = True
    RevertCells = n
    If tell Then Application.StatusBar = TITLE & ":  " & n & " cell(s) put back to their dumped value."
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

' ============================================================ op windows: shared

' The ops a window opens on: every row with a selected cell; with none, the row the cursor
' is on; not an op row either - none (the window then says to select some).
Public Function WindowRows() As Collection
    Dim rows As Collection, r As Long
    Set rows = New Collection
    On Error Resume Next
    If TypeName(Selection) = "Range" Then Set rows = Panel.OpRows(Selection)
    If rows.Count = 0 And ActiveSheet.Name = MAIN_SHEET Then
        r = ActiveCell.Row
        If r >= FIRST_ROW And r <= LastRow Then
            If Not MainSheet.Rows(r).Hidden Then rows.Add r
        End If
    End If
    Set WindowRows = rows
End Function

' The op above or below a row (way -1 / 1), hidden rows skipped; 0 when there is none.
Public Function NextOpRow(ByVal r As Long, ByVal way As Long) As Long
    Dim last As Long
    last = LastRow
    Do
        r = r + way
        If r < FIRST_ROW Or r > last Then Exit Function
    Loop While MainSheet.Rows(r).Hidden
    NextOpRow = r
End Function

' "op 7, op 11" - a few ops by number, the rest as "and 3 more".
Public Function OpNames(ByVal rows As Collection) As String
    Dim i As Long, s As String
    For i = 1 To Application.Min(rows.Count, 5)
        s = s & IIf(s = "", "", ", ") & "op " & MainSheet.Cells(rows(i), 1).Value
    Next
    If rows.Count > 5 Then s = s & " and " & rows.Count - 5 & " more"
    OpNames = s
End Function

' A cell's value as text ("" for an error value).
Public Function CellText(ByVal c As Range) As String
    If Not IsError(c.Value) Then CellText = CStr(c.Value)
End Function

' "Flood x3, Off x1" - each value and how many ops have it, most first (a blank cell is
' "(blank)").
Public Function CountsText(ByVal values As Collection) As String
    Dim nm() As String, n() As Long, k As Long, v As Variant, i As Long, j As Long, best As Long, s As String
    ReDim nm(1 To values.Count + 1): ReDim n(1 To values.Count + 1)
    For Each v In values
        If Trim$(CStr(v)) = "" Then v = "(blank)"
        For i = 1 To k
            If LCase$(nm(i)) = LCase$(CStr(v)) Then Exit For
        Next
        If i > k Then
            k = k + 1
            nm(k) = CStr(v)
        End If
        n(i) = n(i) + 1
    Next
    For j = 1 To k
        best = 0
        For i = 1 To k
            If n(i) > 0 Then
                If best = 0 Then best = i Else If n(i) > n(best) Then best = i
            End If
        Next
        s = s & IIf(s = "", "", ", ") & nm(best) & " x" & n(best)
        n(best) = 0
    Next
    CountsText = s
End Function

' Count an op once in an Apply's "ops set".
Public Sub CountOp(ByVal done As Collection, ByVal r As Long)
    On Error Resume Next
    done.Add r, CStr(r)
End Sub

' An op window's result line after Apply: "3 ops set - Ctrl+Z to undo." and the ops
' refused, with why ("op 7 (with the move): this machine has no Thru-tool").
Public Function ApplyResult(ByVal done As Collection, ByVal bad As Collection, ByVal why As String) As String
    Dim s As String, i As Long
    If done.Count > 0 Then
        s = done.Count & IIf(done.Count = 1, " op", " ops") & " set - Ctrl+Z to undo."
    ElseIf bad.Count = 0 And why = "" Then
        s = "Nothing to change - the ops already have that."
    Else
        s = "Nothing set."
    End If
    If bad.Count > 0 Then
        s = s & vbCrLf & "Not set:  "
        For i = 1 To Application.Min(bad.Count, 4)
            s = s & IIf(i = 1, "", ";  ") & bad(i)
        Next
        If bad.Count > 4 Then s = s & ";  and " & bad.Count - 4 & " more"
    End If
    If why <> "" Then s = s & vbCrLf & "Stopped by an error: " & why
    ApplyResult = s
End Function

' A window closed: forget it, so the next press opens a fresh one.
Public Sub WindowClosed(ByVal f As Object)
    If Not textWin Is Nothing Then
        If f Is textWin Then Set textWin = Nothing
    End If
    If Not coolWin Is Nothing Then
        If f Is coolWin Then Set coolWin = Nothing
    End If
End Sub

' Whether a cell takes input at all (read-only and does-not-apply cells refuse everything).
Public Function TakesInput(ByVal c As Range) As Boolean
    Dim f As String
    TakesInput = True
    On Error Resume Next
    f = UCase$(Replace(c.Validation.Formula1, "=", ""))
    On Error GoTo 0
    If f = "FALSE" Then TakesInput = False
End Function

' The most characters a cell takes, from its own rule (0 = none known).
Public Function TextLimit(ByVal c As Range) As Long
    On Error Resume Next
    If c.Validation.Type = xlValidateTextLength Then TextLimit = CLng(c.Validation.Formula1)
End Function

' ============================================================ coolant

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
Public Function CoolantChoices(ByVal c As Range) As Variant
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

' The Coolant window on the ops selected - one of it: pressed again it comes to the front
' and follows the selection.
Public Sub ShowCoolantWindow()
    If coolWin Is Nothing Then Set coolWin = New CoolantPicker
    coolWin.ShowOn WindowRows()
End Sub

' Set selected on coolant cells: the Coolant window, on their ops. It sets the coolant
' itself (Apply), so nothing comes back for Set to write.
Public Function AskCoolants(ByVal cells As Range, ByVal what As String) As String
    If coolWin Is Nothing Then Set coolWin = New CoolantPicker
    coolWin.ShowOn Panel.OpRows(cells)
End Function

' Coolant names into one X-style cell, checked against THAT cell's own choices (machines
' differ). "A + B" and "A off" are not dropdown choices but are what the load reads, so
' the check is by name; a single choice (or none) still goes through the cell's rule.
' "" when written, else why not.
Public Function WriteCoolant(ByVal c As Range, ByVal names As String) As String
    Dim opts As Variant, want() As String, i As Long, j As Long, found As Boolean, nm As String, offs As Boolean
    opts = CoolantChoices(c)
    If UBound(opts) < 0 Or Not TakesInput(c) Then WriteCoolant = "no coolant on this op": Exit Function
    want = Split(names, "+")
    For i = 0 To UBound(want)
        nm = Trim$(want(i))
        If LCase$(Right$(nm, 4)) = " off" Then
            nm = Trim$(Left$(nm, Len(nm) - 4))
            offs = True
        End If
        found = False
        For j = 0 To UBound(opts)
            If LCase$(Trim$(opts(j))) = LCase$(nm) Then found = True
        Next
        If Not found Then WriteCoolant = "this machine has no " & nm: Exit Function
    Next
    If UBound(want) = 0 And Not offs Then
        If Not TryWrite(c, Trim$(names)) Then WriteCoolant = "the cell does not take " & Trim$(names)
    Else
        Panel.Journal c                     ' for undo
        c.Value = names
    End If
End Function

' Coolant names into cells, each row checked against ITS OWN choices: "done refused".
Public Function SetCoolantCells(ByVal cells As Range, ByVal names As String) As String
    Dim c As Range, ok As Long, bad As Long
    For Each c In cells.Cells
        If WriteCoolant(c, names) = "" Then ok = ok + 1 Else bad = bad + 1
    Next
    SetCoolantCells = ok & " " & bad
End Function

' For tools\check_macros.ps1 (no window shown): the Coolant window on a range's ops, then
' steps joined by ";" (CoolantPicker.Act: "v9=Flood", "with/Flood=1", "!btnApply",
' "follow=A5" ...), read back: the fields asked for, joined by "|" (CoolantPicker.Field).
Public Function CoolantWindowSelfTest(ByVal r As Range, ByVal steps As String, ByVal fields As String) As String
    Dim f As CoolantPicker, t As Variant, s As String
    Set f = New CoolantPicker
    f.LoadRows Panel.OpRows(r)
    If steps <> "" Then
        For Each t In Split(steps, ";")
            f.Act CStr(t)
        Next
    End If
    For Each t In Split(fields, ",")
        s = s & IIf(s = "", "", "|") & f.Field(CStr(t))
    Next
    CoolantWindowSelfTest = s
    Unload f
End Function

' ============================================================ comments and text

' The tab of the Comments & text window for a column: "comment", "insp", "manual", or "".
Public Function TextTab(ByVal hdr As String) As String
    Select Case hdr
        Case "comment": TextTab = "comment"
        Case "insp_comment": TextTab = "insp"
        Case "manual_text": TextTab = "manual"
    End Select
End Function

' The columns the Comments & text window opens on with a double-click.
Public Function IsTextColumn(ByVal hdr As String) As Boolean
    IsTextColumn = (TextTab(hdr) <> "")
End Function

' The Comments & text window on the ops selected - one of it: pressed again it comes to
' the front and follows the selection. Its tab: "comment", "insp", "manual"; "" = the
' one for the column the cursor is in (else as it was).
Public Sub ShowTextWindow(Optional ByVal tabName As String = "")
    On Error Resume Next
    If tabName = "" And ActiveSheet.Name = MAIN_SHEET Then tabName = TextTab(CStr(MainSheet.Cells(HEADER_ROW, ActiveCell.Column).Value))
    On Error GoTo 0
    If textWin Is Nothing Then Set textWin = New TextEditor
    textWin.ShowOn WindowRows(), tabName
End Sub

' Double-click on a comment, inspection comment or manual entry text cell: the window, on
' that tab, for that op. True when it is one of those cells.
Public Function EditText(ByVal c As Range) As Boolean
    Dim tb As String
    tb = TextTab(CStr(MainSheet.Cells(HEADER_ROW, c.Column).Value))
    If tb = "" Or c.Row < FIRST_ROW Then Exit Function
    ShowTextWindow tb
    EditText = True
End Function

' For tools\check_macros.ps1 (no window shown): the Comments & text window on a range's
' ops - on a tab, or "" for the tab of the range's column (as a double-click opens it) -
' then steps joined by ";" (TextEditor.Act: "txtComment=...", "!btnApply", "follow=A5"
' ...), read back: the fields asked for, joined by "|" (TextEditor.Field).
Public Function TextWindowSelfTest(ByVal r As Range, ByVal tabName As String, ByVal steps As String, ByVal fields As String) As String
    Dim f As TextEditor, t As Variant, s As String
    If tabName = "" Then tabName = TextTab(CStr(MainSheet.Cells(HEADER_ROW, r.Cells(1, 1).Column).Value))
    Set f = New TextEditor
    f.LoadRows Panel.OpRows(r), tabName
    If steps <> "" Then
        For Each t In Split(steps, ";")
            f.Act CStr(t)
        Next
    End If
    For Each t In Split(fields, ",")
        s = s & IIf(s = "", "", "|") & f.Field(CStr(t))
    Next
    TextWindowSelfTest = s
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
