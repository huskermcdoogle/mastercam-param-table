Attribute VB_Name = "Panel"
' Parameter Table Tool - what every command and op window shares, so they all work the
' same way:
'
'   UNDO      Every command records each cell it writes (BeginEdit ... EndEdit around it;
'             ParamTable.TryWrite and every other write calls Journal first). Ctrl+Z, or
'             the ribbon's Undo, puts the last command's cells back - only cells still as
'             the command left them, so a value typed since is never lost. Up to 20 back.
'   OP ROWS   A window works on "the ops selected": every row with a selected cell, any
'             column (OpRows). OpLabel / RowsText say which, in words.
'   FOLLOW    An op window stays open beside the sheet (shown modeless) and follows the
'             selection: Opened Me when it shows, Closed Me when it goes; SelectionMoved
'             calls its Public Sub FollowRows(rows As Collection) when other ops are picked.
'   FOCUS     ShowColumns "coolant|coolant_before" rows - opens the columns' groups,
'             selects those rows' cells there and scrolls to them, so the sheet shows what
'             the window is about. GoToGroup does the same for a whole column group and
'             folds the others away (the ribbon's Go to).
'   HELP      ShowHelp "coolant" opens help\coolant.html of the user manual beside the
'             add-in (a window's ? button). Topic names are the manual's page names.
'   ROW MARK  The op row the cursor is on gets a line above and below (and a tint on the
'             frozen columns), so the place is not lost on a wide sheet.
'
' Plain text in the repo (vba\Panel.bas); tools\build_vba.ps1 compiles every vba\*.bas.
Option Explicit

Private Const TITLE As String = "Parameter Table"
Private Const HEADER_ROW As Long = 2
Private Const FIRST_ROW As Long = 3
Private Const MAX_UNDO As Long = 20

' ============================================================ undo

Private jOn As Boolean
Private jLabel As String
Private jEntries As Collection          ' Array(row, col, hadFormula, old)
Private jSeen As Collection             ' "row|col" keys already kept
Private undoStack As Collection         ' Array(label, entries) - entries: Array(row, col, hadF, old, nowF, now)

Private wins As Collection              ' the open windows that follow the selection
Public Following As Boolean             ' a window is moving the selection itself
Private lastKey As String               ' the rows they were last told
Private marked As Boolean               ' the row mark is on the sheet
Private lastFind As String              ' what Find op looked for last

' A command is writing (its cells are being recorded for undo).
Public Function Recording() As Boolean
    Recording = jOn
End Function

' Start recording a command's writes. Nested calls keep the outer command.
Public Sub BeginEdit(ByVal label As String)
    If jOn Then Exit Sub
    jOn = True
    jLabel = label
    Set jEntries = New Collection
    Set jSeen = New Collection
End Sub

' Called before a cell of the main sheet is written: its value (or formula) as it was.
Public Sub Journal(ByVal c As Range)
    Dim k As String, cell As Range
    If Not jOn Then Exit Sub
    If c Is Nothing Then Exit Sub
    If c.Worksheet.Name <> ParamTable.MainSheet.Name Then Exit Sub
    For Each cell In c.Cells
        k = cell.Row & "|" & cell.Column
        On Error Resume Next
        jSeen.Add k, k
        If Err.Number = 0 Then
            On Error GoTo 0
            If cell.HasFormula Then
                jEntries.Add Array(cell.Row, cell.Column, True, cell.Formula)
            Else
                jEntries.Add Array(cell.Row, cell.Column, False, cell.Value)
            End If
        End If
        On Error GoTo 0
    Next
End Sub

' Stop recording. A command that changed something becomes the one Ctrl+Z (and the
' ribbon's Undo) takes back. Returns how many cells it changed.
Public Function EndEdit() As Long
    Dim e As Variant, kept As Collection, c As Range, ws As Worksheet
    If Not jOn Then Exit Function
    jOn = False
    Set ws = ParamTable.MainSheet
    Set kept = New Collection
    For Each e In jEntries
        Set c = ws.Cells(e(0), e(1))
        ' Only cells that really changed - a write the cell's rule refused put it back.
        If e(2) <> c.HasFormula Then
            kept.Add Array(e(0), e(1), e(2), e(3), c.HasFormula, IIf(c.HasFormula, c.Formula, c.Value))
        ElseIf e(2) Then
            If CStr(e(3)) <> c.Formula Then kept.Add Array(e(0), e(1), True, e(3), True, c.Formula)
        ElseIf Not ParamTable.SameValue(e(3), c.Value) Or VarType(e(3)) <> VarType(c.Value) Then
            kept.Add Array(e(0), e(1), False, e(3), False, c.Value)
        End If
    Next
    Set jEntries = Nothing
    Set jSeen = Nothing
    EndEdit = kept.Count
    If kept.Count = 0 Then Exit Function
    If undoStack Is Nothing Then Set undoStack = New Collection
    undoStack.Add Array(jLabel, kept)
    Do While undoStack.Count > MAX_UNDO
        undoStack.Remove 1
    Loop
    On Error Resume Next
    Application.OnUndo "Undo " & jLabel, "'" & ThisWorkbook.Name & "'!Panel.UndoCommand"
End Function

' What Ctrl+Z runs after a command (OnUndo needs a Sub).
Public Sub UndoCommand()
    UndoLast
End Sub

' What Undo would take back ("" when nothing).
Public Function UndoLabel() As String
    If undoStack Is Nothing Then Exit Function
    If undoStack.Count = 0 Then Exit Function
    UndoLabel = undoStack(undoStack.Count)(0)
End Function

' Put the last command's cells back. A cell changed since (typed over) is left alone and
' counted. Returns "restored skipped".
Public Function UndoLast() As String
    Dim top As Variant, e As Variant, c As Range, ws As Worksheet, ok As Long, skip As Long, same As Boolean, ev As Boolean
    If undoStack Is Nothing Then
        Application.StatusBar = TITLE & ":  nothing to undo."
        UndoLast = "0 0"
        Exit Function
    End If
    If undoStack.Count = 0 Then
        Application.StatusBar = TITLE & ":  nothing to undo."
        UndoLast = "0 0"
        Exit Function
    End If
    top = undoStack(undoStack.Count)
    undoStack.Remove undoStack.Count
    Set ws = ParamTable.MainSheet
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    For Each e In top(1)
        Set c = ws.Cells(e(0), e(1))
        If e(4) Then
            same = c.HasFormula And c.Formula = CStr(e(5))
        Else
            same = Not c.HasFormula And ParamTable.SameValue(c.Value, e(5))
        End If
        If same Then
            If e(2) Then c.Formula = e(3) Else c.Value = e(3)
            ok = ok + 1
        Else
            skip = skip + 1
        End If
    Next
Done:
    Application.EnableEvents = ev
    UndoLast = ok & " " & skip
    Application.StatusBar = TITLE & ":  undid " & top(0) & " - " & ok & " cell(s) put back" & _
                            IIf(skip > 0, ", " & skip & " left alone (changed since)", "") & "."
    RefreshWindows
End Function

' ============================================================ op rows

' The rows (op lines) of a selection: every visible data row with a selected cell, in
' sheet order, each once.
Public Function OpRows(ByVal r As Range) As Collection
    Dim cells As Range, area As Range, rr As Long, out As Collection, i As Long, rowsSeen() As Boolean, last As Long
    Set out = New Collection
    Set OpRows = out
    If r Is Nothing Then Exit Function
    Set cells = ParamTable.DataCells(r)
    If cells Is Nothing Then Exit Function
    last = ParamTable.LastRow
    ReDim rowsSeen(FIRST_ROW To Application.Max(FIRST_ROW, last))
    For Each area In cells.Areas
        For rr = area.Row To area.Row + area.Rows.Count - 1
            If rr >= FIRST_ROW And rr <= last Then rowsSeen(rr) = True
        Next
    Next
    For i = FIRST_ROW To last
        If rowsSeen(i) Then out.Add i
    Next
End Function

' A row by its op number, or 0.
Public Function RowOfOp(ByVal op As Variant) As Long
    Dim f As Range
    Set f = ParamTable.OpCell(op)
    If Not f Is Nothing Then RowOfOp = f.Row
End Function

' "op 7  -  T12 FINISH  -  OD finish" - one op line in words.
Public Function OpLabel(ByVal r As Long) As String
    Dim ws As Worksheet, s As String, t As Variant, cm As Variant
    Set ws = ParamTable.MainSheet
    s = "op " & ws.Cells(r, 1).Value
    t = ParamTable.ValueAt(r, "tool")
    If Not IsEmpty(t) Then s = s & "  -  T" & t
    s = s & " " & ws.Cells(r, 2).Value
    cm = ParamTable.ValueAt(r, "comment")
    If Trim$(CStr(cm)) <> "" Then s = s & "  -  " & cm
    OpLabel = s
End Function

' "op 7", "ops 7, 8, 12", "14 ops (7, 8, 12, 15, ...)" - which ops, in few words.
Public Function RowsText(ByVal rows As Collection) As String
    Dim i As Long, s As String
    If rows Is Nothing Then Exit Function
    If rows.Count = 0 Then RowsText = "no ops": Exit Function
    If rows.Count = 1 Then RowsText = "op " & ParamTable.MainSheet.Cells(rows(1), 1).Value: Exit Function
    For i = 1 To Application.Min(rows.Count, 6)
        s = s & IIf(s = "", "", ", ") & ParamTable.MainSheet.Cells(rows(i), 1).Value
    Next
    If rows.Count > 6 Then
        RowsText = rows.Count & " ops (" & s & ", ...)"
    Else
        RowsText = "ops " & s
    End If
End Function

' "7|8|12" - the same rows give the same key.
Public Function RowsKey(ByVal rows As Collection) As String
    Dim v As Variant
    If rows Is Nothing Then Exit Function
    For Each v In rows
        RowsKey = RowsKey & "|" & v
    Next
End Function

' ============================================================ focus: columns and groups

' The column groups of the sheet (row 1's names), in order.
Public Function GroupNames() As Collection
    Dim ws As Worksheet, c As Long, lc As Long, out As Collection, v As String
    Set out = New Collection
    Set ws = ParamTable.MainSheet
    lc = ParamTable.LastCol
    c = 1
    Do While c <= lc
        v = Trim$(CStr(ws.Cells(1, c).MergeArea.Cells(1, 1).Value))
        If v <> "" Then out.Add v
        c = c + ws.Cells(1, c).MergeArea.Columns.Count
    Loop
    Set GroupNames = out
End Function

' Open (unhide) the column group a column is in.
Private Sub OpenGroupOf(ByVal col As Long)
    Dim area As Range
    Set area = ParamTable.MainSheet.Cells(1, col).MergeArea
    If area.EntireColumn.Hidden Or AnyHidden(area) Then area.EntireColumn.Hidden = False
End Sub

Private Function AnyHidden(ByVal area As Range) As Boolean
    Dim c As Range
    For Each c In area.Columns
        If c.EntireColumn.Hidden Then AnyHidden = True: Exit Function
    Next
End Function

' Fold every outlined group except the one holding `keepCol` (0 = open them all).
Private Sub FoldOthers(ByVal keepCol As Long)
    Dim ws As Worksheet, c As Long, lc As Long, area As Range, keep As Range, k As Long
    Set ws = ParamTable.MainSheet
    lc = ParamTable.LastCol
    Dim keepFirst As Long
    If keepCol > 0 Then keepFirst = ws.Cells(1, keepCol).MergeArea.Column
    c = 1
    Do While c <= lc
        Set area = ws.Cells(1, c).MergeArea
        For k = 2 To area.Columns.Count
            ' Only the outlined detail (outline level 2); the handle column stays.
            If ws.Columns(area.Column + k - 1).OutlineLevel > 1 Then
                ws.Columns(area.Column + k - 1).Hidden = (keepCol > 0 And area.Column <> keepFirst)
            End If
        Next
        c = c + area.Columns.Count
    Loop
End Sub

' Show the named columns ("coolant|coolant_before|coolant_with") for these rows: open
' their groups, select those rows' cells there, scroll to them. True when any exists.
Public Function ShowColumns(ByVal names As String, Optional ByVal rows As Collection) As Boolean
    Dim ws As Worksheet, nm As Variant, c As Long, first As Long, sel As Range, r As Variant, cols As Collection
    Set ws = ParamTable.MainSheet
    Set cols = New Collection
    For Each nm In Split(names, "|")
        c = ParamTable.ColOf(CStr(nm))
        If c > 0 Then
            If ws.Columns(c).Hidden Then OpenGroupOf c
            If ws.Columns(c).Hidden Then ws.Columns(c).Hidden = False
            cols.Add c
            If first = 0 Or c < first Then first = c
        End If
    Next
    If cols.Count = 0 Then Exit Function
    ShowColumns = True
    If rows Is Nothing Then Set rows = OpRows(Selection)
    For Each r In rows
        For Each nm In cols
            If sel Is Nothing Then Set sel = ws.Cells(r, nm) Else Set sel = Union(sel, ws.Cells(r, nm))
        Next
    Next
    BringInView first, rows
    If Not sel Is Nothing Then
        Following = True
        On Error Resume Next
        If ActiveSheet.Name <> ws.Name Then ws.Activate
        sel.Select
        On Error GoTo 0
        Following = False
        ' A window moved the sheet: the other open windows follow it too, and a later click
        ' on these same ops is not taken for a move.
        TellWindows rows
    End If
End Function

' The open windows follow these rows (each ignores rows it already shows).
Private Sub TellWindows(ByVal rows As Collection)
    Dim f As Variant, k As String
    k = RowsKey(rows)
    If k = "" Or k = lastKey Then Exit Sub
    lastKey = k
    If wins Is Nothing Then Exit Sub
    On Error Resume Next
    For Each f In wins
        f.FollowRows rows
    Next
End Sub

' Scroll so a column (and the first of the rows) is in view.
Private Sub BringInView(ByVal col As Long, ByVal rows As Collection)
    Dim w As Window, ws As Worksheet, r As Long
    Set ws = ParamTable.MainSheet
    If ActiveSheet.Name <> ws.Name Then ws.Activate
    Set w = ActiveWindow
    On Error Resume Next
    If Intersect(w.VisibleRange, ws.Columns(col)) Is Nothing Or col > w.VisibleRange.Column + w.VisibleRange.Columns.Count - 3 Then
        w.ScrollColumn = col
    End If
    If Not rows Is Nothing Then
        If rows.Count > 0 Then
            r = rows(1)
            If Intersect(w.VisibleRange, ws.Rows(r)) Is Nothing Then w.ScrollRow = Application.Max(FIRST_ROW, r - 2)
        End If
    End If
End Sub

' Go to a column group: open it, fold the others, put the cursor in it on the same rows.
' "" (or "All columns") opens every group.
Public Sub GoToGroup(ByVal groupName As String)
    Dim ws As Worksheet, c As Long, lc As Long, area As Range, rows As Collection, sel As Range, r As Variant
    Set ws = ParamTable.MainSheet
    If ActiveSheet.Name <> ws.Name Then ws.Activate
    Set rows = OpRows(Selection)
    If groupName = "" Or groupName = "All columns" Then
        FoldOthers 0
        ActiveWindow.ScrollColumn = 1
        Exit Sub
    End If
    lc = ParamTable.LastCol
    c = 1
    Do While c <= lc
        Set area = ws.Cells(1, c).MergeArea
        If Trim$(CStr(area.Cells(1, 1).Value)) = groupName Then Exit Do
        c = c + area.Columns.Count
    Loop
    If c > lc Then Exit Sub
    FoldOthers c
    area.EntireColumn.Hidden = False
    If rows.Count = 0 Then rows.Add FIRST_ROW
    For Each r In rows
        If sel Is Nothing Then Set sel = ws.Cells(r, c) Else Set sel = Union(sel, ws.Cells(r, c))
    Next
    BringInView c, rows
    Following = True
    sel.Select
    Following = False
    ActiveWindow.ScrollColumn = c
End Sub

' ============================================================ find an op

' Find an op: "12" its op number, "T12" the first op of tool 12, anything else in its
' comment or type. From the row after the cursor, round to the top. True when found.
Public Function FindOp(ByVal text As String, Optional ByVal fromNext As Boolean = True) As Boolean
    Dim ws As Worksheet, r As Long, start As Long, last As Long, n As Long, t As String, hit As Boolean
    Dim cmt As Long, tl As Long
    text = Trim$(text)
    If text = "" Then Exit Function
    lastFind = text
    Set ws = ParamTable.MainSheet
    If ActiveSheet.Name <> ws.Name Then ws.Activate
    last = ParamTable.LastRow
    If last < FIRST_ROW Then Exit Function
    cmt = ParamTable.ColOf("comment"): tl = ParamTable.ColOf("tool")
    start = ActiveCell.Row
    If start < FIRST_ROW Or start > last Then start = FIRST_ROW - 1
    If Not fromNext Then start = start - 1
    t = LCase$(text)
    For n = 1 To last - FIRST_ROW + 1
        r = start + n
        If r > last Then r = r - (last - FIRST_ROW + 1)
        If Not ws.Rows(r).Hidden Then
            If IsNumeric(t) Then
                hit = (CStr(ws.Cells(r, 1).Value) = t)
            ElseIf Left$(t, 1) = "t" And IsNumeric(Mid$(t, 2)) And tl > 0 Then
                hit = (CStr(ws.Cells(r, tl).Value) = Mid$(t, 2))
            Else
                hit = InStr(1, LCase$(CStr(ws.Cells(r, 2).Value)), t) > 0
                If Not hit And cmt > 0 Then hit = InStr(1, LCase$(CStr(ws.Cells(r, cmt).Value)), t) > 0
            End If
            If hit Then
                If ws.Columns(ActiveCell.Column).Hidden Then ws.Cells(r, 1).Select Else ws.Cells(r, ActiveCell.Column).Select
                If Intersect(ActiveWindow.VisibleRange, ws.Rows(r)) Is Nothing Then ActiveWindow.ScrollRow = Application.Max(FIRST_ROW, r - 2)
                Application.StatusBar = TITLE & ":  " & OpLabel(r)
                FindOp = True
                Exit Function
            End If
        End If
    Next
    Application.StatusBar = TITLE & ":  no op matches '" & text & "' (an op number, T and a tool number, or words in the comment)."
End Function

Public Sub FindNext()
    If lastFind = "" Then
        Application.StatusBar = TITLE & ":  type an op number, T12 or words from a comment in Find op first."
    Else
        FindOp lastFind, True
    End If
End Sub

' ============================================================ windows that follow

' A window that follows the selection has opened (or come back to the front).
Public Sub Opened(ByVal f As Object)
    Dim i As Long
    If wins Is Nothing Then Set wins = New Collection
    For i = 1 To wins.Count
        If wins(i) Is f Then Exit Sub
    Next
    wins.Add f
End Sub

Public Sub Closed(ByVal f As Object)
    Dim i As Long
    If wins Is Nothing Then Exit Sub
    For i = wins.Count To 1 Step -1
        If wins(i) Is f Then wins.Remove i
    Next
End Sub

' From the sheet: the selection moved. Mark the row; open windows follow other ops.
Public Sub SelectionMoved(ByVal Target As Range)
    Dim rows As Collection, k As String, f As Variant
    On Error Resume Next
    MarkRow ActiveCell.Row
    If wins Is Nothing Then Exit Sub
    If wins.Count = 0 Then Exit Sub
    Set rows = OpRows(Target)
    If rows.Count = 0 Then Exit Sub
    k = RowsKey(rows)
    If Following Then lastKey = k: Exit Sub ' a window moved it: remember, so clicking back is a move
    If k = lastKey Then Exit Sub
    lastKey = k
    For Each f In wins
        f.FollowRows rows
    Next
End Sub

' After values changed under them (undo, another window's Apply): every open window
' reads its rows again.
Public Sub RefreshWindows()
    Dim f As Variant
    On Error Resume Next
    If wins Is Nothing Then Exit Sub
    For Each f In wins
        f.Refresh
    Next
End Sub

' ============================================================ the row mark

Private Sub MarkRow(ByVal r As Long)
    Dim nm As Name
    If ActiveSheet.Name <> ParamTable.MainSheet.Name Then Exit Sub
    If r < FIRST_ROW Or r > ParamTable.LastRow Then r = 0
    On Error Resume Next
    If Not marked Then AddRowMark
    marked = True
    Set nm = ThisWorkbook.Names("PT_Row")
    If nm Is Nothing Then
        ThisWorkbook.Names.Add Name:="PT_Row", RefersTo:="=" & r, Visible:=False
    ElseIf nm.RefersTo <> "=" & r Then
        nm.RefersTo = "=" & r
    End If
End Sub

' Once per workbook: the line above and below the current op row, a tint on its frozen
' columns. Lowest priority - the edit highlights still show through.
Private Sub AddRowMark()
    Dim ws As Worksheet, all As Range, fc As FormatCondition, i As Long, frozen As Long
    Set ws = ParamTable.MainSheet
    Set all = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(Application.Max(FIRST_ROW, ParamTable.LastRow), ParamTable.LastCol))
    On Error Resume Next
    For i = 1 To all.FormatConditions.Count
        If InStr(1, all.FormatConditions(i).Formula1, "PT_Row", vbTextCompare) > 0 Then Exit Sub
    Next
    On Error GoTo 0
    If Not NameExists("PT_Row") Then ThisWorkbook.Names.Add Name:="PT_Row", RefersTo:="=0", Visible:=False
    Set fc = all.FormatConditions.Add(Type:=xlExpression, Formula1:="=ROW()=PT_Row")
    fc.SetLastPriority
    fc.StopIfTrue = False
    With fc.Borders(xlTop): .LineStyle = xlContinuous: .Color = RGB(37, 99, 235): End With
    With fc.Borders(xlBottom): .LineStyle = xlContinuous: .Color = RGB(37, 99, 235): End With
    frozen = ActiveWindow.SplitColumn
    If frozen < 1 Then frozen = 4
    Set fc = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(Application.Max(FIRST_ROW, ParamTable.LastRow), frozen)) _
               .FormatConditions.Add(Type:=xlExpression, Formula1:="=ROW()=PT_Row")
    fc.SetLastPriority
    fc.StopIfTrue = False
    fc.Interior.Color = RGB(219, 234, 254)
End Sub

Private Function NameExists(ByVal n As String) As Boolean
    Dim x As Name
    On Error Resume Next
    Set x = ThisWorkbook.Names(n)
    NameExists = Not x Is Nothing
End Function

' ============================================================ help

' The user manual's folder: as the dump recorded it (PT_Help), else beside the add-in in
' Mastercam's own Add-Ins folder. "" when neither is there.
Public Function HelpFolder() As String
    Dim p As String, sh As Object
    On Error Resume Next
    p = Evaluate(ThisWorkbook.Names("PT_Help").RefersTo)
    If p <> "" Then
        If Dir(p & "\index.html") <> "" Then HelpFolder = p: Exit Function
    End If
    Set sh = CreateObject("WScript.Shell")
    p = sh.RegRead("HKCU\SOFTWARE\CNC Software\Mastercam 2026\UserDir")
    If p <> "" Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        p = p & "Add-Ins\ParamTable\help"
        If Dir(p & "\index.html") <> "" Then HelpFolder = p: Exit Function
    End If
End Function

' Open a page of the user manual ("coolant" -> help\coolant.html; "" -> the front page).
Public Sub ShowHelp(Optional ByVal topic As String = "")
    Dim d As String, p As String
    d = HelpFolder()
    If d = "" Then
        MsgBox "The user manual is not installed. It comes with the add-in, in the ParamTable\help folder " & _
               "beside it (Documents\My Mastercam 2026\Mastercam\Add-Ins).", vbInformation, TITLE
        Exit Sub
    End If
    p = d & "\" & IIf(topic = "", "index", topic) & ".html"
    If Dir(p) = "" Then p = d & "\index.html"
    CreateObject("Shell.Application").ShellExecute p
End Sub

' ============================================================ ribbon

' A command from the ribbon, with undo around it.
Private Sub Recorded(ByVal label As String, ByVal proc As String, Optional ByVal arg As Variant)
    Dim why As String
    BeginEdit label
    On Error GoTo Fail
    If IsMissing(arg) Then Application.Run proc Else Application.Run proc, arg
    GoTo Done
Fail:
    why = Err.Description
Done:
    On Error GoTo 0
    EndEdit
    RefreshWindows
    If why <> "" Then MsgBox "Something went wrong: " & why, vbExclamation, TITLE
End Sub

Public Sub RbUndo(control As IRibbonControl): UndoCommand: End Sub
Public Sub RbHelp(control As IRibbonControl): ShowHelp "": End Sub
Public Sub RbFindOp(control As IRibbonControl, text As String): FindOp text, True: End Sub
Public Sub RbFindNext(control As IRibbonControl): FindNext: End Sub

Public Sub RbSet(control As IRibbonControl): Recorded "Set selected", "ParamTable.SetSelected": End Sub
Public Sub RbScale(control As IRibbonControl): Recorded "Scale", "ParamTable.ScaleSelected": End Sub
Public Sub RbCopy(control As IRibbonControl): Recorded "Copy from op", "ParamTable.CopyFromOp": End Sub
Public Sub RbRevertSel(control As IRibbonControl): Recorded "Revert cells", "ParamTable.RevertSelected": End Sub
Public Sub RbRevertRows(control As IRibbonControl): Recorded "Revert rows", "ParamTable.RevertRows": End Sub
Public Sub RbRevertAll(control As IRibbonControl): Recorded "Revert everything", "ParamTable.RevertAll": End Sub
Public Sub RbTarget(control As IRibbonControl): Planner.ShowTargetWindow: End Sub
Public Sub RbScenLoad(control As IRibbonControl): Recorded "Restore scenario", "ParamTable.ScenarioWindow", "scenload": End Sub

' The op windows (each its own window module; these just open them).
Public Sub RbText(control As IRibbonControl): ParamTable.ShowTextWindow: End Sub
Public Sub RbCoolant(control As IRibbonControl): ParamTable.ShowCoolantWindow: End Sub
Public Sub RbSpeed(control As IRibbonControl): ParamTable.CalcSpeed: End Sub      ' modeless: its own writes are each one undo
Public Sub RbInspect(control As IRibbonControl): Planner.ShowInspectWindow: End Sub

' The Go to menu: every column group of this sheet, and All columns.
Public Sub RbGoToMenu(control As IRibbonControl, ByRef content)
    Dim s As String, g As Variant, i As Long
    s = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    s = s & "<button id=""ptGoAll"" label=""All columns (open every group)"" tag="""" onAction=""RbGoToItem""/>"
    s = s & "<menuSeparator id=""ptGoSep""/>"
    On Error Resume Next
    For Each g In GroupNames()
        i = i + 1
        s = s & "<button id=""ptGo" & i & """ label=""" & XmlText(CStr(g)) & """ tag=""" & XmlText(CStr(g)) & """ onAction=""RbGoToItem""/>"
    Next
    content = s & "</menu>"
End Sub

Public Sub RbGoToItem(control As IRibbonControl)
    GoToGroup control.Tag
End Sub

Private Function XmlText(ByVal s As String) As String
    XmlText = Replace(Replace(Replace(Replace(s, "&", "&amp;"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")
End Function

' ============================================================ for the checks

' For tools\check_macros.ps1 (no window): a range's ops in words, then its rows.
Public Function OpRowsSelfTest(ByVal r As Range) As String
    Dim rows As Collection
    Set rows = OpRows(r)
    OpRowsSelfTest = RowsText(rows) & "|" & RowsKey(rows)
End Function

' Show columns for a range's ops: whether any exists, and the selection it leaves.
Public Function ShowColumnsSelfTest(ByVal names As String, ByVal r As Range) As String
    Dim ok As Boolean
    ok = ShowColumns(names, OpRows(r))
    ShowColumnsSelfTest = ok & "|" & Selection.Address(False, False)
End Function
