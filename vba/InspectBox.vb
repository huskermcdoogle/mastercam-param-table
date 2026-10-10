' Parameter Table Tool - the Inspection & inserts window (a UserForm named InspectBox):
' a line per tool of the ops selected (or every tool) - its insert, flips per part, how
' long each edge cuts now and how much of the last edge a part uses - and the edge time
' that uses every edge evenly for the same flips (or one flip fewer / more). A tool whose
' flips come from the toolpath itself (stops every so much cut, at the end of an op) says
' so instead: those change only in Mastercam.
'
' It stays open beside the sheet (modeless) and follows the ops selected - until the goal,
' the tools or the ticks are changed here; then it keeps them (Revert lets it follow
' again). Apply writes the edge time (insp_time) to EVERY op of each ticked tool that
' times its stops - flips are counted over the whole tool - as one Ctrl+Z.
'
' tools\build_vba.ps1 creates it with these named controls - lblHead, btnHelp, lblScope,
' cboScope, lblGoal, cboGoal, lblC1 .. lblC8 (the list's headings), lst, lblDetail,
' lblSummary, lblResult, btnRevert, btnApply, btnClose - and puts this code behind it.
Option Explicit

Private Const PAD As Single = 8
Private Const WIDTHS As String = "34;84;50;52;50;66;48;380"     ' the columns shown, in points
Private Const HINT As String = "Edge time: how long each edge cuts between flips (insp_time).  Last edge: how much of " & _
                               "that the last edge in a part gets to cut.  Click a tool for its note."

Private rows As Collection                  ' the ops it is on
Private busy As Boolean                     ' filling the boxes in: their events wait
Private touched As Boolean                  ' the goal, the tools or the ticks were changed here
Private refreshed As Boolean                ' Refresh ran (Apply asks every window to)
Private lastWhy As String                   ' why there is nothing to apply, when there is not
Private lstBox(3) As Single                 ' the list's place: left, top, width, height

Private Sub UserForm_Initialize()
    Dim w As Single, y As Single, x As Single, i As Long, cw() As String, s As String, v As Variant
    Dim heads As Variant, tips As Variant
    busy = True
    Me.Caption = "Parameter Table - inspection and inserts"
    Me.Width = 820
    lblHead.Font.Size = 10
    lblHead.Font.Bold = True
    btnHelp.Caption = "?"
    btnHelp.Font.Bold = True
    btnHelp.TabStop = False
    btnHelp.ControlTipText = "How Inspection & inserts works (the user manual)"
    lblScope.Caption = "Tools:"
    lblScope.Font.Size = 9
    cboScope.Style = fmStyleDropDownList
    lblGoal.Caption = "Goal:"
    lblGoal.Font.Size = 9
    lblGoal.Font.Bold = True
    cboGoal.Style = fmStyleDropDownList
    For Each v In Planner.InspectGoals()
        cboGoal.AddItem v
    Next
    cboGoal.ListIndex = 0

    ' The list: a tick box per tool, the numbers, a note - then hidden columns the window
    ' reads (whether it can change, the tool, flips now and new, a longer edge).
    cw = Split(WIDTHS, ";")
    For i = 0 To UBound(cw)
        s = s & IIf(s = "", "", ";") & cw(i) & " pt"
    Next
    lst.ColumnCount = 13
    lst.ColumnWidths = s & ";0 pt;0 pt;0 pt;0 pt;0 pt"
    lst.MultiSelect = fmMultiSelectMulti
    lst.ListStyle = fmListStyleOption
    lst.Font.Name = "Segoe UI"
    lst.Font.Size = 9
    lst.IntegralHeight = False
    heads = Array("Tool", "Insert", "Flips / part", "Edge time", "Last edge", "New edge time", "New flips", "Note")
    tips = Array("The tool number", "The insert, from the Tools page", "Insert flips per part now, over all of the tool's ops", _
                 "How long each edge cuts between flips now (insp_time; from the toolpath when it fixes the flips)", _
                 "How much of its edge time the last edge in a part gets to cut", "The edge time for the goal", _
                 "Flips per part with the new edge time", "What to know about this tool")
    lblDetail.Font.Size = 8
    lblDetail.WordWrap = True
    lblDetail.ForeColor = RGB(91, 101, 115)
    lblSummary.Font.Size = 9
    lblSummary.WordWrap = True
    lblResult.Font.Size = 9
    lblResult.WordWrap = True
    btnRevert.Caption = "Revert"
    btnRevert.ControlTipText = "Back to the same flips and the ops selected on the sheet"
    btnApply.Caption = "Apply"
    btnApply.Accelerator = "A"
    btnApply.Enabled = False                ' no Enter key: a click, once the list is read
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc

    w = Me.InsideWidth - 2 * PAD
    y = PAD
    lblHead.Move PAD, y + 2, w - 30, 16
    btnHelp.Move Me.InsideWidth - PAD - 22, y, 22, 20: y = y + 26
    lblScope.Move PAD, y + 3, 40, 14
    cboScope.Move PAD + 44, y, 420, 20: y = y + 26
    lblGoal.Move PAD, y + 3, 40, 14
    cboGoal.Move PAD + 44, y, 260, 20: y = y + 28
    x = PAD + 18                            ' past the tick boxes
    For i = 0 To 7
        With Me.Controls("lblC" & (i + 1))
            .Caption = heads(i)
            .ControlTipText = tips(i)
            .Font.Size = 8
            .Font.Bold = True
            .ForeColor = RGB(55, 65, 81)
            .Move x, y, CSng(cw(i)), 12
        End With
        x = x + CSng(cw(i))
    Next
    y = y + 14
    lst.Move PAD, y, w, 150
    lstBox(0) = PAD: lstBox(1) = y: lstBox(2) = w: lstBox(3) = 150
    y = y + 154
    lblDetail.Move PAD, y, w, 24: y = y + 26
    lblSummary.Move PAD, y, w, 40: y = y + 42
    lblResult.Move PAD, y, w, 28: y = y + 32
    btnRevert.Move PAD, y, 76, 24
    btnApply.Move Me.InsideWidth - 2 * 80 - PAD + 4, y, 76, 24
    btnClose.Move Me.InsideWidth - 80 - PAD + 4, y, 76, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    lblDetail.Caption = HINT
    busy = False
End Sub

' Shown on a scaled display, a list box drops back to its design size: put it back.
Private Sub UserForm_Activate()
    On Error Resume Next
    If lstBox(2) > 0 Then lst.Move lstBox(0), lstBox(1), lstBox(2), lstBox(3)
End Sub

' The ops it starts on (Planner.WindowRows).
Public Sub Setup(ByVal ops As Collection)
    Set rows = ops
    busy = True
    FillScope
    busy = False
    Run
End Sub

' The two choices of tools, in words; the heading.
Private Sub FillScope()
    Dim ch As Variant, i As Long, keep As Long
    keep = cboScope.ListIndex
    ch = Planner.InspectScopes(rows)
    cboScope.Clear
    For i = LBound(ch) To UBound(ch)
        cboScope.AddItem CStr(ch(i))
    Next
    cboScope.ListIndex = IIf(keep = 1, 1, 0)
    lblHead.Caption = Heading()
End Sub

Public Function Heading() As String
    If cboScope.ListIndex = 1 Then
        Heading = "Every tool on the sheet"
    ElseIf rows Is Nothing Then
        Heading = "Select one or more ops"
    ElseIf rows.Count = 0 Then
        Heading = "Select one or more ops"
    ElseIf rows.Count = 1 Then
        Heading = Panel.OpLabel(rows(1))
    Else
        Heading = Panel.RowsText(rows)
    End If
End Function

' Changed here and not applied: the window keeps it rather than follow the sheet.
Public Function Holding() As Boolean
    Holding = touched
End Function

' The sheet's selection moved to other ops (Panel.SelectionMoved).
Public Sub FollowRows(ByVal ops As Collection)
    Dim sameTools As Boolean
    If touched Then
        Say "Not applied yet - Apply, or Revert to follow the sheet again.", RGB(180, 83, 9)
        Exit Sub
    End If
    ' Another op of the same tools: the list stays as it is (no need to work it out again).
    sameTools = (cboScope.ListIndex = 0 And Planner.ToolsKey(ops) = Planner.ToolsKey(rows) And lst.ListCount > 0)
    Set rows = ops
    busy = True
    FillScope
    busy = False
    If Not sameTools Then Run
End Sub

' Values changed under it (undo, another window's Apply): work it out again, keeping the
' ticks made here.
Public Sub Refresh()
    refreshed = True
    Run touched
End Sub

' For the checks: the tools (0 those of the ops, 1 every tool), the goal, tools to untick.
Public Sub SetInputs(ByVal scopeIdx As Long, ByVal goalIdx As Long, ByVal untick As String)
    Dim i As Long
    busy = True
    cboScope.ListIndex = scopeIdx
    cboGoal.ListIndex = goalIdx
    lblHead.Caption = Heading()
    busy = False
    touched = (scopeIdx <> 0 Or goalIdx <> 0)
    Run
    If untick <> "" Then
        busy = True
        For i = 0 To lst.ListCount - 1
            If InStr("|" & untick & "|", "|" & lst.List(i, 9) & "|") > 0 Then lst.Selected(i) = False
        Next
        busy = False
        touched = True
        Totals
    End If
End Sub

' For the checks: "apply", "revert", "enter" (no button takes Enter here).
Public Sub Press(ByVal what As String)
    Select Case LCase$(what)
    Case "apply": btnApply_Click
    Case "revert": btnRevert_Click
    Case "enter"
        If btnApply.Default Then btnApply_Click
    End Select
End Sub

Public Function Summary() As String
    Summary = Replace(lblSummary.Caption, vbCrLf, " / ")
End Function

Public Function Result() As String
    Result = Replace(lblResult.Caption, vbCrLf, " / ")
End Function

Public Function CanApply() As Boolean
    CanApply = btnApply.Enabled
End Function

' The list as text: rows joined by " // ", the columns shown by ";", then "ticked" or "".
Public Function ListRows() As String
    Dim i As Long, j As Long, s As String
    For i = 0 To lst.ListCount - 1
        If i > 0 Then s = s & " // "
        For j = 0 To 7
            s = s & IIf(j = 0, "", ";") & lst.List(i, j)
        Next
        s = s & ";" & IIf(lst.Selected(i), "ticked", "")
    Next
    ListRows = s
End Function

' The tools ticked: "1|12".
Private Function TickedTools() As String
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then TickedTools = TickedTools & IIf(TickedTools = "", "", "|") & lst.List(i, 9)
    Next
End Function

' Work it out on the sheet's own flips and show a line per tool. Ticked to start with:
' every tool Apply would change (or, keeping ticks, those still ticked).
Private Sub Run(Optional ByVal keepTicks As Boolean = False)
    Dim p As String, d As Variant, keep As String, i As Long
    If busy Then Exit Sub
    If keepTicks Then keep = TickedTools()
    Me.MousePointer = fmMousePointerHourGlass
    p = Planner.PlanInspect(rows, cboScope.ListIndex = 1, IIf(cboGoal.ListIndex < 0, 0, cboGoal.ListIndex))
    Me.MousePointer = fmMousePointerDefault
    busy = True
    lst.Clear
    d = Planner.InspectList()
    If IsArray(d) Then lst.List = d
    For i = 0 To lst.ListCount - 1
        If keepTicks Then
            lst.Selected(i) = (lst.List(i, 8) = "1") And InStr("|" & keep & "|", "|" & lst.List(i, 9) & "|") > 0
        Else
            lst.Selected(i) = (lst.List(i, 8) = "1")
        End If
    Next
    busy = False
    lastWhy = IIf(Left$(p, 1) = "0", Mid$(p, 3), "")
    lblDetail.Caption = HINT
    Totals
End Sub

' What Apply would do, from the ticks.
Private Sub Totals()
    Dim i As Long, n As Long, f0 As Double, f1 As Double, warn As Long, tl As String, tot As Double
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            n = n + 1
            f0 = f0 + Val(lst.List(i, 10))
            f1 = f1 + Val(lst.List(i, 11))
            If lst.List(i, 12) = "1" Then warn = warn + 1
            tl = tl & IIf(tl = "", "", ", ") & lst.List(i, 0)
        End If
    Next
    If n = 0 Then
        lblSummary.Caption = IIf(lastWhy <> "", lastWhy, "Tick the tools to change.")
        lblSummary.ForeColor = RGB(91, 101, 115)
    Else
        tot = Planner.InspectTotal()
        lblSummary.Caption = "Apply sets the new edge time on every op of " & tl & " that times its stops:  flips per part " & _
                             Planner.NumText(f0) & " -> " & Planner.NumText(f1) & "  (the whole part " & Planner.NumText(tot) & _
                             " -> " & Planner.NumText(tot - f0 + f1) & ")." & _
                             IIf(warn > 0, vbCrLf & warn & " of them would run each edge longer than now - check the insert can take it.", "")
        lblSummary.ForeColor = IIf(warn > 0, RGB(180, 83, 9), RGB(33, 115, 70))
    End If
    btnApply.Enabled = (n > 0)
End Sub

' The note of the tool the cursor is on, in full.
Private Sub ShowDetail()
    Dim i As Long
    i = lst.ListIndex
    If i < 0 Or i >= lst.ListCount Then Exit Sub
    lblDetail.Caption = lst.List(i, 0) & IIf(lst.List(i, 1) <> "", "  (" & lst.List(i, 1) & ")", "") & ":  " & lst.List(i, 7)
End Sub

Private Sub Say(ByVal s As String, ByVal color As Long)
    lblResult.Caption = s
    lblResult.ForeColor = color
End Sub

' A tick: only a tool Apply would change can be ticked.
Private Sub lst_Change()
    Dim i As Long
    If busy Then Exit Sub
    busy = True
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) And lst.List(i, 8) <> "1" Then lst.Selected(i) = False
    Next
    busy = False
    touched = True
    ShowDetail
    Totals
End Sub

Private Sub lst_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ShowDetail
End Sub

Private Sub lst_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ShowDetail
End Sub

Private Sub cboGoal_Change()
    If busy Then Exit Sub
    touched = True
    Say "", 0
    Run
End Sub

Private Sub cboScope_Change()
    If busy Then Exit Sub
    touched = True
    lblHead.Caption = Heading()
    Say "", 0
    Run
End Sub

' Write the edge times - each through its cell's rule - as one Ctrl+Z.
Private Sub btnApply_Click()
    Dim r As String, p() As String, before As Double, why As String, ticked As String
    ticked = TickedTools()
    If ticked = "" Then Exit Sub
    If Planner.InspectStale() Then
        ' Typed over or undone since: what the list says is not what Apply would do.
        Run True
        Say "The sheet changed since this was worked out - here it is again. Check it, then Apply.", RGB(180, 83, 9)
        Exit Sub
    End If
    before = Planner.InspectTotal()
    Panel.BeginEdit "Inspection & inserts"
    On Error GoTo Fail
    r = Planner.InspectApply(ticked)
    GoTo Done
Fail:
    why = Err.Description
Done:
    On Error GoTo 0
    Panel.EndEdit
    ' Done: back to "the same flips" (a goal of one fewer would offer one fewer again),
    ' and the window follows the sheet again.
    touched = False
    busy = True
    cboGoal.ListIndex = 0
    busy = False
    refreshed = False
    Panel.RefreshWindows                    ' every open window reads the sheet again - this one too
    If Not refreshed Then Run
    p = Split(r & "|||", "|")
    If why <> "" Then
        Say "Stopped part way: " & why & ". Ctrl+Z puts back what was written.", RGB(185, 28, 28)
    Else
        Say "Done: edge time set on op" & IIf(InStr(p(2), ",") > 0, "s ", " ") & p(2) & " (" & p(3) & ")" & _
            IIf(Val(p(1)) > 0, ", " & p(1) & " refused by the cell's limits", "") & " - flips per part " & _
            Planner.NumText(before) & " -> " & Planner.NumText(Planner.InspectTotal()) & ". Ctrl+Z to undo.", RGB(33, 115, 70)
    End If
End Sub

' Back to the same flips and the ops selected on the sheet.
Private Sub btnRevert_Click()
    touched = False
    busy = True
    cboGoal.ListIndex = 0
    cboScope.ListIndex = 0
    busy = False
    Say "", 0
    Set rows = Planner.WindowRows()
    busy = True
    FillScope
    busy = False
    Run
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp "inspection"
End Sub

Private Sub btnClose_Click()
    Panel.Closed Me
    Planner.WindowGone Me
    Unload Me
End Sub

' The window's X closes it the same way.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Panel.Closed Me
        Planner.WindowGone Me
    End If
End Sub
