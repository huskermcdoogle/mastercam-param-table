' Parameter Table Tool - the Hit a target time window (a UserForm named PlanBox): which
' ops, the time wanted, a table of every op before -> after, and a summary - all worked
' out on the sheet's own live formulas before anything is written.
'
' It stays open beside the sheet (modeless) and follows the ops selected - until a target
' is typed: then it holds that plan, so clicking about the sheet to check it does not
' lose it (Revert lets it follow again). Enter, or Preview, works the plan out; Apply - a
' button of its own, never the Enter key - writes it, as one Ctrl+Z. Percent: -10% is 10%
' less time, +10% 10% more; the sign is always typed.
'
' tools\build_vba.ps1 creates it with these named controls - lblHead, btnHelp, lblScope,
' cboScope, lblTarget, cboTarget, lblHint, chkMain, chkSpeeds, lst, lblDetail, lblNote,
' lblSummary, lblResult, btnRevert, btnPreview, btnApply, btnClose - and puts this code
' behind it.
Option Explicit

Private Const PAD As Single = 8

Private rows As Collection                  ' the ops it is on
Private busy As Boolean                     ' filling the boxes in: their events wait
Private stale As Boolean                    ' a target typed, not worked out yet
Private planOk As Boolean
Private pc As Collection, pv As Collection  ' the plan Apply writes: cells, values
Private was As Collection                   ' those cells' values when it was worked out
Private times As Variant                    ' the ops' time before and with the plan (seconds)
Private refreshed As Boolean                ' Refresh ran (Apply asks every window to)
Private lstBox(3) As Single                 ' the list's place: left, top, width, height

Private Sub UserForm_Initialize()
    Dim w As Single, y As Single, v As Variant
    busy = True                             ' setting the boxes up is not a change
    Me.Caption = "Parameter Table - hit a target time"
    Me.Width = 660
    lblHead.Font.Size = 10
    lblHead.Font.Bold = True
    btnHelp.Caption = "?"
    btnHelp.Font.Bold = True
    btnHelp.TabStop = False
    btnHelp.ControlTipText = "How Hit a target time works (the user manual)"
    lblScope.Caption = "Ops:"
    lblScope.Font.Size = 9
    cboScope.Style = fmStyleDropDownList
    lblTarget.Caption = "Target:"
    lblTarget.Font.Size = 9
    lblTarget.Font.Bold = True
    cboTarget.Font.Size = 11
    cboTarget.MatchEntry = fmMatchEntryNone
    For Each v In Array("-5%", "-10%", "-15%", "-20%", "+10%")
        cboTarget.AddItem v
    Next
    lblHint.Caption = "the time you want (12:30) - or a change: -10% = 10% less time, +10% = 10% more"
    lblHint.Font.Size = 8
    lblHint.ForeColor = RGB(91, 101, 115)
    chkMain.Caption = "Main cutting feed only (leave plunge, retract and back feeds as they are)"
    chkMain.Value = True
    chkSpeeds.Caption = "Speeds too (by the same amount)"
    chkSpeeds.Value = False
    lst.Font.Name = "Consolas"
    lst.Font.Size = 9
    lst.ColumnCount = 7
    lst.ColumnWidths = "30 pt;30 pt;80 pt;56 pt;80 pt;326 pt;0 pt"     ' fits the width: no sideways scrolling
    lst.TabStop = False
    lst.IntegralHeight = False
    lblDetail.Font.Size = 8
    lblDetail.WordWrap = True
    lblDetail.ForeColor = RGB(91, 101, 115)
    lblNote.Font.Size = 9
    lblNote.WordWrap = True
    lblNote.ForeColor = RGB(180, 83, 9)
    lblNote.Caption = "* Speeds raised: the flips here count cut time only. A faster speed wears the edge sooner, " & _
                      "so the real flips may not drop - check the edge time."
    lblNote.Visible = False
    lblSummary.Font.Size = 9
    lblSummary.WordWrap = True
    lblResult.Font.Size = 9
    lblResult.WordWrap = True
    btnRevert.Caption = "Revert"
    btnRevert.ControlTipText = "Forget this plan and follow the ops selected on the sheet again"
    btnPreview.Caption = "Preview"
    btnPreview.Default = True               ' Enter works the plan out - and only that
    btnApply.Caption = "Apply"
    btnApply.Accelerator = "A"
    btnApply.Default = False                ' never the Enter key: a plan is read before it is written
    btnApply.Enabled = False
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc

    w = Me.InsideWidth - 2 * PAD
    y = PAD
    lblHead.Move PAD, y + 2, w - 30, 16
    btnHelp.Move Me.InsideWidth - PAD - 22, y, 22, 20: y = y + 26
    lblScope.Move PAD, y + 3, 46, 14
    cboScope.Move PAD + 50, y, w - 50, 20: y = y + 28
    lblTarget.Move PAD, y + 4, 46, 14
    cboTarget.Move PAD + 50, y, 110, 22
    lblHint.Move PAD + 168, y + 5, w - 168, 14: y = y + 28
    chkMain.Move PAD + 50, y, w - 50, 16: y = y + 18
    chkSpeeds.Move PAD + 50, y, w - 50, 16: y = y + 24
    lst.Move PAD, y, w, 160
    lstBox(0) = PAD: lstBox(1) = y: lstBox(2) = w: lstBox(3) = 160
    y = y + 164
    lblDetail.Move PAD, y, w, 22: y = y + 24
    lblNote.Move PAD, y, w, 26: y = y + 28
    lblSummary.Move PAD, y, w, 54: y = y + 56
    lblResult.Move PAD, y, w, 28: y = y + 32
    btnRevert.Move PAD, y, 76, 24
    btnPreview.Move Me.InsideWidth - 3 * 80 - PAD + 4, y, 76, 24
    btnApply.Move Me.InsideWidth - 2 * 80 - PAD + 4, y, 76, 24
    btnClose.Move Me.InsideWidth - 80 - PAD + 4, y, 76, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
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

' Which ops, in words, and the choices: those rows, their tool, every op.
Private Sub FillScope()
    Dim ch As Variant, i As Long, sel As Range
    Set sel = Planner.RowsRange(rows)
    ch = Planner.ScopeChoices(sel)
    cboScope.Clear
    For i = LBound(ch) To UBound(ch)
        cboScope.AddItem CStr(ch(i))
    Next
    cboScope.ListIndex = Planner.ScopeDefault(sel)
    lblHead.Caption = Heading()
End Sub

Public Function Heading() As String
    If rows Is Nothing Then
        Heading = "Select one or more ops"
    ElseIf rows.Count = 0 Then
        Heading = "Select one or more ops"
    ElseIf rows.Count = 1 Then
        Heading = Panel.OpLabel(rows(1))
    Else
        Heading = Panel.RowsText(rows)
    End If
End Function

' A target typed and not applied: the window keeps it rather than follow the sheet.
Public Function Holding() As Boolean
    Holding = (Trim$(cboTarget.Text) <> "")
End Function

' The sheet's selection moved to other ops (Panel.SelectionMoved).
Public Sub FollowRows(ByVal ops As Collection)
    If Holding() Then
        Say "Not applied yet - Apply, or Revert to follow the sheet again.", RGB(180, 83, 9)
        Exit Sub
    End If
    Set rows = ops
    busy = True
    FillScope
    busy = False
    Run
End Sub

' Values changed under it (undo, another window's Apply): work it out again.
Public Sub Refresh()
    refreshed = True
    If Not stale Then Run
End Sub

' For the checks: fill in every input as a person would - a typed target waits for Enter.
Public Sub SetInputs(ByVal scopeIdx As Long, ByVal target As String, ByVal mainOnly As Boolean, ByVal withSpeeds As Boolean)
    busy = True
    cboScope.ListIndex = scopeIdx
    chkMain.Value = mainOnly
    chkSpeeds.Value = withSpeeds
    cboTarget.Text = target
    busy = False
    If Trim$(target) = "" Then Run Else Dirty
End Sub

' For the checks: type a target only, every other box as the window opened.
Public Sub TypeTarget(ByVal target As String)
    cboTarget.Text = target
End Sub

' For the checks: the options ticked - "main", "speeds", "main+speeds" or "".
Public Function Options() As String
    If chkMain.Value Then Options = "main"
    If chkSpeeds.Value Then Options = Options & IIf(Options = "", "", "+") & "speeds"
End Function

' For the checks: "enter" (whichever button Enter presses), "preview", "apply", "revert".
Public Sub Press(ByVal what As String)
    Select Case LCase$(what)
    Case "enter"
        If btnPreview.Default Then btnPreview_Click
        If btnApply.Default Then btnApply_Click
    Case "preview": btnPreview_Click
    Case "apply": btnApply_Click
    Case "revert": btnRevert_Click
    End Select
End Sub

Public Function Summary() As String
    Summary = Replace(lblSummary.Caption, vbCrLf, " / ")
End Function

Public Function Result() As String
    Result = Replace(lblResult.Caption, vbCrLf, " / ")
End Function

Public Function NoteText() As String
    If lblNote.Visible Then NoteText = lblNote.Caption
End Function

Public Function CanApply() As Boolean
    CanApply = btnApply.Enabled And Not stale
End Function

' The table as text: rows joined by " // ", the columns shown by ";".
Public Function ListRows() As String
    Dim i As Long, j As Long, s As String
    For i = 0 To lst.ListCount - 1
        If i > 0 Then s = s & " // "
        For j = 0 To 5
            s = s & IIf(j = 0, "", ";") & lst.List(i, j)
        Next
    Next
    ListRows = s
End Function

' Work the plan out on the sheet's own formulas and show it.
Private Sub Run()
    Dim p As String, d As Variant, c As Variant
    If busy Then Exit Sub
    Me.MousePointer = fmMousePointerHourGlass
    p = Planner.PlanTargetTime(Planner.ScopeRange(Planner.RowsRange(rows), cboScope.ListIndex), cboTarget.Text, _
                               chkSpeeds.Value, chkMain.Value)
    Me.MousePointer = fmMousePointerDefault
    planOk = (Left$(p, 1) = "1")
    Planner.TakePlan pc, pv
    If pc Is Nothing Then planOk = False
    Set was = New Collection
    If planOk Then
        For Each c In pc
            was.Add c.Value
        Next
    End If
    lblSummary.Caption = Mid$(p, 3)
    lblSummary.ForeColor = IIf(planOk, RGB(33, 115, 70), RGB(91, 101, 115))
    lst.Clear
    d = Planner.PlanList()
    If IsArray(d) Then
        lst.List = d
        ' Raised speeds: the flips column is cut time only - its heading says so (the note).
        If chkSpeeds.Value Then lst.List(0, 3) = "flips / part *"
    End If
    lblNote.Visible = chkSpeeds.Value
    lblDetail.Caption = ""
    times = Planner.PlanTimes()
    stale = False
    btnApply.Enabled = planOk
End Sub

' Typed, not worked out yet: Enter (Preview) works it out. Apply waits for that.
Private Sub Dirty()
    If busy Then Exit Sub
    stale = True
    planOk = False
    btnApply.Enabled = False
    Set pc = Nothing: Set pv = Nothing
    lblSummary.Caption = "Press Enter (or Preview) to see what this changes, op by op."
    lblSummary.ForeColor = RGB(91, 101, 115)
    lst.Clear
    lblDetail.Caption = ""
End Sub

Private Sub Say(ByVal s As String, ByVal color As Long)
    lblResult.Caption = s
    lblResult.ForeColor = color
End Sub

Private Sub cboTarget_Change()
    If busy Then Exit Sub
    Say "", 0
    If cboTarget.ListIndex >= 0 Then Run Else Dirty
End Sub

Private Sub cboScope_Change()
    If busy Then Exit Sub
    Run
End Sub

Private Sub chkMain_Click()
    If busy Then Exit Sub
    Run
End Sub

Private Sub chkSpeeds_Click()
    If busy Then Exit Sub
    Run
End Sub

' The changes of the op picked, in full (the list cuts them to fit).
Private Sub lst_Click()
    On Error Resume Next
    If lst.ListIndex > 0 Then lblDetail.Caption = "op " & lst.List(lst.ListIndex, 0) & ":  " & lst.List(lst.ListIndex, 6)
End Sub

Private Sub btnPreview_Click()
    Run
End Sub

' Write the plan - each value through its cell's rule - as one Ctrl+Z.
Private Sub btnApply_Click()
    Dim i As Long, ok As Long, bad As Long, why As String, t As Variant
    If stale Or Not planOk Or pc Is Nothing Then Exit Sub
    For i = 1 To pc.Count
        If Not ParamTable.SameValue(pc(i).Value, was(i)) Then
            ' Typed over or undone since the preview: work it out again before anything goes in.
            Run
            Say "The sheet changed since this was worked out - here it is again. Check it, then Apply.", RGB(180, 83, 9)
            Exit Sub
        End If
    Next
    t = times                               ' the time before; read again after, the time now
    Panel.BeginEdit "Hit a target time"
    On Error GoTo Fail
    For i = 1 To pc.Count
        If ParamTable.TryWrite(pc(i), pv(i)) Then ok = ok + 1 Else bad = bad + 1
    Next
    GoTo Done
Fail:
    why = Err.Description
Done:
    On Error GoTo 0
    Panel.EndEdit
    ' Done with this target: the window shows the sheet as it is now, and follows again.
    Set pc = Nothing: Set pv = Nothing
    busy = True
    cboTarget.Text = ""
    busy = False
    refreshed = False
    Panel.RefreshWindows                    ' every open window reads the sheet again - this one too
    If Not refreshed Then Run
    If why <> "" Then
        Say "Stopped part way: " & why & " - " & ok & " cell(s) were written. Ctrl+Z puts them back.", RGB(185, 28, 28)
    Else
        Say "Done: " & ok & " cell(s) changed" & IIf(bad > 0, ", " & bad & " refused by their cell's limits", "") & _
            " - the ops now take " & Planner.TimeText(times(0)) & " (were " & Planner.TimeText(t(0)) & "). Ctrl+Z to undo.", RGB(33, 115, 70)
    End If
End Sub

' Forget the plan and the options; follow the ops selected on the sheet again.
Private Sub btnRevert_Click()
    busy = True
    cboTarget.Text = ""
    chkMain.Value = True
    chkSpeeds.Value = False
    busy = False
    Say "", 0
    stale = False
    Set rows = Planner.WindowRows()
    busy = True
    FillScope
    busy = False
    Run
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp "target-time"
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
