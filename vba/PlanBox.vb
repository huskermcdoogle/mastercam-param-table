' Parameter Table Tool - the window for Hit a target time and Even out flips (a UserForm
' named PlanBox): which ops, the goal, a table of every op before -> after, and a summary,
' all worked out on the sheet's own live formulas before anything is written.
'
' Typing in the goal box marks the table out of date; Enter (or Preview) works it out,
' and once it is up to date the same button - Apply - writes it. Picking from a list or
' ticking a box works it out at once. Esc / Cancel writes nothing.
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblScope, cboScope,
' lblTarget, cboTarget, lblHow, cboHow, chk1, chk2, lst, lblSummary, btnOK, btnCancel -
' and puts this code behind it.
Option Explicit

Public Accepted As Boolean

Private Const PAD As Single = 8

Private mode As String                      ' "time" or "flips"
Private sel As Range
Private busy As Boolean
Private stale As Boolean
Private planOk As Boolean

Private Sub UserForm_Initialize()
    Accepted = False
    Me.Width = 640
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblScope.Font.Size = 9
    lblTarget.Font.Size = 9
    lblTarget.Font.Bold = True
    lblHow.Font.Size = 9
    cboScope.Style = fmStyleDropDownList
    cboHow.Style = fmStyleDropDownList
    cboTarget.Font.Size = 11
    cboTarget.MatchEntry = fmMatchEntryNone
    lst.Font.Name = "Consolas"
    lst.Font.Size = 9
    lst.ColumnCount = 6
    lst.TabStop = False
    lblSummary.Font.Size = 9
    lblSummary.WordWrap = True
    btnOK.Default = True                    ' Enter: preview, then apply
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                 ' Esc
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    cboTarget.SetFocus
    cboTarget.SelStart = 0
    cboTarget.SelLength = Len(cboTarget.Text)
End Sub

' The tool, the selection, the scope captions and which to start on.
Public Sub Setup(ByVal how As String, ByVal selection As Range, ByVal scopes As Variant, ByVal scopeIdx As Long)
    Dim i As Long, w As Single, y As Single, v As Variant
    busy = True
    mode = how
    Set sel = selection
    cboScope.Clear
    For i = LBound(scopes) To UBound(scopes)
        cboScope.AddItem CStr(scopes(i))
    Next
    cboScope.ListIndex = scopeIdx
    cboTarget.Clear
    cboHow.Clear
    lblScope.Caption = "Operations:"
    If mode = "time" Then
        Me.Caption = "Parameter Table - hit a target time"
        lblInfo.Caption = "Scales the feeds of these ops - all by one factor, each held inside its own limits - so their " & _
                          "estimated time (est_seconds) comes to the target. Worked out on the sheet's own formulas."
        lblTarget.Caption = "Target time for them (12:30, 1:02:00) - or a cut (10%):"
        For Each v In Array("5%", "10%", "15%", "20%", "-10%")
            cboTarget.AddItem v
        Next
        lblHow.Caption = ""
        lblHow.Visible = False
        cboHow.Visible = False
        chk1.Caption = "Speeds too (same factor - time goes with feed x speed)"
        chk2.Caption = "Main cutting feed only (leave plunge, retract, back feeds)"
        lst.ColumnWidths = "40 pt;36 pt;120 pt;80 pt;90 pt;250 pt"
    Else
        Me.Caption = "Parameter Table - even out insert flips"
        lblInfo.Caption = "Each tool's flips per part (flips_part, summed over all its ops - an edge carries on from op to op). " & _
                          "'Even out' keeps the count but uses every edge the same: the shortest edge time (or slowest feed) that still gives it."
        lblTarget.Caption = "Goal for each tool - or type a flip count:"
        For Each v In Planner.FlipTargets()
            cboTarget.AddItem v
        Next
        cboTarget.ListIndex = 0
        lblHow.Caption = "Change:"
        cboHow.AddItem "insp_time - the edge time between stops"
        cboHow.AddItem "feeds - keep insp_time, move the cut time"
        cboHow.ListIndex = 0
        chk1.Visible = False
        chk2.Visible = False
        lst.ColumnWidths = "40 pt;36 pt;110 pt;90 pt;120 pt;120 pt"
    End If

    w = Me.InsideWidth - 2 * PAD
    y = PAD
    lblInfo.Move PAD, y, w, 28: y = y + 32
    lblScope.Move PAD, y + 3, 70, 14
    cboScope.Move PAD + 74, y, w - 74, 20: y = y + 26
    lblTarget.Move PAD, y, w, 14: y = y + 16
    cboTarget.Move PAD, y, 220, 22
    If mode = "time" Then
        chk1.Move PAD + 232, y - 6, w - 232, 16
        chk2.Move PAD + 232, y + 10, w - 232, 16
    Else
        lblHow.Move PAD + 232, y + 4, 50, 14
        cboHow.Move PAD + 284, y, w - 284, 20
    End If
    y = y + 30
    lst.Move PAD, y, w, 170: y = y + 176
    lblSummary.Move PAD, y, w, 54: y = y + 58
    btnOK.Move Me.InsideWidth - 2 * 80 - PAD, y, 76, 24
    btnCancel.Move Me.InsideWidth - 80 - PAD + 4, y, 76, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    busy = False
    Run
End Sub

' For the checks: set every input at once, then work it out.
Public Sub SetInputs(ByVal scopeIdx As Long, ByVal target As String, ByVal howIdx As Long, ByVal opt1 As Boolean, ByVal opt2 As Boolean)
    busy = True
    cboScope.ListIndex = scopeIdx
    cboTarget.Text = target
    If cboHow.ListCount > 0 Then cboHow.ListIndex = howIdx
    chk1.Value = opt1
    chk2.Value = opt2
    busy = False
    Run
End Sub

Public Function Summary() As String
    Summary = Replace(lblSummary.Caption, vbCrLf, " / ")
End Function

Public Function CanAccept() As Boolean
    CanAccept = btnOK.Enabled And Not stale
End Function

' The table as text: rows joined by " // ", columns by ";".
Public Function ListRows() As String
    Dim i As Long, j As Long, s As String
    For i = 0 To lst.ListCount - 1
        If i > 0 Then s = s & " // "
        For j = 0 To lst.ColumnCount - 1
            s = s & IIf(j = 0, "", ";") & lst.List(i, j)
        Next
    Next
    ListRows = s
End Function

' Work the plan out and show it.
Private Sub Run()
    Dim p As String, d As Variant
    If busy Then Exit Sub
    Me.MousePointer = fmMousePointerHourGlass
    p = Planner.PlanRun(mode, sel, cboScope.ListIndex, cboTarget.Text, IIf(cboHow.ListIndex < 0, 0, cboHow.ListIndex), _
                        chk1.Value, chk2.Value)
    Me.MousePointer = fmMousePointerDefault
    planOk = (Left$(p, 1) = "1")
    lblSummary.Caption = Mid$(p, 3)
    lblSummary.ForeColor = IIf(planOk, RGB(33, 115, 70), RGB(91, 101, 115))
    lst.Clear
    d = Planner.PlanList()
    If IsArray(d) Then lst.List = d
    stale = False
    btnOK.Caption = "Apply"
    btnOK.Enabled = planOk
End Sub

' Typed, not worked out yet: Enter works it out.
Private Sub Dirty()
    If busy Then Exit Sub
    stale = True
    btnOK.Caption = "Preview"
    btnOK.Enabled = True
    lblSummary.Caption = "Press Enter to work out what this changes."
    lblSummary.ForeColor = RGB(91, 101, 115)
    lst.Clear
End Sub

Private Sub cboTarget_Change()
    If busy Then Exit Sub
    If cboTarget.ListIndex >= 0 Then Run Else Dirty
End Sub

Private Sub cboScope_Change()
    Run
End Sub

Private Sub cboHow_Change()
    Run
End Sub

Private Sub chk1_Click()
    Run
End Sub

Private Sub chk2_Click()
    Run
End Sub

Private Sub btnOK_Click()
    If stale Then
        Run
    ElseIf planOk Then
        Accepted = True
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

' The window's X is Cancel.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
