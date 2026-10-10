' Parameter Table Tool - the window for Set, Scale, Copy from op and the scenarios (a
' UserForm named EditBox): one box to type or pick in, a live line saying what OK would do
' - how many cells change and how many are refused, and why - and a list of each cell,
' old -> new, before anything is written. It works on the cells selected, so it waits for
' OK or Cancel (modal).
'
' Copy from op: the op to copy FROM (the one the cursor is on, to start with) and where
' INTO - the other selected rows, or every other op of that op's tool.
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblHint, btnHelp,
' lblPrompt, cbo, lblInto, cboInto, lblPreview, lstDetail, btnOK, btnCancel - and puts
' this code behind it.
Option Explicit

Public Accepted As Boolean

Private Const PAD As Single = 8

Private mode As String                      ' "set", "scale", "copy", "scensave", "scenload", "scendel"
Private target As Range
Private keys As Variant                     ' what each list row stands for (copy: op_idn)
Private listOnly As Boolean
Private topic As String                     ' its page of the user manual
Private busy As Boolean                     ' filling the boxes in: their events wait
Private intoPicked As Boolean               ' where a copy goes was picked here: keep it

Private lstBox(3) As Single                 ' the list's place: left, top, width, height

Private Sub UserForm_Initialize()
    Accepted = False
    Me.Width = 520
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblHint.Font.Size = 8
    lblHint.WordWrap = True
    lblHint.ForeColor = RGB(91, 101, 115)
    btnHelp.Caption = "?"
    btnHelp.Font.Bold = True
    btnHelp.TabStop = False
    btnHelp.ControlTipText = "How this works (the user manual)"
    lblPrompt.Font.Size = 9
    lblPrompt.Font.Bold = True
    cbo.Font.Size = 11
    cbo.MatchEntry = fmMatchEntryNone
    lblInto.Caption = "Into:"
    lblInto.Font.Size = 9
    lblInto.Font.Bold = True
    cboInto.Style = fmStyleDropDownList
    cboInto.Font.Size = 9
    lblPreview.Font.Size = 9
    lblPreview.WordWrap = True
    lstDetail.Font.Name = "Consolas"
    lstDetail.Font.Size = 9
    lstDetail.TabStop = False
    lstDetail.IntegralHeight = False
    btnOK.Caption = "OK"
    btnOK.Default = True                    ' Enter
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                 ' Esc
End Sub

' Typing replaces what is in the box: it opens with all of it selected.
' Shown on a scaled display, a list box drops back to its design size: put it back.
Private Sub UserForm_Activate()
    On Error Resume Next
    If lstBox(2) > 0 Then lstDetail.Move lstBox(0), lstBox(1), lstBox(2), lstBox(3)
    cbo.SetFocus
    cbo.SelStart = 0
    cbo.SelLength = Len(cbo.Text)
End Sub

' mode, the cells, a line about them in words, the columns as a small hint, what to type,
' the list (may be empty), what each list row stands for (or Empty: the row's text),
' whether only the list is allowed, what to start with (a list-only window: the row whose
' key or text it is), and its page of the user manual.
Public Sub Setup(ByVal how As String, ByVal cells As Range, ByVal caption As String, ByVal info As String, _
                 ByVal hint As String, ByVal prompt As String, ByVal choices As Variant, ByVal rowKeys As Variant, _
                 ByVal onlyList As Boolean, ByVal start As String, ByVal helpTopic As String)
    Dim i As Long, w As Single, y As Single
    busy = True
    mode = how
    Set target = cells
    keys = rowKeys
    listOnly = onlyList
    topic = helpTopic
    Me.Caption = caption
    lblInfo.Caption = info
    lblHint.Caption = hint
    lblPrompt.Caption = prompt
    cbo.Clear
    If IsArray(choices) Then
        For i = LBound(choices) To UBound(choices)
            cbo.AddItem CStr(choices(i))
        Next
    End If
    cbo.Style = IIf(onlyList, fmStyleDropDownList, fmStyleDropDownCombo)
    cbo.ShowDropButtonWhen = IIf(cbo.ListCount > 0, fmShowDropButtonWhenAlways, fmShowDropButtonWhenNever)
    If onlyList Then
        If start <> "" Then
            For i = 0 To cbo.ListCount - 1
                If KeyOf(i) = start Then cbo.ListIndex = i: Exit For
            Next
        End If
    Else
        cbo.Text = start
    End If

    w = Me.InsideWidth - 2 * PAD
    y = PAD
    lblInfo.Move PAD, y, w - 30, 28
    btnHelp.Move Me.InsideWidth - PAD - 22, y, 22, 20: y = y + 30
    lblHint.Visible = (hint <> "")
    If hint <> "" Then lblHint.Move PAD, y, w, 12: y = y + 16
    lblPrompt.Move PAD, y, w, 14: y = y + 16
    cbo.Move PAD, y, w, 22: y = y + 30
    lblInto.Visible = (mode = "copy")
    cboInto.Visible = (mode = "copy")
    If mode = "copy" Then
        lblInto.Move PAD, y + 3, 34, 14
        cboInto.Move PAD + 38, y, w - 38, 20: y = y + 28
        FillInto -1
    End If
    lblPreview.Move PAD, y, w, 42: y = y + 46
    lstDetail.Move PAD, y, w, 120
    lstBox(0) = PAD: lstBox(1) = y: lstBox(2) = w: lstBox(3) = 120
    y = y + 128
    btnOK.Move Me.InsideWidth - 2 * 72 - PAD, y, 68, 24
    btnCancel.Move Me.InsideWidth - 72 - PAD + 4, y, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    busy = False
    Refresh
End Sub

' Copy: where it can go, for the op copied from - keeping the choice made (-1: the
' default, the other selected rows when there are some, else the tool's other ops).
Private Sub FillInto(ByVal keep As Long)
    Dim ch As Variant, i As Long, was As Boolean
    was = busy
    busy = True
    ch = ParamTable.CopyIntoChoices(target, Value())
    cboInto.Clear
    For i = LBound(ch) To UBound(ch)
        cboInto.AddItem CStr(ch(i))
    Next
    If keep < 0 Then keep = ParamTable.CopyIntoDefault(target, Value())
    cboInto.ListIndex = keep
    busy = was
End Sub

Private Function KeyOf(ByVal i As Long) As String
    If IsArray(keys) Then KeyOf = CStr(keys(LBound(keys) + i)) Else KeyOf = cbo.List(i)
End Function

' What OK applies: the typed text, or the key of the picked row.
Public Function Value() As String
    If IsArray(keys) And cbo.ListIndex >= 0 Then
        Value = CStr(keys(LBound(keys) + cbo.ListIndex))
    Else
        Value = cbo.Text
    End If
End Function

' Copy: 0 into the selected rows, 1 into every other op of the tool.
Public Function IntoIdx() As Long
    If mode = "copy" Then IntoIdx = Application.Max(0, cboInto.ListIndex)
End Function

Public Function Preview() As String
    Preview = lblPreview.Caption
End Function

Public Function CanAccept() As Boolean
    CanAccept = btnOK.Enabled
End Function

Public Function DetailCount() As Long
    DetailCount = lstDetail.ListCount
End Function

Public Function DetailLine(ByVal i As Long) As String
    If i < lstDetail.ListCount Then DetailLine = lstDetail.List(i)
End Function

' For the checks: type into the box, or pick a list row.
Public Sub TypeIn(ByVal s As String)
    Dim i As Long
    If listOnly Then                        ' a pick-only list takes no typing: pick the matching row
        cbo.ListIndex = -1
        For i = 0 To cbo.ListCount - 1
            If cbo.List(i) = s Or KeyOf(i) = s Then cbo.ListIndex = i
        Next
    Else
        cbo.Text = s
    End If
    Refresh
End Sub

Public Sub Pick(ByVal i As Long)
    cbo.ListIndex = i
    Refresh
End Sub

Public Sub PickInto(ByVal i As Long)
    If mode <> "copy" Then Exit Sub
    intoPicked = True
    busy = True
    cboInto.ListIndex = i
    busy = False
    Refresh
End Sub

Private Sub Refresh()
    Dim p As String, d As Variant, i As Long
    If busy Then Exit Sub
    p = ParamTable.EditPreview(mode, target, Value(), IntoIdx())
    btnOK.Enabled = (Left$(p, 1) = "1")
    lblPreview.Caption = Mid$(p, 3)
    lblPreview.ForeColor = IIf(btnOK.Enabled, RGB(33, 115, 70), RGB(91, 101, 115))
    lstDetail.Clear
    d = ParamTable.EditDetail(mode, target, Value(), IntoIdx())
    If IsArray(d) Then
        For i = LBound(d) To UBound(d)
            lstDetail.AddItem CStr(d(i))
        Next
    End If
End Sub

Private Sub cbo_Change()
    If busy Then Exit Sub
    ' Another op copied from: another tool to offer - and, unless the place was picked
    ' here, the place that suits it.
    If mode = "copy" Then FillInto IIf(intoPicked, cboInto.ListIndex, -1)
    Refresh
End Sub

Private Sub cboInto_Change()
    If busy Then Exit Sub
    intoPicked = True
    Refresh
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp topic
End Sub

Private Sub btnOK_Click()
    If Not btnOK.Enabled Then Exit Sub
    Accepted = True
    Me.Hide
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
