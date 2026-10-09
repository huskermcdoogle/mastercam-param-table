' Parameter Table Tool - the window for Set selected, Scale %, Copy from op, Apply to the
' tool's ops and the scenarios (a UserForm named EditBox): one box to type or pick in, a
' live line saying what OK would do - how many cells change and how many are refused -
' and a list of each cell, old -> new, before anything is written.
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblPrompt, cbo,
' lblPreview, lstDetail, btnOK, btnCancel - and puts this code behind it.
Option Explicit

Public Accepted As Boolean

Private Const PAD As Single = 8

Private mode As String                      ' "set", "scale", "copy", "tool", "scensave", "scenload", "scendel"
Private target As Range
Private keys As Variant                     ' what each list row stands for (copy: op_idn)
Private listOnly As Boolean

Private Sub UserForm_Initialize()
    Accepted = False
    Me.Width = 500
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblPrompt.Font.Size = 9
    lblPrompt.Font.Bold = True
    cbo.Font.Size = 11
    cbo.MatchEntry = fmMatchEntryNone
    lblPreview.Font.Size = 9
    lblPreview.WordWrap = True
    lstDetail.Font.Name = "Consolas"
    lstDetail.Font.Size = 9
    lstDetail.TabStop = False
    btnOK.Caption = "OK"
    btnOK.Default = True                    ' Enter
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                 ' Esc
End Sub

' Typing replaces what is in the box: it opens with all of it selected.
Private Sub UserForm_Activate()
    On Error Resume Next
    cbo.SetFocus
    cbo.SelStart = 0
    cbo.SelLength = Len(cbo.Text)
End Sub

' mode, the cells, a line about them, what to type, the list (may be empty), what each
' list row stands for (or Empty: the row's text), whether only the list is allowed, and
' what to start with (a list-only window: the row whose key or text it is).
Public Sub Setup(ByVal how As String, ByVal cells As Range, ByVal caption As String, ByVal info As String, _
                 ByVal prompt As String, ByVal choices As Variant, ByVal rowKeys As Variant, _
                 ByVal onlyList As Boolean, ByVal start As String)
    Dim i As Long, w As Single, y As Single
    mode = how
    Set target = cells
    keys = rowKeys
    listOnly = onlyList
    Me.Caption = caption
    lblInfo.Caption = info
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
    lblInfo.Move PAD, y, w, 30: y = y + 34
    lblPrompt.Move PAD, y, w, 14: y = y + 16
    cbo.Move PAD, y, w, 22: y = y + 30
    lblPreview.Move PAD, y, w, 42: y = y + 46
    lstDetail.Move PAD, y, w, 120: y = y + 128
    btnOK.Move Me.InsideWidth - 2 * 72 - PAD, y, 68, 24
    btnCancel.Move Me.InsideWidth - 72 - PAD + 4, y, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    Refresh
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

Private Sub Refresh()
    Dim p As String, d As Variant, i As Long
    p = ParamTable.EditPreview(mode, target, Value())
    btnOK.Enabled = (Left$(p, 1) = "1")
    lblPreview.Caption = Mid$(p, 3)
    lblPreview.ForeColor = IIf(btnOK.Enabled, RGB(33, 115, 70), RGB(91, 101, 115))
    lstDetail.Clear
    d = ParamTable.EditDetail(mode, target, Value())
    If IsArray(d) Then
        For i = LBound(d) To UBound(d)
            lstDetail.AddItem CStr(d(i))
        Next
    End If
End Sub

Private Sub cbo_Change()
    Refresh
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
