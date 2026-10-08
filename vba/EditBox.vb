' Parameter Table Tool - the window for Set selected, Scale % and Copy from op (a UserForm
' named EditBox): one box to type or pick in, and a live line saying what OK would do -
' how many cells change and how many are refused - before anything is written.
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblPrompt, cbo,
' lblPreview, btnOK, btnCancel - and puts this code behind it.
Option Explicit

Public Accepted As Boolean

Private Const PAD As Single = 8

Private mode As String                      ' "set", "scale" or "copy"
Private target As Range
Private keys As Variant                     ' what each list row stands for (copy: op_idn)
Private listOnly As Boolean

Private Sub UserForm_Initialize()
    Accepted = False
    Me.Width = 420
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblPrompt.Font.Size = 9
    lblPrompt.Font.Bold = True
    cbo.Font.Size = 11
    cbo.MatchEntry = fmMatchEntryNone
    lblPreview.Font.Size = 9
    lblPreview.WordWrap = True
    btnOK.Caption = "OK"
    btnOK.Default = True                    ' Enter
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                 ' Esc
End Sub

' mode, the cells, a line about them, what to type, the list (may be empty), what each
' list row stands for (or Empty: the row's text), whether only the list is allowed, and
' what to start with.
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
    If Not onlyList Then cbo.Text = start

    w = Me.InsideWidth - 2 * PAD
    y = PAD
    lblInfo.Move PAD, y, w, 30: y = y + 34
    lblPrompt.Move PAD, y, w, 14: y = y + 16
    cbo.Move PAD, y, w, 22: y = y + 30
    lblPreview.Move PAD, y, w, 42: y = y + 46
    btnOK.Move Me.InsideWidth - 2 * 72 - PAD, y, 68, 24
    btnCancel.Move Me.InsideWidth - 72 - PAD + 4, y, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    Refresh
End Sub

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

' For the checks: type into the box, or pick a list row.
Public Sub TypeIn(ByVal s As String)
    Dim i As Long
    If listOnly Then                        ' a pick-only list takes no typing: pick the matching row
        cbo.ListIndex = -1
        For i = 0 To cbo.ListCount - 1
            If cbo.List(i) = s Then cbo.ListIndex = i
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
    Dim p As String
    p = ParamTable.EditPreview(mode, target, Value())
    btnOK.Enabled = (Left$(p, 1) = "1")
    lblPreview.Caption = Mid$(p, 3)
    lblPreview.ForeColor = IIf(btnOK.Enabled, RGB(33, 115, 70), RGB(91, 101, 115))
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
