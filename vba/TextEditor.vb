' Parameter Table Tool - the manual-entry text editor (a UserForm named TextEditor).
'
' tools\build_vba.ps1 creates the form with five named controls - txt, lblInfo, lblCount,
' btnOK, btnCancel - and puts this code behind it. Everything else (sizes, fonts, layout,
' behaviour) is set here, so the form is described in text, not in a binary .frx.
Option Explicit

Public Accepted As Boolean

Private Const LIMIT As Long = 3111          ' what the Manual Entry dialog holds
Private Const PAD As Single = 6

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#End If

Private Sub UserForm_Initialize()
    Accepted = False
    Me.Caption = "Parameter Table - manual entry text"
    Me.Width = 560
    Me.Height = 440

    With txt
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .EnterKeyBehavior = True             ' Enter is a new line, not OK
        .TabKeyBehavior = False
        .Font.Name = "Consolas"
        .Font.Size = 11
    End With
    lblInfo.Font.Size = 9
    lblCount.Font.Size = 9
    btnOK.Caption = "OK"
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                  ' Esc
    Layout
End Sub

' Resizable: a plain UserForm has a fixed border; give it a sizing frame and maximise.
Private Sub UserForm_Activate()
    #If VBA7 Then
        Dim h As LongPtr
        h = FindWindowA("ThunderDFrame", Me.Caption)
        If h <> 0 Then
            SetWindowLongPtrA h, -16, GetWindowLongPtrA(h, -16) Or &H40000 Or &H10000   ' WS_THICKFRAME, WS_MAXIMIZEBOX
            DrawMenuBar h
        End If
    #End If
End Sub

Private Sub UserForm_Resize()
    Layout
End Sub

Private Sub Layout()
    Dim w As Single, h As Single, bw As Single, bh As Single
    w = Me.InsideWidth: h = Me.InsideHeight
    If w < 240 Or h < 160 Then Exit Sub
    bw = 72: bh = 24
    lblInfo.Move PAD, PAD, w - 2 * PAD, 14
    txt.Move PAD, PAD + 18, w - 2 * PAD, h - (PAD + 18) - bh - 2 * PAD
    lblCount.Move PAD, h - PAD - bh + 5, w - 2 * bw - 4 * PAD, 14
    btnOK.Move w - 2 * bw - 2 * PAD, h - PAD - bh, bw, bh
    btnCancel.Move w - bw - PAD, h - PAD - bh, bw, bh
End Sub

' What to edit, and a line of context above it.
Public Sub LoadText(ByVal text As String, ByVal info As String)
    ' Every line break as CR LF - what Mastercam's manual entry holds.
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    txt.Text = Replace(text, vbLf, vbCrLf)
    lblInfo.Caption = info
    UpdateCount
    txt.SelStart = 0
End Sub

Public Function EditedText() As String
    EditedText = txt.Text
End Function

Public Function CountText() As String
    CountText = lblCount.Caption
End Function

Public Function CanAccept() As Boolean
    CanAccept = btnOK.Enabled
End Function

Private Sub txt_Change()
    UpdateCount
End Sub

' "1,204 / 3,111 characters, 18 lines" - red, and OK off, past the limit: the load would
' refuse it, so it cannot be accepted here.
Private Sub UpdateCount()
    Dim n As Long, lines As Long
    n = Len(txt.Text)
    If n > 0 Then lines = UBound(Split(txt.Text, vbLf)) + 1
    lblCount.Caption = Format$(n, "#,##0") & " / " & Format$(LIMIT, "#,##0") & " characters,  " & _
                       lines & IIf(lines = 1, " line", " lines")
    If n > LIMIT Then
        lblCount.ForeColor = RGB(180, 35, 24)
        lblCount.Caption = lblCount.Caption & "  -  over the limit by " & Format$(n - LIMIT, "#,##0")
    Else
        lblCount.ForeColor = RGB(91, 101, 115)
    End If
    btnOK.Enabled = (n <= LIMIT)
End Sub

Private Sub btnOK_Click()
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
