' Parameter Table Tool - the coolant picker (a UserForm named CoolantPicker): one tick box
' per coolant the machine has; the ticked ones are turned on at the chosen timing.
'
' tools\build_vba.ps1 creates the form with these named controls - lblInfo, lblResult,
' btnOK, btnCancel, btnClear and the tick boxes chk1..chk12 - and puts this code behind it.
' Layout and wording are all here.
Option Explicit

Public Accepted As Boolean

Private Const MAX_BOXES As Long = 12
Private Const PAD As Single = 8
Private Const ROW_H As Single = 20
Private Const COL_W As Single = 170

Private names() As String                   ' 1-based, matches chk1..chkN
Private nBoxes As Long

Private Sub UserForm_Initialize()
    Dim i As Long
    Accepted = False
    Me.Caption = "Parameter Table - coolant"
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblResult.Font.Size = 9
    lblResult.Font.Bold = True
    btnOK.Caption = "OK"
    btnOK.Default = True                    ' Enter
    btnCancel.Caption = "Cancel"
    btnCancel.Cancel = True                 ' Esc
    btnClear.Caption = "Clear"
    For i = 1 To MAX_BOXES
        Box(i).Visible = False
        Box(i).Font.Size = 10
    Next
End Sub

Private Function Box(ByVal i As Long) As MSForms.CheckBox
    Set Box = Me.Controls("chk" & i)
End Function

' The coolants to offer (no "none"), which of them some rows' machines lack, what to
' tick to start with ("Flood + Mist" or ""), and a line saying what is being set.
Public Sub Setup(ByVal offer As Variant, ByVal partial As Variant, ByVal current As String, ByVal info As String)
    Dim i As Long, cols As Long, perCol As Long, r As Long, k As Long, anyPartial As Boolean, have As String
    nBoxes = UBound(offer) - LBound(offer) + 1
    If nBoxes > MAX_BOXES Then nBoxes = MAX_BOXES
    ReDim names(1 To IIf(nBoxes < 1, 1, nBoxes))
    have = " + " & LCase$(current) & " + "
    For i = 1 To nBoxes
        names(i) = offer(LBound(offer) + i - 1)
        With Box(i)
            .Caption = names(i)
            If partial(LBound(partial) + i - 1) Then
                .Caption = .Caption & "  *"
                anyPartial = True
            End If
            .Value = (InStr(have, " + " & LCase$(names(i)) & " + ") > 0)
            .Visible = True
        End With
    Next
    If anyPartial Then info = info & vbCrLf & "* not on every selected row's machine - those rows are skipped."
    lblInfo.Caption = info

    ' One column up to 6, then two.
    cols = IIf(nBoxes > 6, 2, 1)
    perCol = (nBoxes + cols - 1) \ cols
    Me.Width = Application.Max(300, cols * COL_W + 2 * PAD + 12)
    lblInfo.Move PAD, PAD, Me.InsideWidth - 2 * PAD, IIf(anyPartial, 40, 28)
    For i = 1 To nBoxes
        k = (i - 1) \ perCol: r = (i - 1) Mod perCol
        Box(i).Move PAD + k * COL_W, lblInfo.Top + lblInfo.Height + 4 + r * ROW_H, COL_W - 4, ROW_H
    Next
    r = lblInfo.Top + lblInfo.Height + 4 + perCol * ROW_H + 6
    lblResult.Move PAD, r, Me.InsideWidth - 2 * PAD, 16
    r = r + 22
    btnClear.Move PAD, r, 64, 24
    btnOK.Move Me.InsideWidth - 2 * 72 - 2 * PAD + 4, r, 68, 24
    btnCancel.Move Me.InsideWidth - 72 - PAD + 4, r, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + r + 24 + PAD
    UpdateResult
End Sub

' What OK sets: the ticked names joined by " + ", or "none" when nothing is ticked.
Public Function Result() As String
    Dim i As Long, out As String
    For i = 1 To nBoxes
        If Box(i).Value Then out = out & IIf(out = "", "", " + ") & names(i)
    Next
    If out = "" Then out = "none"
    Result = out
End Function

' The boxes as shown, joined by "|" (for the checks).
Public Function Captions() As String
    Dim i As Long, out As String
    For i = 1 To nBoxes
        out = out & IIf(i = 1, "", "|") & Box(i).Caption
    Next
    Captions = out
End Function

' For the checks: tick or untick a box by its coolant's position.
Public Sub Tick(ByVal i As Long, ByVal on_ As Boolean)
    If i >= 1 And i <= nBoxes Then Box(i).Value = on_
End Sub

Private Sub UpdateResult()
    lblResult.Caption = "Will set:  " & Result()
End Sub

Private Sub chk1_Click(): UpdateResult: End Sub
Private Sub chk2_Click(): UpdateResult: End Sub
Private Sub chk3_Click(): UpdateResult: End Sub
Private Sub chk4_Click(): UpdateResult: End Sub
Private Sub chk5_Click(): UpdateResult: End Sub
Private Sub chk6_Click(): UpdateResult: End Sub
Private Sub chk7_Click(): UpdateResult: End Sub
Private Sub chk8_Click(): UpdateResult: End Sub
Private Sub chk9_Click(): UpdateResult: End Sub
Private Sub chk10_Click(): UpdateResult: End Sub
Private Sub chk11_Click(): UpdateResult: End Sub
Private Sub chk12_Click(): UpdateResult: End Sub

Private Sub btnClear_Click()
    Dim i As Long
    For i = 1 To nBoxes
        Box(i).Value = False
    Next
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
