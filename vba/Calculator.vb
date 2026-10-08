' Parameter Table Tool - the speed and feed calculator (a UserForm named Calculator).
' Type in any box and the others follow: diameter + surface speed <-> RPM, and
' feed per rev <-> per minute at that RPM. Filled from the row you are on.
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblDia, txtDia,
' lblSurf, txtSurf, lblRpm, txtRpm, lblRev, txtRev, lblMin, txtMin, chkMetric, lblNote,
' btnClose - and puts this code behind it.
Option Explicit

Private Const PAD As Single = 8
Private busy As Boolean                     ' a box is being filled in by the others
Private maxRpm As Double

Private Sub UserForm_Initialize()
    Dim y As Single
    Me.Caption = "Parameter Table - speed and feed"
    Me.Width = 330
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblNote.Font.Size = 9
    lblNote.WordWrap = True
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc
    chkMetric.Caption = "Metric (mm, m/min)"

    y = PAD
    lblInfo.Move PAD, y, Me.InsideWidth - 2 * PAD, 28: y = y + 32
    Row lblDia, txtDia, y: y = y + 26
    Row lblSurf, txtSurf, y: y = y + 26
    Row lblRpm, txtRpm, y: y = y + 34
    Row lblRev, txtRev, y: y = y + 26
    Row lblMin, txtMin, y: y = y + 30
    chkMetric.Move PAD, y, 200, 18: y = y + 24
    lblNote.Move PAD, y, Me.InsideWidth - 2 * PAD, 28: y = y + 32
    btnClose.Move Me.InsideWidth - 72 - PAD + 4, y, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    Units
End Sub

Private Sub Row(ByVal lbl As MSForms.Label, ByVal box As MSForms.TextBox, ByVal y As Single)
    lbl.Move PAD, y + 3, 150, 16
    lbl.Font.Size = 9
    box.Move PAD + 156, y, Me.InsideWidth - 2 * PAD - 156, 20
    box.Font.Size = 11
    box.TextAlign = fmTextAlignRight
End Sub

Private Sub Units()
    Dim m As Boolean
    m = chkMetric.Value
    lblDia.Caption = "Diameter (" & IIf(m, "mm", "in") & ")"
    lblSurf.Caption = "Surface speed (" & IIf(m, "m/min", "SFM") & ")"
    lblRpm.Caption = "Spindle speed (RPM)"
    lblRev.Caption = "Feed per rev (" & IIf(m, "mm", "in") & ")"
    lblMin.Caption = "Feed per minute (" & IIf(m, "mm", "in") & ")"
End Sub

' Start values (blank = leave empty), metric or not, the row's max RPM (0 = none) and a line about it.
Public Sub Setup(ByVal dia As String, ByVal surf As String, ByVal rpm As String, ByVal rev As String, _
                 ByVal perMin As String, ByVal metric As Boolean, ByVal maxSpindle As Double, ByVal info As String)
    busy = True
    chkMetric.Value = metric
    Units
    txtDia.Text = dia
    txtSurf.Text = surf
    txtRpm.Text = rpm
    txtRev.Text = rev
    txtMin.Text = perMin
    maxRpm = maxSpindle
    lblInfo.Caption = info
    busy = False
    If surf <> "" Then FromSurface Else If rpm <> "" Then FromRpm
    If rev <> "" Then FromRev Else If perMin <> "" Then FromMin
    Note
End Sub

Private Function Num(ByVal box As MSForms.TextBox) As Double
    If IsNumeric(box.Text) Then Num = CDbl(box.Text)
End Function

Private Sub SetBox(ByVal box As MSForms.TextBox, ByVal v As Double, ByVal fmt As String)
    busy = True
    box.Text = IIf(v > 0, Format$(v, fmt), "")
    busy = False
End Sub

Private Function K() As Double
    K = IIf(chkMetric.Value, 1000, 12)
End Function

Private Sub FromSurface()                   ' diameter + surface -> RPM
    If Num(txtDia) > 0 And Num(txtSurf) > 0 Then SetBox txtRpm, K() * Num(txtSurf) / (Application.Pi() * Num(txtDia)), "0"
    FromRev
End Sub

Private Sub FromRpm()                       ' diameter + RPM -> surface
    If Num(txtDia) > 0 And Num(txtRpm) > 0 Then SetBox txtSurf, Num(txtRpm) * Application.Pi() * Num(txtDia) / K(), "0.0"
    FromRev
End Sub

Private Sub FromRev()                       ' per rev -> per minute
    If Num(txtRev) > 0 And Num(txtRpm) > 0 Then SetBox txtMin, Num(txtRev) * Num(txtRpm), "0.0###"
    Note
End Sub

Private Sub FromMin()                       ' per minute -> per rev
    If Num(txtMin) > 0 And Num(txtRpm) > 0 Then SetBox txtRev, Num(txtMin) / Num(txtRpm), "0.00000"
    Note
End Sub

Private Sub Note()
    If maxRpm > 0 And Num(txtRpm) > maxRpm Then
        lblNote.Caption = "Above this row's max_ss of " & Format$(maxRpm, "0") & " RPM - at this diameter the spindle stops at " & _
                          Format$(maxRpm, "0") & " and the surface speed drops."
        lblNote.ForeColor = RGB(180, 35, 24)
    Else
        lblNote.Caption = IIf(maxRpm > 0, "max_ss on this row: " & Format$(maxRpm, "0") & " RPM.", "")
        lblNote.ForeColor = RGB(91, 101, 115)
    End If
End Sub

Private Sub txtDia_Change()
    If Not busy Then FromSurface
End Sub

Private Sub txtSurf_Change()
    If Not busy Then FromSurface
End Sub

Private Sub txtRpm_Change()
    If Not busy Then FromRpm
End Sub

Private Sub txtRev_Change()
    If Not busy Then FromRev
End Sub

Private Sub txtMin_Change()
    If Not busy Then FromMin
End Sub

Private Sub chkMetric_Click()
    Units
    If Not busy Then FromSurface
End Sub

' For the checks: type into a box by name, read one back.
Public Sub TypeIn(ByVal name As String, ByVal s As String)
    Me.Controls(name).Text = s
End Sub

Public Function Field(ByVal name As String) As String
    If name = "lblNote" Then Field = lblNote.Caption Else Field = Me.Controls(name).Text
End Function

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
