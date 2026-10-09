' Parameter Table Tool - the speed and feed calculator (a UserForm named Calculator).
' Type in any box and the others follow: diameter + surface speed <-> RPM, feed per rev
' <-> per minute at that RPM, and - for a round insert in dynamic turning - the chip
' thickness <-> the feed per rev that gives it at a depth. Filled from the row you are on;
' "Set op N feed" puts the result into that row (through the cell's own rule).
'
' tools\build_vba.ps1 creates it with these named controls - lblInfo, lblDia, txtDia,
' lblSurf, txtSurf, lblRpm, txtRpm, lblRev, txtRev, lblMin, txtMin, chkMetric, lblNote,
' btnClose, lblChip, lblIC, txtIC, lblAp, txtAp, lblHex, txtHex, lblFn, txtFn, lblThin,
' btnUseFeed, btnSetRow - and puts this code behind it.
Option Explicit

Private Const PAD As Single = 8
Private busy As Boolean                     ' a box is being filled in by the others
Private maxRpm As Double
Private chipFrom As String                  ' "hex" or "fn": which of the two was typed last
Private feedCell As Range                   ' the row's feed (Nothing: no row)
Private feedPerMin As Boolean

Private Sub UserForm_Initialize()
    Dim y As Single
    Me.Caption = "Parameter Table - speed, feed and chip"
    Me.Width = 340
    lblInfo.Font.Size = 9
    lblInfo.WordWrap = True
    lblNote.Font.Size = 9
    lblNote.WordWrap = True
    lblThin.Font.Size = 9
    lblThin.WordWrap = True
    lblThin.ForeColor = RGB(91, 101, 115)
    lblChip.Font.Size = 9
    lblChip.Font.Bold = True
    lblChip.Caption = "Round insert, dynamic turning - chip thinning"
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc
    btnUseFeed.Caption = "Use as feed per rev"
    btnSetRow.Caption = "Set the row's feed"
    btnSetRow.Enabled = False
    chkMetric.Caption = "Metric (mm, m/min)"
    chipFrom = "hex"

    y = PAD
    lblInfo.Move PAD, y, Me.InsideWidth - 2 * PAD, 28: y = y + 32
    Row lblDia, txtDia, y: y = y + 26
    Row lblSurf, txtSurf, y: y = y + 26
    Row lblRpm, txtRpm, y: y = y + 34
    Row lblRev, txtRev, y: y = y + 26
    Row lblMin, txtMin, y: y = y + 30
    chkMetric.Move PAD, y, 200, 18: y = y + 26
    lblChip.Move PAD, y, Me.InsideWidth - 2 * PAD, 14: y = y + 18
    Row lblIC, txtIC, y: y = y + 26
    Row lblAp, txtAp, y: y = y + 26
    Row lblHex, txtHex, y: y = y + 26
    Row lblFn, txtFn, y: y = y + 26
    lblThin.Move PAD, y, Me.InsideWidth - 2 * PAD, 28: y = y + 30
    btnUseFeed.Move Me.InsideWidth - PAD - 130, y, 130, 22: y = y + 30
    lblNote.Move PAD, y, Me.InsideWidth - 2 * PAD, 28: y = y + 32
    btnSetRow.Move PAD, y, Me.InsideWidth - 2 * PAD - 80, 24
    btnClose.Move Me.InsideWidth - 72 - PAD + 4, y, 68, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
    Units
End Sub

Private Sub Row(ByVal lbl As MSForms.Label, ByVal box As MSForms.TextBox, ByVal y As Single)
    lbl.Move PAD, y + 3, 170, 16
    lbl.Font.Size = 9
    box.Move PAD + 176, y, Me.InsideWidth - 2 * PAD - 176, 20
    box.Font.Size = 11
    box.TextAlign = fmTextAlignRight
End Sub

Private Sub Units()
    Dim m As Boolean, u As String
    m = chkMetric.Value
    u = IIf(m, "mm", "in")
    lblDia.Caption = "Diameter (" & u & ")"
    lblSurf.Caption = "Surface speed (" & IIf(m, "m/min", "SFM") & ")"
    lblRpm.Caption = "Spindle speed (RPM)"
    lblRev.Caption = "Feed per rev (" & u & ")"
    lblMin.Caption = "Feed per minute (" & u & ")"
    lblIC.Caption = "Insert diameter (" & u & ")"
    lblAp.Caption = "Depth / stepover (" & u & ")"
    lblHex.Caption = "Chip thickness wanted (" & u & ")"
    lblFn.Caption = "Feed per rev to program (" & u & ")"
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

' The row the calculator can write to (its feed cell and whether that feed is per minute),
' and the round insert's diameter and depth to start the chip boxes with.
Public Sub SetRow(ByVal cell As Range, ByVal perMinute As Boolean, ByVal insertDia As String, ByVal depth As String)
    Set feedCell = cell
    feedPerMin = perMinute
    busy = True
    txtIC.Text = insertDia
    txtAp.Text = depth
    If Not perMinute Then txtFn.Text = txtRev.Text
    busy = False
    chipFrom = "fn"                         ' the row's feed: show the chip it makes
    Chip
    SetRowCaption
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
    SetRowCaption
End Sub

' Chip thinning: the box not typed last follows the one that was.
Private Sub Chip()
    Dim f As Double, ang As Double
    f = Planner.ChipFactor(Num(txtIC), Num(txtAp))
    If f <= 0 Then
        lblThin.Caption = "Type the insert diameter and the depth (or stepover) it cuts at."
        Exit Sub
    End If
    If chipFrom = "hex" Then
        If Num(txtHex) > 0 Then SetBox txtFn, Num(txtHex) / f, "0.00000" Else SetBox txtFn, 0, "0"
    Else
        If Num(txtFn) > 0 Then SetBox txtHex, Num(txtFn) * f, "0.00000" Else SetBox txtHex, 0, "0"
    End If
    If f >= 0.99999 Then
        lblThin.Caption = "At half the insert or deeper the chip is as thick as the feed - no thinning."
    Else
        ang = Atn(f / Sqr(1 - f * f)) * 180 / Application.Pi()
        lblThin.Caption = "Entering angle " & Format$(ang, "0") & Chr$(176) & ": the chip is " & Format$(f, "0.00") & _
                          " x the feed - program " & Format$(1 / f, "0.00") & " x the chip you want."
    End If
End Sub

' "Set op 5 feed: 0.01 -> 0.0123 per rev" - what the button would write, live.
Private Sub SetRowCaption()
    Dim v As Double, cur As Variant
    If feedCell Is Nothing Then Exit Sub
    v = IIf(feedPerMin, Num(txtMin), Num(txtRev))
    cur = feedCell.Value
    If v <= 0 Then
        btnSetRow.Caption = "Set the row's feed (type a feed " & IIf(feedPerMin, "per minute", "per rev") & ")"
        btnSetRow.Enabled = False
    ElseIf IsNumeric(cur) And Not IsEmpty(cur) And VarType(cur) <> vbString Then
        btnSetRow.Enabled = Abs(CDbl(cur) - RoundFeed(v)) > 1E-12
        btnSetRow.Caption = "Set op " & feedCell.Worksheet.Cells(feedCell.Row, 1).Value & " feed:  " & cur & "  ->  " & RoundFeed(v) & _
                            IIf(feedPerMin, " per min", " per rev")
    Else
        btnSetRow.Enabled = True
        btnSetRow.Caption = "Set op " & feedCell.Worksheet.Cells(feedCell.Row, 1).Value & " feed to " & RoundFeed(v)
    End If
End Sub

Private Function RoundFeed(ByVal v As Double) As Double
    Dim e As Long
    e = Int(Log(v) / Log(10#))
    RoundFeed = Round(v, Application.Max(0, 3 - e))
End Function

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

Private Sub txtIC_Change()
    If Not busy Then Chip
End Sub

Private Sub txtAp_Change()
    If Not busy Then Chip
End Sub

Private Sub txtHex_Change()
    If busy Then Exit Sub
    chipFrom = "hex"
    Chip
End Sub

Private Sub txtFn_Change()
    If busy Then Exit Sub
    chipFrom = "fn"
    Chip
End Sub

Private Sub chkMetric_Click()
    Units
    If Not busy Then FromSurface
End Sub

' The chip feed into the feed-per-rev box (and on to per minute).
Private Sub btnUseFeed_Click()
    If Num(txtFn) <= 0 Then Exit Sub
    txtRev.Text = txtFn.Text
End Sub

Private Sub btnSetRow_Click()
    Dim v As Double
    If feedCell Is Nothing Then Exit Sub
    v = RoundFeed(IIf(feedPerMin, Num(txtMin), Num(txtRev)))
    If ParamTable.TryWrite(feedCell, v) Then
        SetRowCaption
        btnSetRow.Caption = "Done - op " & feedCell.Worksheet.Cells(feedCell.Row, 1).Value & " feed is " & v
    Else
        btnSetRow.Caption = "Refused - " & v & " is outside this feed's limits"
    End If
End Sub

' For the checks: type into a box by name, read one back, press a button.
Public Sub TypeIn(ByVal name As String, ByVal s As String)
    Me.Controls(name).Text = s
End Sub

Public Function Field(ByVal name As String) As String
    Select Case name
    Case "lblNote", "lblThin", "btnSetRow": Field = Me.Controls(name).Caption
    Case Else: Field = Me.Controls(name).Text
    End Select
End Function

Public Sub Press(ByVal name As String)
    Select Case name
    Case "btnUseFeed": btnUseFeed_Click
    Case "btnSetRow": If btnSetRow.Enabled Then btnSetRow_Click
    End Select
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
