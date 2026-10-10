' Parameter Table Tool - the Speed & feed window (a UserForm named Calculator). It stays
' open beside the sheet (modeless) and follows the op you are on. Two tabs:
'
'   SPEED & FEED   diameter + surface speed <-> RPM, feed per rev <-> per minute at that
'                  RPM, and a warning past the op's max spindle speed.
'   CHIP THINNING  how thick the chip really is at this op's feed, and the feed that gives
'                  the chip you want. A round insert's chip thins with the depth of cut; a
'                  straight edge's with the holder's ENTERING ANGLE - the chip is feed x
'                  sin(angle): 90 = square shoulder, no thinning; 45 = 0.71 x the feed.
'                  (US shops say lead angle = 90 - entering angle; both are shown.) The
'                  insert, its size and the angle come from the Tools page.
'
' Type in any box and the others follow. Filled from the op's row, in the op's units (its
' units column). "Set this op's feed" writes the feed through the cell's own rule, and one
' Ctrl+Z takes it back. How op windows behave together is in Panel.bas.
'
' tools\build_vba.ps1 creates it with these named controls - on the form btnPrev, cboOp,
' btnNext, btnHelp, lblInfo, mpg (a MultiPage), lblResult, btnClose; on its first page
' lbl / txt / hnt + Dia, Surf, Rpm, Rev, Min, and lblNote, btnSetRow; on its second
' lblInsert, optLead, optRound, lbl / txt / hnt + Kr, Lead, IC, Ap, Hex, Fn, and lblNow,
' lblThin, btnSetChip - and puts this code behind it.
Option Explicit

Private Const PAD As Single = 8
Private Const FIRST_ROW As Long = 3
Private Const COLS As String = "feed|feed_mode|speed|speed_mode|max_ss"      ' what the sheet shows of an op

Private busy As Boolean                     ' boxes are being filled in by the code
Private curRow As Long                      ' the op's row (0: none)
Private opRow() As Long                     ' the dropdown's rows, by list index
Private metric As Boolean                   ' the op's units are mm
Private css As Boolean                      ' the op's speed is a surface speed
Private maxRpm As Double                    ' its max spindle speed (0: none)
Private feedCell As Range                   ' its feed cell (Nothing: no op, or no feed)
Private feedPerMin As Boolean               ' its feed is per minute
Private speedFrom As String                 ' "surf" or "rpm": which speed box leads
Private feedFrom As String                  ' "rev" or "min": which feed box leads
Private kind As String                      ' its insert: round, lead, groove, prime, mill, none
Private chipFrom As String                  ' "hex" or "fn": which chip box was typed last
Private chipTyped As Boolean                ' a chip box was typed in (else they follow the feed)

' ============================================================ set up

Private Sub UserForm_Initialize()
    Dim c As Control
    busy = True
    Me.Caption = "Parameter Table - Speed & feed"
    For Each c In Me.Controls
        Select Case TypeName(c)
        Case "Label", "CommandButton", "OptionButton", "ComboBox": c.Font.Size = 9
        Case "TextBox"
            c.Font.Size = 11
            c.TextAlign = fmTextAlignRight
        End Select
        If Left$(c.Name, 3) = "hnt" Then         ' a column name: a small grey hint
            c.Font.Size = 8
            c.ForeColor = RGB(120, 128, 140)
        End If
    Next
    cboOp.Style = fmStyleDropDownList
    cboOp.ControlTipText = "The op this window is on - pick another"
    btnPrev.Caption = "<"
    btnPrev.ControlTipText = "The op above (hidden rows skipped)"
    btnNext.Caption = ">"
    btnNext.ControlTipText = "The op below (hidden rows skipped)"
    btnHelp.Caption = "?"
    btnHelp.ControlTipText = "How this window works (the user manual)"
    lblInfo.WordWrap = True
    lblInfo.ForeColor = RGB(91, 101, 115)
    lblResult.WordWrap = True
    lblNote.WordWrap = True
    lblInsert.WordWrap = True
    lblThin.WordWrap = True
    lblThin.ForeColor = RGB(91, 101, 115)
    lblNow.Font.Bold = True
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc
    mpg.Pages(0).Caption = "Speed & feed"
    mpg.Pages(1).Caption = "Chip thinning"
    mpg.Value = 0
    optLead.Caption = "Straight edge (lead angle)"
    optRound.Caption = "Round insert"
    hntDia.Caption = "where the cut is"
    hntSurf.Caption = "speed (CSS)"
    hntRpm.Caption = "speed (RPM)"
    hntRev.Caption = "feed (per rev)"
    hntMin.Caption = "feed (per min)"
    hntKr.Caption = "Tools page"
    hntLead.Caption = "= 90 - entering angle"
    hntIC.Caption = ""
    hntAp.Caption = ""
    hntHex.Caption = ""
    hntFn.Caption = "feed"
    speedFrom = "surf"
    feedFrom = "rev"
    chipFrom = "fn"
    kind = "none"
    optLead.Value = True
    busy = False
    Layout
    Units
    ShowKind
End Sub

' Shown on a scaled display, MSForms can drop controls back to their design size: put
' everything back where it belongs.
Private Sub UserForm_Activate()
    Layout
End Sub

Private Sub Layout()
    Dim w As Single, y As Single, pw As Single
    Me.Width = 400
    w = Me.InsideWidth - 2 * PAD
    y = PAD
    btnPrev.Move PAD, y, 24, 22
    cboOp.Move PAD + 28, y, w - 28 - 58, 22
    btnNext.Move PAD + w - 54, y, 24, 22
    btnHelp.Move PAD + w - 24, y, 24, 22
    y = y + 28
    lblInfo.Move PAD, y, w, 26: y = y + 30
    mpg.Move PAD, y, w, 284: y = y + 290
    lblResult.Move PAD, y, w - 84, 30
    btnClose.Move PAD + w - 76, y + 2, 76, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 32 + PAD
    pw = mpg.Width - 6

    ' Speed & feed
    y = 8
    Place lblDia, txtDia, hntDia, y: y = y + 26
    Place lblSurf, txtSurf, hntSurf, y: y = y + 26
    Place lblRpm, txtRpm, hntRpm, y: y = y + 34
    Place lblRev, txtRev, hntRev, y: y = y + 26
    Place lblMin, txtMin, hntMin, y: y = y + 32
    lblNote.Move 6, y, pw - 12, 42: y = y + 48
    btnSetRow.Move 6, y, pw - 12, 26

    ' Chip thinning: the angle rows and the round-insert rows share their places.
    y = 6
    lblInsert.Move 6, y, pw - 12, 38: y = y + 40
    optLead.Move 6, y, 170, 18
    optRound.Move 182, y, pw - 188, 18: y = y + 24
    Place lblKr, txtKr, hntKr, y
    Place lblIC, txtIC, hntIC, y: y = y + 26
    Place lblLead, txtLead, hntLead, y
    Place lblAp, txtAp, hntAp, y: y = y + 30
    lblNow.Move 6, y, pw - 12, 16: y = y + 20
    Place lblHex, txtHex, hntHex, y: y = y + 26
    Place lblFn, txtFn, hntFn, y: y = y + 28
    lblThin.Move 6, y, pw - 12, 28: y = y + 32
    btnSetChip.Move 6, y, pw - 12, 26
End Sub

Private Sub Place(ByVal lbl As MSForms.Label, ByVal box As MSForms.TextBox, ByVal hint As MSForms.Label, ByVal y As Single)
    lbl.Move 6, y + 3, 162, 16
    box.Move 170, y, 84, 20
    hint.Move 260, y + 4, mpg.Width - 270, 14
End Sub

Private Sub Units()
    Dim u As String
    u = IIf(metric, "mm", "in")
    lblDia.Caption = "Diameter (" & u & ")"
    lblSurf.Caption = "Surface speed (" & IIf(metric, "m/min", "SFM") & ")"
    lblRpm.Caption = "Spindle speed (RPM)"
    lblRev.Caption = "Feed per rev (" & u & ")"
    lblMin.Caption = "Feed per minute (" & u & ")"
    lblKr.Caption = "Entering angle (" & Chr$(176) & ")"
    lblLead.Caption = "Lead angle, US (" & Chr$(176) & ")"
    lblIC.Caption = "Insert diameter (" & u & ")"
    lblAp.Caption = "Depth of cut (" & u & ")"
    lblHex.Caption = "Chip thickness you want (" & u & ")"
    lblFn.Caption = "Feed per rev to program (" & u & ")"
End Sub

' ============================================================ the op

' Show an op (0: none) - from the ribbon (showCols: its feed and speed cells selected on
' the sheet), or set up for the checks (not).
Public Sub GoToRow(ByVal r As Long, ByVal showCols As Boolean)
    Dim rows As Collection
    If r <> curRow Then lblResult.Caption = ""
    Fill r, False
    If showCols And r > 0 Then
        Set rows = New Collection
        rows.Add r
        On Error Resume Next
        Panel.ShowColumns COLS, rows
    End If
End Sub

' The selection moved to other ops (Panel.SelectionMoved): follow the first. What was
' typed here and not written goes - this is a calculator.
Public Sub FollowRows(ByVal rows As Collection)
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub
    If rows(1) = curRow Then
        Fill curRow, True
    Else
        lblResult.Caption = ""
        Fill rows(1), False
    End If
End Sub

' Values changed under the window (undo, another window): read the same op again. What
' was typed about the insert and the diameter stays.
Public Sub Refresh()
    If curRow <= 0 Then Exit Sub
    If curRow > ParamTable.LastRow Then
        Fill 0, False
    Else
        Fill curRow, True
    End If
End Sub

Private Sub NoOp()
    busy = True
    curRow = 0
    Set feedCell = Nothing
    maxRpm = 0
    css = False
    feedPerMin = False
    speedFrom = "surf"
    feedFrom = "rev"
    kind = "none"
    Units
    BuildList
    cboOp.ListIndex = -1
    lblInfo.Caption = "Select an op (any cell in its row) - or just type in any box: the others follow."
    lblInsert.Caption = "No op: pick the kind of insert, then type its entering angle or its size and depth."
    optLead.Value = True
    busy = False
    ShowKind
    Note
End Sub

' Fill the window from an op's row (0: no op). keep: the same op read again.
Private Sub Fill(ByVal r As Long, ByVal keep As Boolean)
    Dim sp As Variant, fd As Variant, mx As Variant, cd As Variant
    If r <= 0 Then
        NoOp
        Exit Sub
    End If
    busy = True
    curRow = r
    sp = ParamTable.ValueAt(r, "speed")
    fd = ParamTable.ValueAt(r, "feed")
    mx = ParamTable.ValueAt(r, "max_ss")
    metric = (CStr(ParamTable.ValueAt(r, "units")) = "mm")
    css = (CStr(ParamTable.ValueAt(r, "speed_mode")) = "CSS")
    feedPerMin = (CStr(ParamTable.ValueAt(r, "feed_mode")) = "per min")
    speedFrom = IIf(css, "surf", "rpm")
    feedFrom = IIf(feedPerMin, "min", "rev")
    maxRpm = 0
    If IsNum(mx) Then maxRpm = CDbl(mx)
    Units
    If Not keep Then
        ' The op's own average cutting diameter, when the sheet has it.
        cd = ParamTable.ValueAt(r, "cut_dia")
        If Positive(cd) Then
            txtDia.Text = CStr(cd)
            hntDia.Caption = "the op's average (cut_dia)"
        Else
            txtDia.Text = ""
            hntDia.Caption = "where the cut is"
        End If
    End If
    txtSurf.Text = ""
    txtRpm.Text = ""
    txtRev.Text = ""
    txtMin.Text = ""
    If IsNum(sp) Then
        If css Then txtSurf.Text = CStr(sp) Else txtRpm.Text = CStr(sp)
    End If
    If IsNum(fd) Then
        If feedPerMin Then txtMin.Text = CStr(fd) Else txtRev.Text = CStr(fd)
    End If
    If ParamTable.ColOf("feed") > 0 Then
        Set feedCell = ParamTable.MainSheet.Cells(r, ParamTable.ColOf("feed"))
    Else
        Set feedCell = Nothing
    End If
    If Not keep Then InsertFrom r
    chipTyped = False
    chipFrom = "fn"
    busy = False
    If css Then FromSurface Else FromRpm
    Header
End Sub

' The op in the dropdown (its line refreshed - a comment may have changed).
Private Sub Header()
    Dim i As Long, found As Boolean
    busy = True
    For i = 0 To cboOp.ListCount - 1
        If opRow(i) = curRow Then found = True: Exit For
    Next
    If Not found Then
        BuildList
        For i = 0 To cboOp.ListCount - 1
            If opRow(i) = curRow Then found = True: Exit For
        Next
    End If
    If found Then
        cboOp.List(i) = Panel.OpLabel(curRow)
        cboOp.ListIndex = i
    Else
        cboOp.ListIndex = -1
    End If
    busy = False
End Sub

' Every op showing on the sheet (and the window's own, if it is filtered away).
Private Sub BuildList()
    Dim ws As Worksheet, r As Long, last As Long, n As Long
    Set ws = ParamTable.MainSheet
    last = ParamTable.LastRow
    cboOp.Clear
    ReDim opRow(0 To Application.Max(0, last - FIRST_ROW + 1))
    For r = FIRST_ROW To last
        If Not IsEmpty(ws.Cells(r, 1).Value) And (Not ws.Rows(r).Hidden Or r = curRow) Then
            cboOp.AddItem Panel.OpLabel(r)
            opRow(n) = r
            n = n + 1
        End If
    Next
End Sub

' The op's insert, from the Tools page: what kind it is, its size and the holder's angle,
' and the depth of cut from the row.
Private Sub InsertFrom(ByVal r As Long)
    Dim tp As String, t As Variant, lbl As String, code As String, ic As Variant, kr As Variant
    Dim roundDia As String, ap As String, apCol As String, what As String
    tp = UCase$(Trim$(CStr(ParamTable.ValueAt(r, "type"))))
    t = ParamTable.ValueAt(r, "tool")
    lbl = Trim$(CStr(ParamTable.ToolValue(r, "Insert")))
    code = Trim$(CStr(ParamTable.ToolValue(r, "Insert code")))
    ic = ParamTable.ToolValue(r, "Insert size (IC)")
    kr = ParamTable.ToolValue(r, "Entering angle")
    roundDia = ParamTable.RoundInsertDia(r, metric)
    ap = ParamTable.CalcDepth(r, apCol)
    kind = KindOf(tp, lbl, code, Positive(kr), roundDia <> "")
    If Positive(ic) Then
        txtIC.Text = CStr(ic)
        hntIC.Caption = "Tools page"
    Else
        txtIC.Text = roundDia
        hntIC.Caption = IIf(roundDia <> "", "2 x tool_radius", "")
    End If
    txtAp.Text = ap
    hntAp.Caption = apCol
    If Positive(kr) Then
        txtKr.Text = CStr(kr)
        hntKr.Caption = "Tools page"
    ElseIf kind = "groove" Then
        txtKr.Text = "90"
        hntKr.Caption = "a straight plunge"
    Else
        txtKr.Text = ""
        hntKr.Caption = "the holder's"
    End If
    txtLead.Text = Other90(txtKr.Text)
    optRound.Value = (kind = "round")
    optLead.Value = (kind <> "round")
    ShowKind
    If IsEmpty(t) Then
        what = "This op has no tool"
    ElseIf lbl = "" Then
        what = "T" & t & " - not on the Tools page"
    Else
        what = "T" & t & " insert: " & lbl
    End If
    Select Case kind
    Case "round": what = what & ".  A round insert: the chip thins with the depth of cut."
    Case "groove": what = what & ".  A grooving insert: a straight plunge, no thinning."
    Case "prime": what = what & ".  PrimeTurning ops set their own feeds from a chip thickness (pt_rough_chip) - not worked out here."
    Case "mill": what = "A mill op - radial chip thinning is not worked out here (the sheet has no cutter diameter or flute count)."
    Case "none": what = what & ".  Nothing to thin on this kind of op."
    Case Else
        If Positive(kr) Then
            what = what & ".  Holder's entering angle " & kr & Chr$(176) & " (Tools page)."
        Else
            what = what & ".  No entering angle on the Tools page - type the holder's (PCLNR / MCLNR: 95)."
        End If
    End Select
    lblInsert.Caption = what
End Sub

' What kind of insert an op cuts with: its kind first, then the Tools page's insert name
' and code; with neither saying, a tool radius only a round insert has.
Private Function KindOf(ByVal tp As String, ByVal lbl As String, ByVal code As String, ByVal hasAngle As Boolean, _
                        ByVal bigRadius As Boolean) As String
    Dim l As String
    l = LCase$(lbl)
    If tp = "CONTOUR" Or InStr(tp, "MILL") > 0 Then
        KindOf = "mill"
    ElseIf tp = "" Or tp = "DRILL" Or tp = "MANUAL" Or tp = "TRANSFORM" Then
        KindOf = "none"
    ElseIf tp = "PRIME" Or l Like "primeturning*" Then
        KindOf = "prime"
    ElseIf l Like "round*" Or l = "r" Or IsRoundCode(code) Or IsRoundCode(lbl) Then
        KindOf = "round"
    ElseIf tp = "GROOVE" Or tp = "PLUNGE ROUGH" Or l Like "groove*" Then
        KindOf = "groove"
    ElseIf hasAngle Then
        KindOf = "lead"
    ElseIf bigRadius Then
        KindOf = "round"
    Else
        KindOf = "lead"
    End If
End Function

' An ISO round insert's code: R, a clearance letter, letters, then figures (RCMT 10T3MO).
Private Function IsRoundCode(ByVal s As String) As Boolean
    IsRoundCode = UCase$(Trim$(s)) Like "R[ABCDEFGNP][A-Z]*#*"
End Function

' The angle rows for a straight edge, the size and depth rows for a round insert.
Private Sub ShowKind()
    Dim isRound As Boolean
    isRound = optRound.Value
    lblKr.Visible = Not isRound: txtKr.Visible = Not isRound: hntKr.Visible = Not isRound
    lblLead.Visible = Not isRound: txtLead.Visible = Not isRound: hntLead.Visible = Not isRound
    lblIC.Visible = isRound: txtIC.Visible = isRound: hntIC.Visible = isRound
    lblAp.Visible = isRound: txtAp.Visible = isRound: hntAp.Visible = isRound
End Sub

' ============================================================ working it out

Private Function Num(ByVal box As MSForms.TextBox) As Double
    If IsNumeric(box.Text) Then Num = CDbl(box.Text)
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    IsNum = IsNumeric(v)
End Function

Private Function Positive(ByVal v As Variant) As Boolean
    If IsNum(v) Then Positive = (CDbl(v) > 0)
End Function

' 90 minus an angle typed in a box ("" when it is not a number): entering <-> lead.
Private Function Other90(ByVal s As String) As String
    If IsNumeric(s) Then Other90 = CStr(Round(90 - CDbl(s), 3))
End Function

Private Sub SetBox(ByVal box As MSForms.TextBox, ByVal v As Double, ByVal fmt As String)
    Dim was As Boolean
    was = busy
    busy = True
    box.Text = IIf(v > 0, Format$(v, fmt), "")
    busy = was
End Sub

Private Function K() As Double
    K = IIf(metric, 1000, 12)
End Function

' The spindle speed the feeds go with: the RPM box - held to the max spindle speed when
' the op's speed is CSS (the control stops it there).
Private Function RunRpm() As Double
    RunRpm = Num(txtRpm)
    If css And maxRpm > 0 And RunRpm > maxRpm Then RunRpm = maxRpm
End Function

Private Sub FromSurface()                   ' diameter + surface speed -> RPM
    If Num(txtDia) > 0 And Num(txtSurf) > 0 Then SetBox txtRpm, K() * Num(txtSurf) / (Application.Pi() * Num(txtDia)), "0"
    AfterRpm
End Sub

Private Sub FromRpm()                       ' diameter + RPM -> surface speed
    If Num(txtDia) > 0 And Num(txtRpm) > 0 Then SetBox txtSurf, Num(txtRpm) * Application.Pi() * Num(txtDia) / K(), "0.0"
    AfterRpm
End Sub

' A new RPM: the feed that leads stays, the other follows.
Private Sub AfterRpm()
    If feedFrom = "min" Then FromMin Else FromRev
End Sub

Private Sub FromRev()                       ' per rev -> per minute
    If Num(txtRev) > 0 And RunRpm() > 0 Then SetBox txtMin, Num(txtRev) * RunRpm(), "0.0###"
    Note
End Sub

Private Sub FromMin()                       ' per minute -> per rev
    If Num(txtMin) > 0 And RunRpm() > 0 Then SetBox txtRev, Num(txtMin) / RunRpm(), "0.00000"
    Note
End Sub

' The lines under the boxes, the Set button, and the chip tab's feed, after any change.
Private Sub Note()
    Dim mx As String
    mx = Format$(maxRpm, "0")
    If maxRpm > 0 And Num(txtRpm) > maxRpm Then
        If css Then
            lblNote.Caption = "Above this op's max spindle speed of " & mx & " RPM: at this diameter the spindle stops at " & _
                              mx & " and the surface speed drops. The feed per minute here is at " & mx & " RPM."
        Else
            lblNote.Caption = "Above this op's max spindle speed of " & mx & " RPM - the spindle will not go faster."
        End If
        lblNote.ForeColor = RGB(180, 35, 24)
    Else
        If maxRpm > 0 Then
            lblNote.Caption = "This op's max spindle speed: " & mx & " RPM."
        ElseIf css And curRow > 0 Then
            lblNote.Caption = "This op's speed is CSS with no max spindle speed."
        Else
            lblNote.Caption = ""
        End If
        lblNote.ForeColor = RGB(91, 101, 115)
    End If
    Info
    SetCaption btnSetRow, SpeedWrite(), "type a feed"
    FollowFeed
End Sub

Private Sub Info()
    If curRow <= 0 Then Exit Sub
    lblInfo.Caption = "From op " & ParamTable.MainSheet.Cells(curRow, 1).Value & "'s row, in " & IIf(metric, "mm", "inches") & _
                      ". Type in any box - the others follow." & _
                      IIf(css And Num(txtDia) <= 0, "  Its speed is CSS: type a diameter for the RPM there.", "")
End Sub

' Untouched, the chip boxes go with the op's feed per rev.
Private Sub FollowFeed()
    If Not chipTyped Then
        busy = True
        txtFn.Text = IIf(Num(txtRev) > 0, txtRev.Text, "")
        busy = False
        chipFrom = "fn"
    End If
    Chip
End Sub

' How much thinner than the feed the chip is (0: not known yet).
Private Function Factor() As Double
    If optRound.Value Then
        Factor = Planner.ChipFactor(Num(txtIC), Num(txtAp))
    Else
        Factor = Planner.AngleChipFactor(Num(txtKr))
    End If
End Function

' Chip thinning: the box not typed last follows the one that was; then the lines that say
' what it means, and the Set button.
Private Sub Chip()
    Dim f As Double, ang As Double, feedNow As Double, deg As String
    deg = Chr$(176)
    f = Factor()
    feedNow = Num(txtRev)
    If f <= 0 Then
        If optRound.Value Then
            lblThin.Caption = "Type the insert diameter and the depth of cut."
        Else
            lblThin.Caption = "Type the holder's entering angle (90 = square shoulder, no thinning) - or its US lead angle."
        End If
        If chipFrom = "hex" Then SetBox txtFn, 0, "0" Else SetBox txtHex, 0, "0"
        lblNow.Caption = IIf(feedNow > 0, "Now: feed " & RoundFeed(feedNow) & " per rev.", "")
        SetCaption btnSetChip, ChipWrite(), "type a chip thickness or a feed"
        Exit Sub
    End If
    If chipFrom = "hex" Then
        If Num(txtHex) > 0 Then SetBox txtFn, Num(txtHex) / f, "0.00000" Else SetBox txtFn, 0, "0"
    Else
        If Num(txtFn) > 0 Then SetBox txtHex, Num(txtFn) * f, "0.00000" Else SetBox txtHex, 0, "0"
    End If
    If feedNow > 0 Then
        lblNow.Caption = "Now: feed " & RoundFeed(feedNow) & " per rev makes a " & Format$(feedNow * f, "0.00000") & " chip."
    ElseIf feedPerMin And curRow > 0 Then
        lblNow.Caption = "Now: the feed is per minute - give a diameter or RPM on the first tab."
    Else
        lblNow.Caption = ""
    End If
    If f >= 0.99999 Then
        If optRound.Value Then
            lblThin.Caption = "At half the insert or deeper the chip is as thick as the feed - no thinning."
        Else
            lblThin.Caption = "Entering angle 90" & deg & " (lead 0" & deg & "): the chip is as thick as the feed - no thinning."
        End If
    ElseIf optRound.Value Then
        ang = Atn(f / Sqr(1 - f * f)) * 180 / Application.Pi()
        lblThin.Caption = "Entering angle " & Format$(ang, "0") & deg & ": the chip is " & Format$(f, "0.00") & _
                          " x the feed - program " & Format$(1 / f, "0.00") & " x the chip you want."
    Else
        lblThin.Caption = "Entering angle " & CStr(Round(Num(txtKr), 2)) & deg & " (lead " & Other90(txtKr.Text) & deg & _
                          "): the chip is " & Format$(f, "0.00") & " x the feed - program " & Format$(1 / f, "0.00") & _
                          " x the chip you want."
    End If
    SetCaption btnSetChip, ChipWrite(), "type a chip thickness or a feed"
End Sub

' A feed to 4 figures (0.01155, 12.35) - what the Set buttons write.
Private Function RoundFeed(ByVal v As Double) As Double
    Dim e As Long
    If v <= 0 Then Exit Function
    e = Int(Log(v) / Log(10#))
    RoundFeed = Round(v, Application.Max(0, 3 - e))
End Function

' What the first tab's Set writes: the feed box in the op's own units (0: none).
Private Function SpeedWrite() As Double
    SpeedWrite = RoundFeed(IIf(feedPerMin, Num(txtMin), Num(txtRev)))
End Function

' What the chip tab's Set writes: the feed to program, per rev - or per minute at the
' first tab's RPM when the op's feed is per minute (0: cannot).
Private Function ChipWrite() As Double
    If Num(txtFn) <= 0 Then Exit Function
    If feedPerMin Then
        If RunRpm() > 0 Then ChipWrite = RoundFeed(Num(txtFn) * RunRpm())
    Else
        ChipWrite = RoundFeed(Num(txtFn))
    End If
End Function

' A Set button says what it would write, live; off when there is nothing to write.
Private Sub SetCaption(ByVal btn As MSForms.CommandButton, ByVal v As Double, ByVal ask As String)
    Dim cur As Variant, u As String
    u = IIf(feedPerMin, " per min", " per rev")
    If feedCell Is Nothing Then
        btn.Caption = "Set this op's feed (select an op first)"
        btn.Enabled = False
    ElseIf v <= 0 Then
        If feedPerMin And btn Is btnSetChip And Num(txtFn) > 0 Then ask = "its feed is per minute: give a diameter or RPM on the first tab"
        btn.Caption = "Set this op's feed (" & ask & ")"
        btn.Enabled = False
    Else
        cur = feedCell.Value
        If IsNum(cur) Then
            If ParamTable.SameValue(cur, v) Then
                btn.Caption = "This op's feed is " & cur & u & " - no change"
                btn.Enabled = False
            Else
                btn.Caption = "Set this op's feed:  " & cur & "  ->  " & v & u
                btn.Enabled = True
            End If
        Else
            btn.Caption = "Set this op's feed to " & v & u
            btn.Enabled = True
        End If
    End If
End Sub

' Write the op's feed through its cell's rule, as one undo; say what happened here.
Private Sub WriteFeed(ByVal v As Double)
    Dim old As Variant, ok As Boolean, op As Variant, why As String, u As String
    If feedCell Is Nothing Or v <= 0 Then Exit Sub
    op = ParamTable.MainSheet.Cells(feedCell.Row, 1).Value
    u = IIf(feedPerMin, " per min", " per rev")
    old = feedCell.Value
    Panel.BeginEdit "Speed & feed (op " & op & " feed)"
    ok = ParamTable.TryWrite(feedCell, v)
    Panel.EndEdit
    Panel.RefreshWindows
    Refresh
    If ok Then
        lblResult.Caption = "op " & op & " feed " & old & " -> " & v & u & " - Ctrl+Z to undo"
        lblResult.ForeColor = RGB(33, 115, 70)
    Else
        On Error Resume Next
        why = feedCell.Validation.ErrorMessage
        On Error GoTo 0
        If why = "" Then why = v & " is outside this feed's limits, or this op takes no feed."
        lblResult.Caption = "Refused - op " & op & " feed: " & why
        lblResult.ForeColor = RGB(180, 35, 24)
    End If
End Sub

' ============================================================ moving to another op

Private Sub StepOp(ByVal dirn As Long)
    Dim ws As Worksheet, r As Long, last As Long
    Set ws = ParamTable.MainSheet
    last = ParamTable.LastRow
    r = curRow
    If r = 0 Then r = IIf(dirn > 0, FIRST_ROW - 1, last + 1)
    Do
        r = r + dirn
        If r < FIRST_ROW Or r > last Then
            lblResult.Caption = IIf(dirn > 0, "This is the last op.", "This is the first op.")
            lblResult.ForeColor = RGB(91, 101, 115)
            Exit Sub
        End If
    Loop While ws.Rows(r).Hidden Or IsEmpty(ws.Cells(r, 1).Value)
    MoveTo r
End Sub

' The window's own move to another op: fill from it, show its feed and speed on the sheet,
' and tell the other open windows - the sheet's selection is that op now.
Private Sub MoveTo(ByVal r As Long)
    Dim rows As Collection
    lblResult.Caption = ""
    Fill r, False
    Set rows = New Collection
    rows.Add r
    On Error Resume Next
    Panel.ShowColumns COLS, rows
    Panel.SelectionMoved Selection
End Sub

' ============================================================ events

Private Sub cboOp_Change()
    If busy Then Exit Sub
    If cboOp.ListIndex < 0 Then Exit Sub
    If opRow(cboOp.ListIndex) <> curRow Then MoveTo opRow(cboOp.ListIndex)
End Sub

Private Sub btnPrev_Click()
    StepOp -1
End Sub

Private Sub btnNext_Click()
    StepOp 1
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp "calculator"
End Sub

Private Sub txtDia_Change()
    If busy Then Exit Sub
    If speedFrom = "rpm" Then FromRpm Else FromSurface
End Sub

Private Sub txtSurf_Change()
    If busy Then Exit Sub
    speedFrom = "surf"
    FromSurface
End Sub

Private Sub txtRpm_Change()
    If busy Then Exit Sub
    speedFrom = "rpm"
    FromRpm
End Sub

Private Sub txtRev_Change()
    If busy Then Exit Sub
    feedFrom = "rev"
    FromRev
End Sub

Private Sub txtMin_Change()
    If busy Then Exit Sub
    feedFrom = "min"
    FromMin
End Sub

Private Sub txtKr_Change()                  ' entering angle -> lead angle
    If busy Then Exit Sub
    busy = True
    txtLead.Text = Other90(txtKr.Text)
    busy = False
    Chip
End Sub

Private Sub txtLead_Change()                ' lead angle -> entering angle
    If busy Then Exit Sub
    busy = True
    txtKr.Text = Other90(txtLead.Text)
    busy = False
    Chip
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
    chipTyped = True
    Chip
End Sub

Private Sub txtFn_Change()
    If busy Then Exit Sub
    chipFrom = "fn"
    chipTyped = True
    Chip
End Sub

Private Sub optLead_Click()
    If busy Then Exit Sub
    ShowKind
    Chip
End Sub

Private Sub optRound_Click()
    If busy Then Exit Sub
    ShowKind
    Chip
End Sub

Private Sub btnSetRow_Click()
    WriteFeed SpeedWrite()
End Sub

Private Sub btnSetChip_Click()
    WriteFeed ChipWrite()
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' However it closes (its button, its X, Unload): out of the windows that follow.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Panel.Closed Me
    ParamTable.CalcClosed Me
End Sub

' ============================================================ for the checks

' Type into a box by name (a tick or option: 1 / 0; a list: "#2" picks its item).
Public Sub TypeIn(ByVal name As String, ByVal s As String)
    Dim c As Object
    Set c = Me.Controls(name)
    Select Case TypeName(c)
    Case "OptionButton", "CheckBox": c.Value = (s = "1" Or LCase$(s) = "true")
    Case "ComboBox"
        If Left$(s, 1) = "#" Then c.ListIndex = CLng(Mid$(s, 2)) Else c.Text = s
    Case Else: c.Text = s
    End Select
End Sub

' Read a control back: a box's text, a label's or button's caption, an option as 1 / 0.
Public Function Field(ByVal name As String) As String
    Dim c As Object
    Set c = Me.Controls(name)
    Select Case TypeName(c)
    Case "Label", "CommandButton": Field = c.Caption
    Case "OptionButton", "CheckBox": Field = IIf(c.Value, "1", "0")
    Case Else: Field = c.Text
    End Select
End Function

' Press a button (only when it can be pressed).
Public Sub Press(ByVal name As String)
    Select Case name
    Case "btnSetRow": If btnSetRow.Enabled Then btnSetRow_Click
    Case "btnSetChip": If btnSetChip.Enabled Then btnSetChip_Click
    Case "btnPrev": btnPrev_Click
    Case "btnNext": btnNext_Click
    End Select
End Sub
