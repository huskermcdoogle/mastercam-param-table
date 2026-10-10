' Parameter Table Tool - the Coolant window (a UserForm named CoolantPicker): how the
' coolant of the selected ops is set now, and one place to set it for all of them.
'
' Machines differ, and each op follows its own - decided per row by which of its coolant
' cells take input. A machine set up for V9 coolant has ONE setting: Off, Flood, Mist or
' Thru-tool (the coolant column). An X-style machine turns its own coolants on before,
' with and after the move (coolant_before / with / after - several at once, joined by
' " + "). A selection with both kinds shows both parts, each applying to its own ops.
'
' Where the ops agree the choice is picked (the box ticked); where they differ the part
' says "Mixed: Flood x3, none x1" and the box is grey - left grey, each op keeps its own.
' Apply writes only the parts changed, every cell checked against its own machine's
' coolants, as one Undo; ops whose machine lacks a coolant are listed with why.
'
' It stays open beside the sheet (modeless) and follows the ops selected (Panel.bas).
' tools\build_vba.ps1 creates the controls named here and puts this code behind them.
' Layout and wording are all here.
Option Explicit

Private Const PAD As Single = 8
Private Const MAX_NAMES As Long = 12
Private Const ROW_H As Single = 18
Private Const NAME_W As Single = 150
Private Const COL_W As Single = 76

Private opRows As Collection                ' the ops the window is on (sheet rows)
Private pending As Collection               ' where the sheet went while changes were not applied
Private busy As Boolean                     ' boxes set by code, not clicked
Private note As String                      ' a line about these ops
Private result As String                    ' what the last Apply did

Private colV9 As Long                       ' the V9 coolant column (0 = not on the sheet)
Private colX(0 To 2) As Long                ' coolant_before / with / after
Private v9Rows As Collection                ' ops on a V9 machine
Private xRows As Collection                 ' ops on an X-style machine
Private noRows As Collection                ' ops with no coolant at all
Private v9Start As Long                     ' the V9 choice picked at the start (0 = they differ)
Private xNames(1 To MAX_NAMES) As String    ' the X-style coolants offered, in the machines' order
Private xHits(1 To MAX_NAMES) As Long       ' how many of the ops' machines have each
Private xCount As Long                      ' how many are offered
Private xMore As String                     ' coolants past the twelve shown ("|a|b|")
Private xStart(1 To MAX_NAMES, 0 To 2) As Variant   ' each box as it started: True, False or Null (mixed)

Private Sub UserForm_Initialize()
    Dim i As Long, t As Long, lb As Variant
    Set opRows = New Collection
    Set v9Rows = New Collection
    Set xRows = New Collection
    Set noRows = New Collection
    Me.Caption = "Parameter Table - Coolant"
    Me.StartUpPosition = 0                  ' placed beside the sheet, not over it (ShowOn)
    Me.Width = 460

    lblOps.Font.Size = 10
    lblOps.Font.Bold = True
    btnPrev.Caption = "< Prev"
    btnPrev.ControlTipText = "The op above on the sheet"
    btnNext.Caption = "Next >"
    btnNext.ControlTipText = "The op below on the sheet"
    btnHelp.Caption = "?"
    btnHelp.ControlTipText = "How this window works (the user manual)"
    For Each lb In Array(btnPrev, btnNext, btnHelp)
        lb.TabStop = False                  ' Enter is Apply, never a step to another op
        lb.TakeFocusOnClick = False
    Next

    lblV9Head.Caption = "Coolant  -  machines with one coolant setting (V9)"
    lblV9Hint.Caption = "Pick one for all of these ops.          column: coolant"
    For i = 1 To 4
        With Me.Controls("optV" & i)
            .Caption = Choose(i, "Off", "Flood", "Mist", "Thru-tool")
            .GroupName = "v9"
            .Font.Size = 10
        End With
    Next
    lblXHead.Caption = "Coolant before / with / after the move  -  machines with X-style coolant"
    lblXHint.Caption = "Tick the coolants to turn on. A grey box: some of these ops have it - left as each op has it." & _
                       vbCrLf & "columns: coolant_before / coolant_with / coolant_after"
    lblXB.Caption = "Before"
    lblXW.Caption = "With"
    lblXA.Caption = "After"
    For Each lb In Array(lblXB, lblXW, lblXA)
        lb.Font.Size = 9
        lb.Font.Bold = True
        lb.TextAlign = fmTextAlignCenter
    Next
    For i = 1 To MAX_NAMES
        Me.Controls("lblX" & i).Font.Size = 10
        For t = 0 To 2
            Box(i, t).Caption = ""
        Next
    Next
    For Each lb In Array(lblV9Head, lblXHead)
        lb.Font.Size = 9
        lb.Font.Bold = True
    Next
    For Each lb In Array(lblV9Hint, lblXHint, lblXNote)
        lb.Font.Size = 8
        lb.ForeColor = RGB(91, 101, 115)
        lb.WordWrap = True
    Next
    For Each lb In Array(lblV9Now, lblXNow, lblNone, lblStatus)
        lb.Font.Size = 9
        lb.WordWrap = True
    Next
    linSep.Caption = ""
    linSep.BackColor = RGB(200, 205, 212)

    btnRevert.Caption = "Revert"
    btnRevert.ControlTipText = "Throw away what was changed here - show what the sheet has"
    btnApply.Caption = "Apply"
    btnApply.Default = True                 ' Enter
    btnApply.ControlTipText = "Write it to the sheet - Ctrl+Z takes it back"
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc
    Layout
    UpdateAll
End Sub

' The first choice has the focus - so Enter is Apply, not whichever button came first.
Private Sub UserForm_Activate()
    Dim t As Long
    On Error Resume Next
    If v9Rows.Count > 0 Then
        optV1.SetFocus
    ElseIf xRows.Count > 0 Then
        For t = 0 To 2
            If colX(t) > 0 Then Box(1, t).SetFocus: Exit For
        Next
    End If
End Sub

' A tick box by coolant (1..12) and timing (0 before, 1 with, 2 after).
Private Function Box(ByVal i As Long, ByVal t As Long) As MSForms.CheckBox
    Set Box = Me.Controls(Choose(t + 1, "chkB", "chkW", "chkA") & i)
End Function

Private Function Timing(ByVal t As Long) As String
    Timing = Choose(t + 1, "before the move", "with the move", "after the move")
End Function

' Place what these ops have; the window is as tall as they need.
Private Sub Layout()
    Dim w As Single, y As Single, i As Long, t As Long, on9 As Boolean, onX As Boolean, x0 As Single
    w = Me.InsideWidth
    btnHelp.Move w - PAD - 22, PAD, 22, 20
    btnNext.Move btnHelp.Left - 8 - 58, PAD, 58, 20
    btnPrev.Move btnNext.Left - 4 - 58, PAD, 58, 20
    lblOps.Move PAD, PAD + 3, btnPrev.Left - 2 * PAD, 16
    y = PAD + 30
    on9 = v9Rows.Count > 0
    onX = xRows.Count > 0

    ' One setting (V9).
    lblV9Head.Visible = on9
    lblV9Hint.Visible = on9
    lblV9Now.Visible = on9
    For i = 1 To 4
        Me.Controls("optV" & i).Visible = on9
    Next
    If on9 Then
        lblV9Head.Move PAD, y, w - 2 * PAD, 14: y = y + 15
        lblV9Hint.Move PAD, y, w - 2 * PAD, 12: y = y + 16
        For i = 1 To 4
            Me.Controls("optV" & i).Move PAD + 8 + (i - 1) * 100, y, 96, 18
        Next
        y = y + 22
        lblV9Now.Move PAD + 8, y, w - 2 * PAD - 8, 14: y = y + 22
    End If
    linSep.Visible = on9 And onX
    If on9 And onX Then linSep.Move PAD, y, w - 2 * PAD, 1: y = y + 9

    ' Before / with / after (X-style): a row per coolant, a column per timing.
    lblXHead.Visible = onX
    lblXHint.Visible = onX
    lblXNow.Visible = onX
    lblXNote.Visible = onX And lblXNote.Caption <> ""
    x0 = PAD + 8 + NAME_W
    For t = 0 To 2
        Me.Controls(Choose(t + 1, "lblXB", "lblXW", "lblXA")).Visible = onX And colX(t) > 0
    Next
    For i = 1 To MAX_NAMES
        Me.Controls("lblX" & i).Visible = onX And i <= xCount
        For t = 0 To 2
            Box(i, t).Visible = onX And i <= xCount And colX(t) > 0
        Next
    Next
    If onX Then
        lblXHead.Move PAD, y, w - 2 * PAD, 14: y = y + 15
        lblXHint.Move PAD, y, w - 2 * PAD, 24: y = y + 28
        For t = 0 To 2
            Me.Controls(Choose(t + 1, "lblXB", "lblXW", "lblXA")).Move x0 + t * COL_W, y, COL_W, 14
        Next
        y = y + 16
        For i = 1 To xCount
            Me.Controls("lblX" & i).Move PAD + 8, y + 1, NAME_W - 4, 15
            For t = 0 To 2
                Box(i, t).Move x0 + t * COL_W + COL_W / 2 - 7, y, 16, 16
            Next
            y = y + ROW_H
        Next
        y = y + 4
        lblXNow.Move PAD + 8, y, w - 2 * PAD - 8, 40: y = y + 44
        If lblXNote.Visible Then lblXNote.Move PAD + 8, y, w - 2 * PAD - 8, 24: y = y + 26
    End If

    lblNone.Visible = lblNone.Caption <> ""
    If lblNone.Visible Then lblNone.Move PAD, y, w - 2 * PAD, 28: y = y + 30
    lblStatus.Move PAD, y, w - 2 * PAD, 42: y = y + 46
    btnClose.Move w - PAD - 72, y, 72, 24
    btnApply.Move btnClose.Left - 6 - 72, y, 72, 24
    btnRevert.Move btnApply.Left - 6 - 72, y, 72, 24
    Me.Height = (Me.Height - Me.InsideHeight) + y + 24 + PAD
End Sub

' ---------------------------------------------------------------- the ops it is on

' From the ribbon (or Set on coolant cells): the window on these ops, the sheet brought to
' the coolant columns - shown beside the sheet, or brought to the front when it already is.
Public Sub ShowOn(ByVal rows As Collection)
    Dim placed As Boolean
    placed = Me.Visible
    OpenOn rows
    If Not placed Then
        Me.Left = Application.Left + Application.Width - Me.Width - 40
        Me.Top = Application.Top + 180
    End If
    Me.Show vbModeless
    Panel.Opened Me
End Sub

' The same without showing it. Changes not applied yet stay (the window says so) - they
' are never thrown away by a click elsewhere.
Public Sub OpenOn(ByVal rows As Collection)
    If Dirty() Then
        If Panel.RowsKey(rows) <> Panel.RowsKey(opRows) Then Set pending = rows
        UpdateAll
    Else
        LoadRows rows
        Focus
    End If
End Sub

' From the sheet (Panel.SelectionMoved): other ops picked - read them, unless changes here
' are not applied yet.
Public Sub FollowRows(ByVal rows As Collection)
    If Dirty() Then
        Set pending = rows
        UpdateAll
        Exit Sub
    End If
    LoadRows rows
End Sub

' After Undo or another window's Apply: the same ops read again (changes waiting are kept).
Public Sub Refresh()
    If Dirty() Then Exit Sub
    LoadRows opRows
End Sub

' Read the ops' coolant: the part each op is set in (its machine), and how each part is set
' now - the same on all of them (picked / ticked) or mixed (grey, and said in words).
Public Sub LoadRows(ByVal rows As Collection)
    Dim ws As Worksheet, r As Variant, t As Long, i As Long, n As Long, vals As Collection, s As String, part As Boolean
    Set ws = ParamTable.MainSheet
    busy = True
    Set opRows = rows
    Set pending = Nothing
    note = ""
    result = ""
    colV9 = ParamTable.ColOf("coolant")
    colX(0) = ParamTable.ColOf("coolant_before")
    colX(1) = ParamTable.ColOf("coolant_with")
    colX(2) = ParamTable.ColOf("coolant_after")

    If opRows.Count = 0 Then
        lblOps.Caption = "Select one or more ops (any cell in their rows)"
    ElseIf opRows.Count = 1 Then
        lblOps.Caption = Panel.OpLabel(opRows(1))
    Else
        lblOps.Caption = Panel.RowsText(opRows)
    End If

    ' Each op follows its own machine: the part whose cells take input on its row.
    Set v9Rows = New Collection
    Set xRows = New Collection
    Set noRows = New Collection
    xCount = 0
    xMore = ""
    For Each r In opRows
        Select Case Style(r)
            Case "v9": v9Rows.Add r
            Case "x": xRows.Add r: AddNames r
            Case Else: noRows.Add r
        End Select
    Next

    ' V9: one choice, picked when every op has the same.
    Set vals = New Collection
    For Each r In v9Rows
        vals.Add ParamTable.CellText(ws.Cells(r, colV9))
    Next
    V9Choices
    v9Start = 0
    If vals.Count > 0 Then
        If InStr(ParamTable.CountsText(vals), ",") = 0 Then
            For i = 1 To 4
                If LCase$(Me.Controls("optV" & i).Caption) = LCase$(Trim$(vals(1))) Then v9Start = i
            Next
        End If
    End If
    For i = 1 To 4
        Me.Controls("optV" & i).Value = (i = v9Start)
    Next
    lblV9Now.Caption = NowLine("Now", vals)
    lblV9Head.Caption = "Coolant  -  machines with one coolant setting (V9)" & IIf(v9Rows.Count > 1, "  -  " & v9Rows.Count & " ops", "")

    ' X-style: a box per coolant and timing - ticked when every op has it on, grey when
    ' only some do.
    s = ""
    For t = 0 To 2
        Set vals = New Collection
        If colX(t) > 0 Then
            For Each r In xRows
                vals.Add ParamTable.CellText(ws.Cells(r, colX(t)))
            Next
            If vals.Count > 0 Then s = s & IIf(s = "", "", vbCrLf) & NowLine(UCase$(Left$(Timing(t), 1)) & Mid$(Timing(t), 2), vals)
        End If
        For i = 1 To MAX_NAMES
            n = 0
            If i <= xCount And colX(t) > 0 Then
                For Each r In xRows
                    If HasOn(ParamTable.CellText(ws.Cells(r, colX(t))), xNames(i)) Then n = n + 1
                Next
            End If
            If n = 0 Then
                xStart(i, t) = False
            ElseIf n = xRows.Count Then
                xStart(i, t) = True
            Else
                xStart(i, t) = Null
            End If
            Box(i, t).TripleState = IsNull(xStart(i, t))     ' so a click can put it back to grey
            Box(i, t).Value = xStart(i, t)
        Next
    Next
    lblXNow.Caption = s
    lblXHead.Caption = "Coolant before / with / after the move  -  machines with X-style coolant" & _
                       IIf(xRows.Count > 1, "  -  " & xRows.Count & " ops", "")
    For i = 1 To xCount
        Me.Controls("lblX" & i).Caption = xNames(i) & IIf(xHits(i) < xRows.Count, "  *", "")
        If xHits(i) < xRows.Count Then part = True
    Next
    s = ""
    If part Then s = "* not on every one of these ops' machines - an op whose machine lacks it is left as it is, and listed."
    If xMore <> "" Then s = s & IIf(s = "", "", vbCrLf) & "More coolants than shown here (" & _
                           Replace(Mid$(xMore, 2, Len(xMore) - 2), "|", ", ") & ") - set those in the cells."
    lblXNote.Caption = s

    s = ""
    If noRows.Count > 0 And opRows.Count > 0 Then
        s = ParamTable.OpNames(noRows) & ":  no coolant on this kind of op" & IIf(noRows.Count < opRows.Count, " - left alone.", ".")
    End If
    lblNone.Caption = s
    busy = False
    Layout
    UpdateAll
End Sub

' How an op's machine sets coolant, by which of its cells take input: "v9" the one
' setting, "x" before / with / after by the machine's names, "" none at all.
Private Function Style(ByVal r As Long) As String
    Dim ws As Worksheet, t As Long, opts As Variant, i As Long
    Set ws = ParamTable.MainSheet
    If colV9 > 0 Then
        If ParamTable.TakesInput(ws.Cells(r, colV9)) And ParamTable.ListOf(ws.Cells(r, colV9)) <> "" Then Style = "v9": Exit Function
    End If
    For t = 0 To 2
        If colX(t) > 0 Then
            If ParamTable.TakesInput(ws.Cells(r, colX(t))) Then
                ' A V9 machine's leftover X-style entries offer only "none": not X-style.
                opts = ParamTable.CoolantChoices(ws.Cells(r, colX(t)))
                For i = 0 To UBound(opts)
                    If LCase$(Trim$(opts(i))) <> "none" And Trim$(opts(i)) <> "" Then Style = "x": Exit Function
                Next
            End If
        End If
    Next
End Function

' An X-style op's machine coolants, added to those offered - each counted once per op.
Private Sub AddNames(ByVal r As Long)
    Dim ws As Worksheet, t As Long, opts As Variant, i As Long, k As Long, seen As String, nm As String
    Set ws = ParamTable.MainSheet
    For t = 0 To 2
        If colX(t) > 0 Then
            opts = ParamTable.CoolantChoices(ws.Cells(r, colX(t)))
            For i = 0 To UBound(opts)
                nm = Trim$(opts(i))
                If LCase$(nm) <> "none" And nm <> "" And InStr(seen, "|" & LCase$(nm) & "|") = 0 Then
                    seen = seen & "|" & LCase$(nm) & "|"
                    For k = 1 To xCount
                        If LCase$(xNames(k)) = LCase$(nm) Then Exit For
                    Next
                    If k > xCount Then
                        If xCount < MAX_NAMES Then
                            xCount = xCount + 1
                            xNames(xCount) = nm
                            xHits(xCount) = 0
                        Else
                            If InStr(LCase$(xMore), "|" & LCase$(nm) & "|") = 0 Then xMore = IIf(xMore = "", "|", xMore) & nm & "|"
                            k = 0
                        End If
                    End If
                    If k > 0 Then xHits(k) = xHits(k) + 1
                End If
            Next
        End If
    Next
End Sub

' The V9 choices from the V9 ops' own list (Off, Flood, Mist, Thru-tool).
Private Sub V9Choices()
    Dim its As Variant, i As Long
    If v9Rows.Count = 0 Then Exit Sub
    its = Split(ParamTable.ListOf(ParamTable.MainSheet.Cells(v9Rows(1), colV9)), ",")
    For i = 1 To 4
        If i - 1 <= UBound(its) Then Me.Controls("optV" & i).Caption = Trim$(its(i - 1))
        Me.Controls("optV" & i).Visible = (i - 1 <= UBound(its))
    Next
End Sub

' "Now: Flood", "With the move: none  (all 3 ops)", "Before the move: Mixed: Flood x2, none x1".
Private Function NowLine(ByVal label As String, ByVal vals As Collection) As String
    Dim c As String, v As String
    If vals.Count = 0 Then Exit Function
    c = ParamTable.CountsText(vals)
    If InStr(c, ",") = 0 Then
        v = Trim$(vals(1))
        If v = "" Then v = "(blank)"
        NowLine = label & ":  " & v & IIf(vals.Count > 1, IIf(vals.Count = 2, "   (both ops)", "   (all " & vals.Count & " ops)"), "")
    Else
        NowLine = label & ":  Mixed: " & c
    End If
End Function

' Whether a cell's coolants ("Flood + Mist off") turn this one on.
Private Function HasOn(ByVal v As String, ByVal nm As String) As Boolean
    Dim tok As Variant
    For Each tok In Split(v, "+")
        If LCase$(Trim$(tok)) = LCase$(nm) Then HasOn = True: Exit Function
    Next
End Function

' The sheet shows what the window is on: the coolant columns, on its ops.
Private Sub Focus()
    If opRows.Count = 0 Then Exit Sub
    Panel.ShowColumns "coolant|coolant_before|coolant_with|coolant_after", opRows
End Sub

' ---------------------------------------------------------------- changes

Private Function V9Pick() As Long
    Dim i As Long
    For i = 1 To 4
        If Me.Controls("optV" & i).Value = True Then V9Pick = i
    Next
End Function

Private Function V9Dirty() As Boolean
    If v9Rows.Count > 0 Then V9Dirty = (V9Pick() > 0 And V9Pick() <> v9Start)
End Function

Private Function SameState(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        SameState = (IsNull(a) And IsNull(b))
    Else
        SameState = (CBool(a) = CBool(b))
    End If
End Function

' A timing whose boxes were changed - only those are written.
Private Function XDirty(ByVal t As Long) As Boolean
    Dim i As Long
    If xRows.Count = 0 Or colX(t) = 0 Then Exit Function
    For i = 1 To xCount
        If Not SameState(Box(i, t).Value, xStart(i, t)) Then XDirty = True: Exit Function
    Next
End Function

' Changes here that are not applied yet.
Public Function Dirty() As Boolean
    Dirty = V9Dirty() Or XDirty(0) Or XDirty(1) Or XDirty(2)
End Function

' One op's coolants at a timing after Apply: a ticked box on, an unticked one not, a grey
' one as the op has it; what has no box ("Mist off", a coolant past the twelve shown) is
' kept - an "off" only goes when that coolant is now on.
Private Function NewX(ByVal cur As String, ByVal t As Long) As String
    Dim i As Long, s As String, st As Variant, tok As Variant, nm As String, boxed As Boolean
    For i = 1 To xCount
        st = Box(i, t).Value
        If IsNull(st) Then
            If HasOn(cur, xNames(i)) Then s = s & " + " & xNames(i)
        ElseIf st Then
            s = s & " + " & xNames(i)
        End If
    Next
    For Each tok In Split(cur, "+")
        nm = Trim$(tok)
        If nm <> "" And LCase$(nm) <> "none" Then
            boxed = False
            For i = 1 To xCount
                If LCase$(nm) = LCase$(xNames(i)) Then boxed = True
                If LCase$(nm) = LCase$(xNames(i)) & " off" And HasOn(s, xNames(i)) Then boxed = True
            Next
            If Not boxed Then s = s & " + " & nm
        End If
    Next
    If s = "" Then NewX = "none" Else NewX = Mid$(s, 4)
End Function

' Buttons and the line at the bottom, after anything changed.
Private Sub UpdateAll()
    Dim s As String, warn As Boolean, d As Boolean
    d = Dirty()
    btnApply.Enabled = d
    btnRevert.Enabled = d Or Not pending Is Nothing
    btnPrev.Enabled = opRows.Count > 0
    btnNext.Enabled = opRows.Count > 0
    If opRows.Count = 0 Then
        s = ""
    ElseIf Not pending Is Nothing Then
        s = "Not applied yet - Apply, or Revert to follow the sheet again."
        warn = True
    ElseIf result <> "" Then
        s = result
    ElseIf d Then
        s = note & IIf(note = "", "", vbCrLf) & "Not applied yet - Apply writes it to the sheet (Ctrl+Z takes it back)."
    ElseIf v9Rows.Count + xRows.Count = 0 Then
        s = "Nothing to set - " & IIf(opRows.Count = 1, "this op has", "these ops have") & " no coolant."
    Else
        s = note
    End If
    lblStatus.Caption = s
    lblStatus.ForeColor = IIf(warn, RGB(180, 35, 24), RGB(33, 37, 41))
End Sub

Private Sub Changed()
    If busy Then Exit Sub
    result = ""
    UpdateAll
End Sub

Private Sub optV1_Click(): Changed: End Sub
Private Sub optV2_Click(): Changed: End Sub
Private Sub optV3_Click(): Changed: End Sub
Private Sub optV4_Click(): Changed: End Sub
Private Sub chkB1_Change(): Changed: End Sub
Private Sub chkB2_Change(): Changed: End Sub
Private Sub chkB3_Change(): Changed: End Sub
Private Sub chkB4_Change(): Changed: End Sub
Private Sub chkB5_Change(): Changed: End Sub
Private Sub chkB6_Change(): Changed: End Sub
Private Sub chkB7_Change(): Changed: End Sub
Private Sub chkB8_Change(): Changed: End Sub
Private Sub chkB9_Change(): Changed: End Sub
Private Sub chkB10_Change(): Changed: End Sub
Private Sub chkB11_Change(): Changed: End Sub
Private Sub chkB12_Change(): Changed: End Sub
Private Sub chkW1_Change(): Changed: End Sub
Private Sub chkW2_Change(): Changed: End Sub
Private Sub chkW3_Change(): Changed: End Sub
Private Sub chkW4_Change(): Changed: End Sub
Private Sub chkW5_Change(): Changed: End Sub
Private Sub chkW6_Change(): Changed: End Sub
Private Sub chkW7_Change(): Changed: End Sub
Private Sub chkW8_Change(): Changed: End Sub
Private Sub chkW9_Change(): Changed: End Sub
Private Sub chkW10_Change(): Changed: End Sub
Private Sub chkW11_Change(): Changed: End Sub
Private Sub chkW12_Change(): Changed: End Sub
Private Sub chkA1_Change(): Changed: End Sub
Private Sub chkA2_Change(): Changed: End Sub
Private Sub chkA3_Change(): Changed: End Sub
Private Sub chkA4_Change(): Changed: End Sub
Private Sub chkA5_Change(): Changed: End Sub
Private Sub chkA6_Change(): Changed: End Sub
Private Sub chkA7_Change(): Changed: End Sub
Private Sub chkA8_Change(): Changed: End Sub
Private Sub chkA9_Change(): Changed: End Sub
Private Sub chkA10_Change(): Changed: End Sub
Private Sub chkA11_Change(): Changed: End Sub
Private Sub chkA12_Change(): Changed: End Sub

' ---------------------------------------------------------------- buttons

' One op up or down the sheet (hidden rows skipped); the sheet goes with it.
Private Sub StepOp(ByVal way As Long)
    Dim r As Long, one As Collection
    If opRows.Count = 0 Then Exit Sub
    If Dirty() Then
        note = "Apply or Revert first - then move to another op."
        UpdateAll
        Exit Sub
    End If
    r = ParamTable.NextOpRow(IIf(way < 0, opRows(1), opRows(opRows.Count)), way)
    If r = 0 Then
        note = IIf(way < 0, "No op above this one.", "No op below this one.")
        UpdateAll
        Exit Sub
    End If
    Set one = New Collection
    one.Add r
    LoadRows one
    Focus
End Sub

Private Sub btnPrev_Click()
    StepOp -1
End Sub

Private Sub btnNext_Click()
    StepOp 1
End Sub

' Throw away what was changed: the ops as the sheet has them - the ones it moved to, if it did.
Private Sub btnRevert_Click()
    If pending Is Nothing Then LoadRows opRows Else LoadRows pending
End Sub

Private Sub btnApply_Click()
    If btnApply.Enabled Then DoApply
End Sub

' Write the parts changed - every cell checked against its own machine - as one Undo;
' then read the sheet back and say here what happened.
Private Sub DoApply()
    Dim ws As Worksheet, r As Variant, t As Long, c As Range, done As Collection, bad As Collection
    Dim why As String, s As String, v As String, nv As String, w As String
    If Not Dirty() Then Exit Sub
    Set ws = ParamTable.MainSheet
    Set done = New Collection
    Set bad = New Collection
    Panel.BeginEdit "Coolant"
    On Error GoTo Fail
    If V9Dirty() Then
        v = Me.Controls("optV" & V9Pick()).Caption
        For Each r In v9Rows
            Set c = ws.Cells(r, colV9)
            If LCase$(Trim$(ParamTable.CellText(c))) <> LCase$(v) Then
                If ParamTable.TryWrite(c, v) Then
                    ParamTable.CountOp done, CLng(r)
                Else
                    bad.Add "op " & ws.Cells(r, 1).Value & ": this machine has no " & v
                End If
            End If
        Next
    End If
    For t = 0 To 2
        If XDirty(t) Then
            For Each r In xRows
                Set c = ws.Cells(r, colX(t))
                If ParamTable.TakesInput(c) Then
                    nv = NewX(ParamTable.CellText(c), t)
                    If LCase$(nv) <> LCase$(Trim$(ParamTable.CellText(c))) Then
                        w = ParamTable.WriteCoolant(c, nv)
                        If w = "" Then
                            ParamTable.CountOp done, CLng(r)
                        Else
                            bad.Add "op " & ws.Cells(r, 1).Value & " (" & Timing(t) & "): " & w
                        End If
                    End If
                End If
            Next
        End If
    Next
    GoTo Finish
Fail:
    why = Err.Description
Finish:
    On Error GoTo 0
    Panel.EndEdit
    s = ParamTable.ApplyResult(done, bad, why)
    ' Read back - the ops the sheet went to meanwhile, if it did - and the other windows too.
    If Not pending Is Nothing Then Set opRows = pending
    LoadRows opRows
    Panel.RefreshWindows
    result = s
    UpdateAll
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp "coolant"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' Closing (Close, Esc or the X): the sheet stops telling it about the selection.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Panel.Closed Me
    ParamTable.WindowClosed Me
End Sub

' ---------------------------------------------------------------- for the checks

' One step: "v9=Flood" picks a V9 coolant, "with/Flood=1" ticks a box (0 unticks, ? greys
' it), "!button" presses it, "follow=A5:A6" is the sheet moving to those ops, "open=A5"
' the ribbon pressed there, "refresh" another window's Apply.
Public Sub Act(ByVal stepText As String)
    Dim p As Long, k As String, v As String, i As Long, t As Long, parts() As String
    If Left$(stepText, 1) = "!" Then
        Select Case Mid$(stepText, 2)
            Case "btnApply": btnApply_Click
            Case "btnRevert": btnRevert_Click
            Case "btnPrev": btnPrev_Click
            Case "btnNext": btnNext_Click
        End Select
        Exit Sub
    End If
    If stepText = "refresh" Then Refresh: Exit Sub
    p = InStr(stepText, "=")
    If p = 0 Then Exit Sub
    k = Left$(stepText, p - 1)
    v = Mid$(stepText, p + 1)
    Select Case k
    Case "v9"
        For i = 1 To 4
            If LCase$(Me.Controls("optV" & i).Caption) = LCase$(v) Then Me.Controls("optV" & i).Value = True
        Next
    Case "follow"
        FollowRows Panel.OpRows(ParamTable.MainSheet.Range(v))
    Case "open"
        OpenOn Panel.OpRows(ParamTable.MainSheet.Range(v))
    Case Else
        parts = Split(k, "/")
        If UBound(parts) <> 1 Then Exit Sub
        t = (InStr("bwa", LCase$(Left$(parts(0), 1)))) - 1
        If t < 0 Then Exit Sub
        For i = 1 To xCount
            If LCase$(xNames(i)) = LCase$(parts(1)) Then
                If v = "?" Then
                    Box(i, t).TripleState = True
                    Box(i, t).Value = Null
                Else
                    Box(i, t).Value = (v = "1")
                End If
            End If
        Next
    End Select
End Sub

' What the window shows: "parts" how many ops each part has ("v9 2, x 2, none 0"), "v9"
' the V9 choice picked ("" = none), "grid" the X-style boxes ("Flood:-x-" - ticked x, not
' -, grey ?, before / with / after); else a label's words, a button "on" / "off". Line
' breaks read as " / ".
Public Function Field(ByVal name As String) As String
    Dim c As Object, i As Long, t As Long, s As String, st As Variant
    Select Case name
    Case "parts"
        s = "v9 " & v9Rows.Count & ", x " & xRows.Count & ", none " & noRows.Count
    Case "v9"
        If V9Pick() > 0 Then s = Me.Controls("optV" & V9Pick()).Caption
    Case "grid"
        For i = 1 To xCount
            s = s & IIf(s = "", "", ", ") & Me.Controls("lblX" & i).Caption & ":"
            For t = 0 To 2
                If colX(t) > 0 Then
                    st = Box(i, t).Value
                    If IsNull(st) Then s = s & "?" Else s = s & IIf(st, "x", "-")
                End If
            Next
        Next
    Case Else
        Set c = Me.Controls(name)
        Select Case TypeName(c)
            Case "CommandButton": s = IIf(c.Enabled, "on", "off")
            Case "Label": s = c.Caption
            Case Else: If IsNull(c.Value) Then s = "Null" Else s = CStr(c.Value)
        End Select
    End Select
    Field = Replace(s, vbCrLf, " / ")
End Function
