' Parameter Table Tool - the Comments & text window (a UserForm named TextEditor): the text
' on the selected ops, a tab for each kind - the op's comment, the tool inspection comment
' and a manual entry's text - each with a live count against its own limit.
'
' It stays open beside the sheet (modeless) and follows the ops selected (Panel.bas). With
' several ops, a comment typed goes on all of them ("Mixed" when they differ now); a manual
' entry's text is edited one op at a time (the op list on its tab), each op's typing kept
' until Apply. Apply writes every cell through its own rule (ParamTable.TryWrite), as one
' Undo, and says in the window what it did.
'
' tools\build_vba.ps1 creates the controls - lblOps, btnPrev, btnNext, btnHelp, the tabs
' mpg, lblStatus, btnRevert, btnApply, btnClose, and on each tab's page its own boxes - and
' puts this code behind them. Layout and wording are all here.
Option Explicit

Private Const PAD As Single = 6
Private Const PG_CMT As Long = 0
Private Const PG_INSP As Long = 1
Private Const PG_MAN As Long = 2

Private opRows As Collection                ' the ops the window is on (sheet rows)
Private pending As Collection               ' where the sheet went while typing was not applied
Private busy As Boolean                     ' boxes filled by code, not typed
Private note As String                      ' a line about these ops (a tab they do not have ...)
Private result As String                    ' what the last Apply did
Private colCmt As Long, colInsp As Long, colInspOn As Long, colMan As Long, colCode As Long

' The two comments: the rows that take one (the sheet's own rule), the text there now (""
' when they differ), and the tightest limit among them.
Private cmtRows As Collection, cmtNow As String, cmtMixed As Boolean, cmtMax As Long
Private inspRows As Collection, inspNow As String, inspMixed As Boolean, inspMax As Long
Private inspOff As Collection               ' of those, stops that show no comment (insp_comment_on off)
Private inspOnTouched As Boolean            ' "switch it on" ticked or unticked by the person

' Manual entries, one op at a time: the text there now, as typed (kept per op), the limit.
Private manRows As Collection, manNow() As String, manText() As String, manMax() As Long, manIdx As Long

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
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private Sub UserForm_Initialize()
    Dim t As Variant, i As Long
    Set opRows = New Collection
    Set cmtRows = New Collection
    Set inspRows = New Collection
    Set inspOff = New Collection
    Set manRows = New Collection
    Me.Caption = "Parameter Table - Comments & text"
    Me.StartUpPosition = 0                  ' placed beside the sheet, not over it (ShowOn)
    Me.Width = 540
    Me.Height = 440

    lblOps.Font.Size = 10
    lblOps.Font.Bold = True
    btnPrev.Caption = "< Prev"
    btnPrev.ControlTipText = "The op above on the sheet"
    btnNext.Caption = "Next >"
    btnNext.ControlTipText = "The op below on the sheet"
    btnHelp.Caption = "?"
    btnHelp.ControlTipText = "How this window works (the user manual)"
    For Each t In Array(btnPrev, btnNext, btnHelp)
        t.TabStop = False                   ' Enter is Apply, never a step to another op
        t.TakeFocusOnClick = False
    Next
    mpg.Pages(PG_CMT).Caption = "Op comment"
    mpg.Pages(PG_INSP).Caption = "Inspection comment"
    mpg.Pages(PG_MAN).Caption = "Manual text"

    For Each t In Array(txtComment, txtInsp)
        t.MultiLine = True                  ' a long comment wraps; Enter still applies
        t.WordWrap = True
        t.EnterKeyBehavior = False
        t.ScrollBars = fmScrollBarsVertical
        t.Font.Size = 11
    Next
    With txtManual
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsBoth
        .EnterKeyBehavior = True             ' Enter is a new line; Ctrl+Enter applies
        .TabKeyBehavior = False
        .Font.Name = "Consolas"
        .Font.Size = 11
    End With
    For Each t In Array(lblCmtHint, lblInspHint, lblQuickNote)
        t.Font.Size = 8
        t.ForeColor = RGB(91, 101, 115)
    Next
    For Each t In Array(lblCmtCount, lblInspCount, lblManCount, lblCmtNow, lblInspNow, lblStatus, lblQuick, lblManOp, lblManHint)
        t.Font.Size = 9
    Next
    lblCmtNow.WordWrap = True
    lblInspNow.WordWrap = True
    lblStatus.WordWrap = True

    ' The comments the shop uses most; the note says which ones this tool counts as flips.
    lblQuick.Caption = "Common comments - click one to use it:"
    For i = 1 To 4
        With Me.Controls("btnQ" & i)
            .Caption = Choose(i, "ROTATE INSERT", "CHANGE INSERT", "INDEX INSERT", "CHECK INSERT")
            .Font.Size = 8
            .TakeFocusOnClick = False
        End With
    Next
    lblQuickNote.Caption = "Stops whose comment says ROTATE / FLIP / CHANGE / INDEX are counted as insert flips by this tool."
    chkInspOn.Font.Size = 9
    chkInspOn.Visible = False

    lblManOp.Caption = "Op:"
    cboOp.Style = fmStyleDropDownList
    cboOp.Font.Size = 9

    btnRevert.Caption = "Revert"
    btnRevert.ControlTipText = "Throw away what was typed here - show what the sheet has"
    btnApply.Caption = "Apply"
    btnApply.Default = True                 ' Enter (Ctrl+Enter on manual text)
    btnApply.ControlTipText = "Write it to the sheet - Ctrl+Z takes it back"
    btnClose.Caption = "Close"
    btnClose.Cancel = True                  ' Esc
    Layout
    UpdateAll
End Sub

' Resizable: a plain UserForm has a fixed border; give it a sizing frame and maximise.
' Typing goes straight into the tab's text.
Private Sub UserForm_Activate()
    #If VBA7 Then
        Dim h As LongPtr, pid As Long
        h = FindWindowA("ThunderDFrame", Me.Caption)
        If h <> 0 Then GetWindowThreadProcessId h, pid        ' only this Excel's window, not another's of the same name
        If h <> 0 And pid = GetCurrentProcessId() Then
            SetWindowLongPtrA h, -16, GetWindowLongPtrA(h, -16) Or &H40000 Or &H10000   ' WS_THICKFRAME, WS_MAXIMIZEBOX
            DrawMenuBar h
        End If
    #End If
    On Error Resume Next
    Select Case mpg.Value
        Case PG_INSP: txtInsp.SetFocus
        Case PG_MAN: txtManual.SetFocus
        Case Else: txtComment.SetFocus
    End Select
End Sub

Private Sub UserForm_Resize()
    Layout
End Sub

Private Sub Layout()
    Dim w As Single, h As Single, pw As Single, ph As Single, bw As Single, bh As Single, x As Single, i As Long
    w = Me.InsideWidth: h = Me.InsideHeight
    If w < 400 Or h < 330 Then Exit Sub
    bw = 72: bh = 24
    btnHelp.Move w - PAD - 22, PAD, 22, 20
    btnNext.Move btnHelp.Left - 8 - 58, PAD, 58, 20
    btnPrev.Move btnNext.Left - 4 - 58, PAD, 58, 20
    lblOps.Move PAD, PAD + 3, btnPrev.Left - 2 * PAD, 16
    btnClose.Move w - PAD - bw, h - PAD - bh, bw, bh
    btnApply.Move btnClose.Left - 6 - bw, btnClose.Top, bw, bh
    btnRevert.Move btnApply.Left - 6 - bw, btnClose.Top, bw, bh
    lblStatus.Move PAD, btnClose.Top - 4 - 42, w - 2 * PAD, 42
    mpg.Move PAD, PAD + 26, w - 2 * PAD, lblStatus.Top - 4 - (PAD + 26)

    ' Inside a page: its own area below the tabs.
    pw = mpg.Width - 8: ph = mpg.Height - 26
    lblCmtHint.Move PAD, PAD, pw - 2 * PAD, 12
    txtComment.Move PAD, PAD + 14, pw - 2 * PAD, 44
    lblCmtCount.Move PAD, PAD + 62, pw - 2 * PAD, 14
    lblCmtNow.Move PAD, PAD + 82, pw - 2 * PAD, Application.Max(14, ph - PAD - 82 - PAD)

    lblInspHint.Move PAD, PAD, pw - 2 * PAD, 12
    txtInsp.Move PAD, PAD + 14, pw - 2 * PAD, 44
    lblInspCount.Move PAD, PAD + 62, pw - 2 * PAD, 14
    lblQuick.Move PAD, PAD + 82, pw - 2 * PAD, 14
    x = PAD
    For i = 1 To 4
        Me.Controls("btnQ" & i).Move x, PAD + 97, 100, 20
        x = x + 104
    Next
    lblQuickNote.Move PAD, PAD + 120, pw - 2 * PAD, 12
    chkInspOn.Move PAD, PAD + 138, pw - 2 * PAD, 18
    lblInspNow.Move PAD, PAD + 160, pw - 2 * PAD, Application.Max(14, ph - PAD - 160 - PAD)

    lblManOp.Move PAD, PAD + 3, 22, 14
    cboOp.Move PAD + 24, PAD, Application.Min(300, pw / 2), 20
    lblManHint.Move cboOp.Left + cboOp.Width + 8, PAD + 3, pw - (cboOp.Left + cboOp.Width + 8) - PAD, 14
    txtManual.Move PAD, PAD + 26, pw - 2 * PAD, ph - (PAD + 26) - 22
    lblManCount.Move PAD, ph - 18, pw - 2 * PAD, 14
End Sub

' ---------------------------------------------------------------- the ops it is on

' From the ribbon or a double-click: the window on these ops, on a tab ("comment",
' "insp", "manual"; "" = as it is), the sheet brought to that column - shown beside the
' sheet, or brought to the front when it already is.
Public Sub ShowOn(ByVal rows As Collection, Optional ByVal tabName As String = "")
    Dim placed As Boolean
    placed = Me.Visible
    OpenOn rows, tabName
    If Not placed Then
        Me.Left = Application.Left + Application.Width - Me.Width - 40
        Me.Top = Application.Top + 140
    End If
    Me.Show vbModeless
    Panel.Opened Me
End Sub

' The same without showing it. Typing not applied yet stays (the window says so) - it is
' never thrown away by a click elsewhere.
Public Sub OpenOn(ByVal rows As Collection, Optional ByVal tabName As String = "")
    If Dirty() Then
        If Panel.RowsKey(rows) <> Panel.RowsKey(opRows) Then Set pending = rows
        PickTab tabName
        UpdateAll
    Else
        LoadRows rows, tabName
        Focus
    End If
End Sub

' From the sheet (Panel.SelectionMoved): other ops picked - read them, unless typing here
' is not applied yet.
Public Sub FollowRows(ByVal rows As Collection)
    If Dirty() Then
        Set pending = rows
        UpdateAll
        Exit Sub
    End If
    LoadRows rows
End Sub

' After Undo or another window's Apply: the same ops read again (typing waiting is kept).
Public Sub Refresh()
    If Dirty() Then Exit Sub
    LoadRows opRows
End Sub

' Read the ops' text into the tabs. tabName picks the tab; "" keeps the one showing when
' these ops have it.
Public Sub LoadRows(ByVal rows As Collection, Optional ByVal tabName As String = "")
    Dim ws As Worksheet, r As Variant, i As Long, keepRow As Long
    Set ws = ParamTable.MainSheet
    If manIdx > 0 Then keepRow = manRows(manIdx)
    busy = True
    Set opRows = rows
    Set pending = Nothing
    note = ""
    result = ""
    colCmt = ParamTable.ColOf("comment")
    colInsp = ParamTable.ColOf("insp_comment")
    colInspOn = ParamTable.ColOf("insp_comment_on")
    colMan = ParamTable.ColOf("manual_text")
    colCode = ParamTable.ColOf("manual_gcode")

    If opRows.Count = 0 Then
        lblOps.Caption = "Select one or more ops (any cell in their rows)"
    ElseIf opRows.Count = 1 Then
        lblOps.Caption = Panel.OpLabel(opRows(1))
    Else
        lblOps.Caption = Panel.RowsText(opRows)
    End If

    ' The op comment.
    Set cmtRows = Taking(colCmt)
    ReadText colCmt, cmtRows, 119, cmtNow, cmtMixed, cmtMax
    txtComment.Text = cmtNow
    lblCmtHint.Caption = "The op's name in the Operation Manager - one line, up to " & cmtMax & " characters.          column: comment"
    lblCmtNow.Caption = NowText(cmtRows, cmtMixed, "comments", "no comment to change")

    ' The tool inspection comment - and the stops that show none (it does nothing there
    ' until it is switched on).
    Set inspRows = Taking(colInsp)
    ReadText colInsp, inspRows, 49, inspNow, inspMixed, inspMax
    txtInsp.Text = inspNow
    lblInspHint.Caption = "Shown at each tool inspection stop - one line, up to " & inspMax & " characters.          column: insp_comment"
    lblInspNow.Caption = NowText(inspRows, inspMixed, "inspection comments", "no tool inspection")
    Set inspOff = New Collection
    If colInspOn > 0 Then
        For Each r In inspRows
            If ParamTable.TakesInput(ws.Cells(r, colInspOn)) And ParamTable.CellText(ws.Cells(r, colInspOn)) <> "1" Then inspOff.Add r
        Next
    End If
    chkInspOn.Visible = inspOff.Count > 0
    chkInspOn.Value = False
    inspOnTouched = False
    If inspOff.Count > 0 Then
        chkInspOn.Caption = "Also switch the comment on at the stop - it is off on " & _
                            IIf(inspOff.Count = inspRows.Count And inspRows.Count > 1, "all of these ops", ParamTable.OpNames(inspOff)) & _
                            "          (insp_comment_on)"
    End If

    ' Manual entries, one at a time - the same op as before when it is still here.
    Set manRows = Taking(colMan)
    cboOp.Clear
    manIdx = 0
    If manRows.Count > 0 Then
        ReDim manNow(1 To manRows.Count): ReDim manText(1 To manRows.Count): ReDim manMax(1 To manRows.Count)
        For Each r In manRows
            i = i + 1
            manNow(i) = Lines(ParamTable.CellText(ws.Cells(r, colMan)))
            manText(i) = manNow(i)
            manMax(i) = ParamTable.TextLimit(ws.Cells(r, colMan))
            If manMax(i) = 0 Then manMax(i) = 3111      ' what the Manual Entry dialog holds
            cboOp.AddItem Panel.OpLabel(r)
            If r = keepRow Then manIdx = i
        Next
        If manIdx = 0 Then manIdx = 1
        cboOp.ListIndex = manIdx - 1
        txtManual.Text = manText(manIdx)
    Else
        txtManual.Text = ""
    End If
    cboOp.Enabled = manRows.Count > 1
    ManHint

    ' Only the tabs these ops have are on.
    mpg.Pages(PG_CMT).Enabled = cmtRows.Count > 0
    mpg.Pages(PG_INSP).Enabled = inspRows.Count > 0
    mpg.Pages(PG_MAN).Enabled = manRows.Count > 0
    If opRows.Count > 0 And cmtRows.Count + inspRows.Count + manRows.Count = 0 Then
        note = "Nothing to edit here - " & IIf(opRows.Count = 1, "this op takes", "these ops take") & _
               " no comment, inspection comment or manual text."
    End If
    PickTab tabName
    busy = False
    UpdateAll
End Sub

' The window's rows whose cell in a column takes typing (the sheet's own rule says).
Private Function Taking(ByVal col As Long) As Collection
    Dim r As Variant, out As Collection
    Set out = New Collection
    If col > 0 Then
        For Each r In opRows
            If ParamTable.TakesInput(ParamTable.MainSheet.Cells(r, col)) Then out.Add r
        Next
    End If
    Set Taking = out
End Function

' The window's rows that are not in rs.
Private Function RowsBut(ByVal rs As Collection) As Collection
    Dim r As Variant, x As Variant, out As Collection, inIt As Boolean
    Set out = New Collection
    For Each r In opRows
        inIt = False
        For Each x In rs
            If x = r Then inIt = True: Exit For
        Next
        If Not inIt Then out.Add r
    Next
    Set RowsBut = out
End Function

' A comment's text on the rows now: one text, or "" and mixed when they differ; and the
' tightest limit among their cells (fallback when none says).
Private Sub ReadText(ByVal col As Long, ByVal rs As Collection, ByVal fallback As Long, _
                     ByRef txt As String, ByRef mixed As Boolean, ByRef most As Long)
    Dim ws As Worksheet, r As Variant, v As String, n As Long, k As Long
    Set ws = ParamTable.MainSheet
    txt = "": mixed = False: most = 0
    For Each r In rs
        v = OneLine(ParamTable.CellText(ws.Cells(r, col)))
        n = n + 1
        If n = 1 Then
            txt = v
        ElseIf v <> txt Then
            mixed = True
        End If
        k = ParamTable.TextLimit(ws.Cells(r, col))
        If k > 0 And (most = 0 Or k < most) Then most = k
    Next
    If mixed Then txt = ""
    If most = 0 Then most = fallback
End Sub

' What a comment tab says about the ops: one text for all of them, or Mixed - and the
' ops left alone because they have none.
Private Function NowText(ByVal rs As Collection, ByVal mixed As Boolean, ByVal what As String, ByVal noneWhy As String) As String
    Dim s As String, rest As Collection
    If rs.Count > 1 Then
        If mixed Then
            s = "Mixed - these " & rs.Count & " ops have different " & what & " now. What you type goes on all of them."
        Else
            s = "The same on " & IIf(rs.Count = 2, "both", "all " & rs.Count) & " ops - what you type goes on all of them."
        End If
    End If
    Set rest = RowsBut(rs)
    If rest.Count > 0 And rs.Count > 0 Then s = s & IIf(s = "", "", vbCrLf) & ParamTable.OpNames(rest) & ":  " & noneWhy & " - left alone."
    NowText = s
End Function

' Show a tab: the one asked for when these ops have it (else say why not), or the one
' showing, or the first they have.
Private Sub PickTab(ByVal tabName As String)
    Dim want As Long, i As Long, was As Boolean
    was = busy
    busy = True
    want = -1
    Select Case tabName
        Case "comment": want = PG_CMT
        Case "insp": want = PG_INSP
        Case "manual": want = PG_MAN
    End Select
    If want >= 0 And opRows.Count > 0 Then
        If Not mpg.Pages(want).Enabled Then
            note = WhyOff(want)
            want = -1
        End If
    End If
    If want < 0 And mpg.Value >= 0 Then
        If mpg.Pages(mpg.Value).Enabled Then want = mpg.Value
    End If
    If want < 0 Then
        For i = 0 To 2
            If mpg.Pages(i).Enabled Then want = i: Exit For
        Next
    End If
    If want >= 0 Then mpg.Value = want
    busy = was
End Sub

' Why a tab is off for these ops.
Private Function WhyOff(ByVal pg As Long) As String
    Dim one As Boolean, who As String
    one = (opRows.Count = 1)
    If one Then who = "op " & ParamTable.MainSheet.Cells(opRows(1), 1).Value
    Select Case pg
    Case PG_MAN
        WhyOff = IIf(one, who & " is not a manual entry - it has no manual text.", "None of these ops is a manual entry - no manual text.")
    Case PG_INSP
        WhyOff = IIf(one, who & " has no tool inspection comment (no tool inspection on this kind of op).", _
                      "None of these ops has a tool inspection comment.")
    Case Else
        WhyOff = IIf(one, who & "'s comment cannot be changed here.", "These ops' comments cannot be changed here.")
    End Select
End Function

' Above the manual text: how the op puts it out (manual_gcode).
Private Sub ManHint()
    Dim g As String
    If manIdx > 0 And colCode > 0 Then g = ParamTable.CellText(ParamTable.MainSheet.Cells(manRows(manIdx), colCode))
    Select Case g
        Case "1006": lblManHint.Caption = "Output as CODE.     Ctrl+Enter = Apply"
        Case "1005": lblManHint.Caption = "Output as a COMMENT.     Ctrl+Enter = Apply"
        Case Else: lblManHint.Caption = "Ctrl+Enter = Apply"
    End Select
End Sub

' The sheet shows what the window is on: the tab's column, on its ops.
Private Sub Focus()
    Dim one As Collection
    If opRows.Count = 0 Then Exit Sub
    Select Case mpg.Value
    Case PG_INSP
        Panel.ShowColumns "insp_comment", opRows
    Case PG_MAN
        If manIdx > 0 Then
            Set one = New Collection
            one.Add manRows(manIdx)
            Panel.ShowColumns "manual_text", one
        End If
    Case Else
        Panel.ShowColumns "comment", opRows
    End Select
End Sub

' ---------------------------------------------------------------- the text

' A comment is one line: line breaks become spaces.
Private Function OneLine(ByVal s As String) As String
    OneLine = Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function

' Every line break as CR LF - what Mastercam's manual entry holds.
Private Function Lines(ByVal s As String) As String
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    Lines = Replace(s, vbLf, vbCrLf)
End Function

Private Function CmtText() As String
    CmtText = OneLine(txtComment.Text)
End Function

Private Function InspText() As String
    InspText = OneLine(txtInsp.Text)
End Function

Private Function CmtDirty() As Boolean
    If cmtRows.Count > 0 Then CmtDirty = (CmtText() <> cmtNow)
End Function

Private Function InspDirty() As Boolean
    If inspRows.Count > 0 Then InspDirty = (InspText() <> inspNow)
End Function

' Switching the stops' comment on is a change of its own.
Private Function InspOnWanted() As Boolean
    If chkInspOn.Visible Then
        If Not IsNull(chkInspOn.Value) Then InspOnWanted = chkInspOn.Value
    End If
End Function

Private Function ManDirty() As Boolean
    Dim i As Long
    For i = 1 To manRows.Count
        If manText(i) <> manNow(i) Then ManDirty = True: Exit Function
    Next
End Function

' Typing here that is not applied yet.
Public Function Dirty() As Boolean
    Dirty = CmtDirty() Or InspDirty() Or InspOnWanted() Or ManDirty()
End Function

' What is typed and past its limit ("" = nothing) - the load would refuse it, so Apply
' stays off.
Private Function OverLimit() As String
    Dim i As Long
    If CmtDirty() And Len(CmtText()) > cmtMax Then OverLimit = "The op comment is longer than " & cmtMax & " characters."
    If InspDirty() And Len(InspText()) > inspMax Then OverLimit = "The inspection comment is longer than " & inspMax & " characters."
    For i = 1 To manRows.Count
        If manText(i) <> manNow(i) And Len(manText(i)) > manMax(i) Then
            OverLimit = "The manual text of op " & ParamTable.MainSheet.Cells(manRows(i), 1).Value & " is longer than " & _
                        Format$(manMax(i), "#,##0") & " characters."
        End If
    Next
End Function

' "12 / 119 characters" (and ", 18 lines" for manual text) - red past the limit.
Private Sub CountLine(ByVal lbl As MSForms.Label, ByVal s As String, ByVal most As Long, ByVal withLines As Boolean)
    Dim n As Long, ln As Long
    n = Len(s)
    lbl.Caption = Format$(n, "#,##0") & " / " & Format$(most, "#,##0") & " characters"
    If withLines Then
        If n > 0 Then ln = UBound(Split(s, vbLf)) + 1
        lbl.Caption = lbl.Caption & ",  " & ln & IIf(ln = 1, " line", " lines")
    End If
    If n > most Then
        lbl.ForeColor = RGB(180, 35, 24)
        lbl.Caption = lbl.Caption & "  -  over the limit by " & Format$(n - most, "#,##0")
    Else
        lbl.ForeColor = RGB(91, 101, 115)
    End If
End Sub

' Counts, buttons and the line at the bottom, after anything changed.
Private Sub UpdateAll()
    Dim over As String, s As String, warn As Boolean, d As Boolean
    CountLine lblCmtCount, CmtText(), cmtMax, False
    CountLine lblInspCount, InspText(), inspMax, False
    If manIdx > 0 Then CountLine lblManCount, manText(manIdx), manMax(manIdx), True Else lblManCount.Caption = ""
    over = OverLimit()
    d = Dirty()
    btnApply.Enabled = d And over = ""
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
    ElseIf over <> "" Then
        s = over & "  Shorten it to apply."
        warn = True
    ElseIf d Then
        s = note & IIf(note = "", "", vbCrLf) & "Not applied yet - Apply writes it to the sheet (Ctrl+Z takes it back)."
    Else
        s = note
    End If
    lblStatus.Caption = s
    lblStatus.ForeColor = IIf(warn, RGB(180, 35, 24), RGB(33, 37, 41))
End Sub

' ---------------------------------------------------------------- typing

Private Sub txtComment_Change()
    If busy Then Exit Sub
    result = ""
    UpdateAll
End Sub

' An inspection comment typed for stops that show none: switch it on too, unless the
' person has said otherwise.
Private Sub txtInsp_Change()
    If busy Then Exit Sub
    result = ""
    If chkInspOn.Visible And Not inspOnTouched Then
        busy = True
        chkInspOn.Value = InspDirty()
        busy = False
    End If
    UpdateAll
End Sub

Private Sub chkInspOn_Click()
    If busy Then Exit Sub
    inspOnTouched = True
    result = ""
    UpdateAll
End Sub

Private Sub txtManual_Change()
    If busy Or manIdx = 0 Then Exit Sub
    manText(manIdx) = Lines(txtManual.Text)
    result = ""
    UpdateAll
End Sub

' Enter is a new line in manual text, so Ctrl+Enter applies.
Private Sub txtManual_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And 2) <> 0 Then
        KeyCode = 0
        If btnApply.Enabled Then DoApply
    End If
End Sub

' Another manual entry of the ops: its own text (what was typed for the last one is kept).
Private Sub cboOp_Change()
    If busy Or cboOp.ListIndex < 0 Then Exit Sub
    manIdx = cboOp.ListIndex + 1
    busy = True
    txtManual.Text = manText(manIdx)
    busy = False
    ManHint
    Focus
    UpdateAll
End Sub

Private Sub mpg_Change()
    If busy Then Exit Sub
    Focus
    UpdateAll
End Sub

Private Sub QuickPick(ByVal s As String)
    txtInsp.Text = s
    On Error Resume Next
    If Me.Visible Then txtInsp.SetFocus
End Sub

Private Sub btnQ1_Click(): QuickPick btnQ1.Caption: End Sub
Private Sub btnQ2_Click(): QuickPick btnQ2.Caption: End Sub
Private Sub btnQ3_Click(): QuickPick btnQ3.Caption: End Sub
Private Sub btnQ4_Click(): QuickPick btnQ4.Caption: End Sub

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

' Throw away what was typed: the ops as the sheet has them - the ones it moved to, if it did.
Private Sub btnRevert_Click()
    If pending Is Nothing Then LoadRows opRows Else LoadRows pending
End Sub

Private Sub btnApply_Click()
    If btnApply.Enabled Then DoApply
End Sub

' Write what was typed - each cell through its own rule - as one Undo; then read the
' sheet back and say here what happened.
Private Sub DoApply()
    Dim ws As Worksheet, r As Variant, i As Long, done As Collection, bad As Collection, why As String, s As String
    If Not Dirty() Or OverLimit() <> "" Then Exit Sub
    Set ws = ParamTable.MainSheet
    Set done = New Collection
    Set bad = New Collection
    Panel.BeginEdit "Comments & text"
    On Error GoTo Fail
    If CmtDirty() Then
        For Each r In cmtRows
            Put1 ws.Cells(r, colCmt), CmtText(), "comment", done, bad
        Next
    End If
    If InspDirty() Then
        For Each r In inspRows
            Put1 ws.Cells(r, colInsp), InspText(), "inspection comment", done, bad
        Next
    End If
    If InspOnWanted() Then
        For Each r In inspOff
            Put1 ws.Cells(r, colInspOn), 1, "comment switch", done, bad
        Next
    End If
    For i = 1 To manRows.Count
        If manText(i) <> manNow(i) Then Put1 ws.Cells(manRows(i), colMan), manText(i), "manual text", done, bad
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

' One cell: left alone when it holds that already; else written through its own rule and
' its op counted - or refused, with why.
Private Sub Put1(ByVal c As Range, ByVal v As Variant, ByVal what As String, ByVal done As Collection, ByVal bad As Collection)
    Dim k As Long
    If ParamTable.CellText(c) = CStr(v) Then Exit Sub
    If ParamTable.TryWrite(c, v) Then
        ParamTable.CountOp done, c.Row
    Else
        k = ParamTable.TextLimit(c)
        bad.Add "op " & c.Worksheet.Cells(c.Row, 1).Value & " " & what & ": " & _
                IIf(k > 0 And Len(CStr(v)) > k, "longer than its " & k & " characters", "the cell does not take it")
    End If
End Sub

Private Sub btnHelp_Click()
    Panel.ShowHelp "text"
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

' One step: "box=text" types into a box, "!button" presses it, "tab=insp" shows a tab,
' "op=1" picks the manual op (from 0), "chkInspOn=1" ticks it, "follow=A5:A6" is the sheet
' moving to those ops, "open=A5" the ribbon pressed there, "refresh" another window's Apply.
Public Sub Act(ByVal stepText As String)
    Dim p As Long, k As String, v As String
    If Left$(stepText, 1) = "!" Then
        Select Case Mid$(stepText, 2)
            Case "btnApply": btnApply_Click
            Case "btnRevert": btnRevert_Click
            Case "btnPrev": btnPrev_Click
            Case "btnNext": btnNext_Click
            Case "btnQ1", "btnQ2", "btnQ3", "btnQ4": QuickPick Me.Controls(Mid$(stepText, 2)).Caption
        End Select
        Exit Sub
    End If
    If stepText = "refresh" Then Refresh: Exit Sub
    p = InStr(stepText, "=")
    If p = 0 Then Exit Sub
    k = Left$(stepText, p - 1)
    v = Mid$(stepText, p + 1)
    Select Case k
        Case "tab": mpg.Value = IIf(v = "insp", PG_INSP, IIf(v = "manual", PG_MAN, PG_CMT))
        Case "op": cboOp.ListIndex = CLng(v)
        Case "chkInspOn": chkInspOn.Value = (v = "1")
        Case "follow": FollowRows Panel.OpRows(ParamTable.MainSheet.Range(v))
        Case "open": OpenOn Panel.OpRows(ParamTable.MainSheet.Range(v))
        Case Else: Me.Controls(k).Text = v
    End Select
End Sub

' What a control shows: a label's words, a box's text, a button "on" / "off", a tick box
' True / False ("hidden" when not shown); "tabs" which tabs are on ("101"), "tab" the one
' showing. Line breaks read as " / ".
Public Function Field(ByVal name As String) As String
    Dim c As Object, i As Long, s As String
    Select Case name
    Case "tabs"
        For i = 0 To 2
            s = s & IIf(mpg.Pages(i).Enabled, "1", "0")
        Next
    Case "tab"
        If mpg.Value >= 0 Then s = Choose(mpg.Value + 1, "comment", "insp", "manual")
    Case Else
        Set c = Me.Controls(name)
        Select Case TypeName(c)
            Case "CommandButton": s = IIf(c.Enabled, "on", "off")
            Case "Label": s = c.Caption
            Case "CheckBox": If c.Visible Then s = CStr(c.Value) Else s = "hidden"
            Case Else: s = c.Text
        End Select
    End Select
    Field = Replace(s, vbCrLf, " / ")
End Function
