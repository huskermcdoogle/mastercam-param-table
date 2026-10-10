Attribute VB_Name = "Checks"
' Parameter Table Tool - the continuous-improvement tools on the ribbon:
'
'   PROGRAM CHECK   (Improve > Program check) the whole program looked over for what is
'                   worth a second look before it goes to the floor. One line per finding
'                   on a "Program check" sheet: Look at / FYI, the op (a link to it), the
'                   tool, what was found, why it matters, how to fix it - and an Ignore
'                   column. Most useful first; ignored ones greyed at the bottom.
'   SLOWEST OPS     (Find > Slowest ops) where the time is: the sheet's own filter shows
'                   only the ops that together make SLOW_SHARE of the estimated cycle time,
'                   and a "Slowest ops" sheet ranks them with links. Show all clears it.
'   CHANGE REPORT   (Undo and review > Change report) a page to print: before -> after for
'                   a part and a batch, then every change by op, with a Why to type in.
'
' NONE OF THEM EDITS THE MAIN SHEET - they read it (Slowest ops sets its filter, nothing
' else), so there is nothing to undo. Each sheet is made again on every click.
'
' THE RULES. A finding's key is "rule|subject" - an op ("op 7"), or a tool and what it
' carries ("T5|RPGV 1204"). The key is how a finding is known again, so an ignored one
' stays ignored; a finding no longer there simply stops showing.
'   needs-regen   Look at  the toolpath is out of date (needs_regen = yes)
'   feed-unit     Look at  a feed that looks like the other unit (ParamTable.CheckRow's test)
'   css-no-max    Look at  CSS with no max spindle speed (ParamTable.CheckRow's test)
'   edge-time     Look at  a tool's edge cuts far longer than its insert usually does: more
'                          than EDGE_OVER_USUAL x the insert's "Usual edge time" (Tools
'                          page); with none typed, more than EDGE_OVER_MEDIAN x the median
'                          of the other tools carrying it here, and EDGE_MIN_EXTRA longer
'   air           Look at  over AIR_PCT % of an op's cut time in air, and over AIR_MIN_SEC
'   mixed-speeds  FYI      one tool, one kind of op, speeds or feeds more than SPREAD_MAX apart
'   tool-changes  FYI      one tool put in more than once in a toolpath group
'   coolant-off   FYI      a cutting op with no coolant on
'   no-comment    FYI      an op with no comment
' A comment never decides a flip here, and tool inspection switched off is a choice - so
' neither is a finding on its own.
'
' IGNORE. The Ignore column (yes / blank) goes into the hidden "Ignored findings" sheet at
' the next check. A dump writes that sheet from the part's .ptconfig, and reads it - and
' this Ignore column - back from the saved workbook: a finding ignored once stays ignored
' on later checks AND later dumps.
'
' Plain text in the repo (vba\Checks.bas); tools\build_vba.ps1 compiles every vba\*.bas.
' tools\check_macros.ps1 drives the quiet Functions (ProgramCheck, SlowestOps,
' ChangeReport); only the ribbon Subs at the bottom show anything.
Option Explicit

Private Const TITLE As String = "Parameter Table"
Private Const HEADER_ROW As Long = 2
Private Const FIRST_ROW As Long = 3

Public Const CHECK_SHEET As String = "Program check"
Public Const IGNORE_SHEET As String = "Ignored findings"
Public Const SLOW_SHEET As String = "Slowest ops"
Public Const REPORT_SHEET As String = "Change report"

' ---- What counts - every threshold in one place.
Private Const EDGE_OVER_USUAL As Double = 1.25      ' an edge more than 25% over the insert's usual edge time
Private Const EDGE_OVER_MEDIAN As Double = 1.5      ' none typed: over 1.5 x the median of the insert's other tools
Private Const EDGE_MIN_EXTRA As Double = 120        ' ... and at least 2 minutes (seconds) longer than it
Private Const AIR_PCT As Double = 30                ' an op cutting air over 30% of its cut time
Private Const AIR_MIN_SEC As Double = 60            ' ... and over a minute of air per part
Private Const SPREAD_MAX As Double = 0.1            ' one tool, one kind of op: speeds or feeds over 10% apart
Private Const SLOW_SHARE As Double = 0.8            ' Slowest ops: the ops that make 80% of the cycle time
Private Const SLOW_NAMED As Long = 6                ' ... how many the status line names

' Severity, and the order rules come in (most useful first).
Private Const LOOK_AT As Long = 0
Private Const FYI As Long = 1

' A finding is an Array; these are its places.
Private Const F_SEV As Long = 0         ' LOOK_AT / FYI
Private Const F_RANK As Long = 1        ' the rule's place in the order
Private Const F_IMPACT As Long = 2      ' seconds (or a count) - bigger first within a rule
Private Const F_ROW As Long = 3         ' the main-sheet row it is about (the link)
Private Const F_COL As Long = 4         ' the column the link goes to
Private Const F_OPS As Long = 5         ' "op 7" / "ops 2, 9, 11"
Private Const F_TOOL As Long = 6        ' "T5  RPGV 1204"
Private Const F_TOOLNO As Long = 7      ' "5" - for the link to its Tools row
Private Const F_WHAT As Long = 8
Private Const F_WHY As Long = 9
Private Const F_FIX As Long = 10
Private Const F_KEY As Long = 11
Private Const F_IGN As Long = 12        ' ignored (set when sorted)

' The sheet as read for one command.
Private mMain As Worksheet
Private mV As Variant                   ' the main sheet's values, (row, column) from A1
Private mLast As Long, mLc As Long
Private mCol As Object                  ' column name -> column
Private mRowOf As Object                ' op_idn -> row
Private mD As Variant                   ' the Dumped sheet's values (Empty when there is none)
Private mDLast As Long
Private mDCol As Object                 ' its column name -> column
Private mDRowOf As Object               ' op_idn -> its Dumped row
Private mOrder As Collection            ' main-sheet rows in Operation Manager order (as dumped)

' The Tools page.
Private mToolRow As Object              ' tool number -> its Tools row
Private mInsertOf As Object             ' tool number -> the insert it carries (as named there)
Private mIns As Object                  ' LCase(insert) -> Array(name, edges, cost, parts per edge, usual s, row)

Private found As Collection
Private lastErr As String
Private mLine As String                 ' the check's count line

' ============================================================ reading the sheets

Private Function SheetNamed(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetNamed = ThisWorkbook.Worksheets(name)
End Function

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    SafeText = Trim$(CStr(v))
End Function

Private Function NumOf(ByVal v As Variant, ByRef x As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Or Not IsNumeric(v) Then Exit Function
    ElseIf Not IsNumeric(v) Then
        Exit Function
    End If
    x = CDbl(v)
    NumOf = True
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

' The main sheet (and Dumped) into memory - one read, then every rule works from it.
Private Function Snapshot() As Boolean
    Dim d As Worksheet, c As Long, r As Long, h As String, dc As Long
    Set mMain = ParamTable.MainSheet
    Application.Calculate
    mLast = ParamTable.LastRow
    mLc = ParamTable.LastCol
    Set mCol = NewDict()
    Set mRowOf = NewDict()
    Set mDCol = NewDict()
    Set mDRowOf = NewDict()
    Set mOrder = New Collection
    mD = Empty
    mDLast = 0
    If mLast < FIRST_ROW Or mLc < 2 Then Exit Function
    mV = mMain.Range(mMain.Cells(1, 1), mMain.Cells(mLast, mLc)).Value2
    For c = 1 To mLc
        h = SafeText(mV(HEADER_ROW, c))
        If h <> "" And Not mCol.Exists(h) Then mCol.Add h, c
    Next
    For r = FIRST_ROW To mLast
        h = SafeText(mV(r, 1))
        If h <> "" And Not mRowOf.Exists(h) Then mRowOf.Add h, r
    Next
    ' Dumped: the values as written, rows in Operation Manager order (it is never sorted).
    Set d = SheetNamed("Dumped")
    If Not d Is Nothing Then
        mDLast = d.Cells(d.Rows.Count, 1).End(xlUp).Row
        dc = d.Cells(HEADER_ROW, d.Columns.Count).End(xlToLeft).Column
        If mDLast >= FIRST_ROW And dc >= 2 Then
            mD = d.Range(d.Cells(1, 1), d.Cells(mDLast, dc)).Value2
            For c = 1 To dc
                h = SafeText(mD(HEADER_ROW, c))
                If h <> "" And Not mDCol.Exists(h) Then mDCol.Add h, c
            Next
        Else
            mDLast = 0
        End If
    End If
    ' The program order: Dumped's, then any row it does not have, in sheet order.
    Dim seen As Object
    Set seen = NewDict()
    For r = FIRST_ROW To mDLast
        h = SafeText(mD(r, 1))
        If h <> "" And Not mDRowOf.Exists(h) Then mDRowOf.Add h, r
        If mRowOf.Exists(h) And Not seen.Exists(h) Then
            seen.Add h, True
            mOrder.Add mRowOf(h)
        End If
    Next
    For r = FIRST_ROW To mLast
        h = SafeText(mV(r, 1))
        If h <> "" And Not seen.Exists(h) Then
            seen.Add h, True
            mOrder.Add r
        End If
    Next
    Snapshot = True
End Function

Private Function ColN(ByVal name As String) As Long
    If mCol.Exists(name) Then ColN = mCol(name)
End Function

' A main-sheet value by row and column name (Empty when the sheet has no such column).
Private Function Cv(ByVal r As Long, ByVal name As String) As Variant
    If mCol.Exists(name) Then Cv = mV(r, mCol(name))
End Function

Private Function Tx(ByVal r As Long, ByVal name As String) As String
    Tx = SafeText(Cv(r, name))
End Function

Private Function Nm(ByVal r As Long, ByVal name As String, ByRef x As Double) As Boolean
    Nm = NumOf(Cv(r, name), x)
End Function

' A Dumped value by its row and column name.
Private Function Dv(ByVal dr As Long, ByVal name As String) As Variant
    If dr < FIRST_ROW Or dr > mDLast Then Exit Function
    If mDCol.Exists(name) Then Dv = mD(dr, mDCol(name))
End Function

' The Dumped row of a main-sheet row (by op_idn), or 0.
Private Function DRow(ByVal r As Long) As Long
    Dim k As String
    k = SafeText(mV(r, 1))
    If mDRowOf.Exists(k) Then DRow = mDRowOf(k)
End Function

Private Function OpOf(ByVal r As Long) As String
    OpOf = SafeText(mV(r, 1))
End Function

' The row's tool number, "" for none (a manual entry, a transform).
Private Function ToolOf(ByVal r As Long) As String
    ToolOf = Tx(r, "tool")
    If ToolOf = "0" Then ToolOf = ""
End Function

' What the operator's seconds read as: 754 -> 12:34.
Private Function Tm(ByVal sec As Double) As String
    Tm = Planner.TimeText(sec)
End Function

' The column a link goes to: the first of the names the sheet has and shows - else
' op_idn, so a click never selects a cell out of sight in a folded group.
Private Function LinkCol(ByVal names As String) As Long
    Dim nm As Variant, c As Long
    LinkCol = 1
    For Each nm In Split(names, "|")
        c = ColN(CStr(nm))
        If c > 0 Then
            If Not mMain.Columns(c).Hidden Then LinkCol = c: Exit Function
        End If
    Next
End Function

' The batch quantity typed on the Summary (1 when there is none).
Public Function BatchQty() As Double
    Dim s As Worksheet, r As Long, x As Double
    BatchQty = 1
    Set s = SheetNamed("Summary")
    If s Is Nothing Then Exit Function
    For r = 1 To 300
        If LCase$(Left$(SafeText(s.Cells(r, 2).Value2), 14)) = "batch quantity" Then
            If NumOf(s.Cells(r, 3).Value2, x) Then
                If x >= 1 Then BatchQty = x
            End If
            Exit Function
        End If
    Next
End Function

' The Tools page: each tool's row and insert; the inserts table's edges, cost, parts per
' edge and usual edge time. Columns are found by their headings.
Private Sub ReadTools()
    Dim t As Worksheet, c As Long, lc As Long, r As Long, h As String, cTool As Long, cIns As Long, num As String
    Dim hr As Long, last As Long, cName As Long, cEdges As Long, cCost As Long, cPpe As Long, cUsual As Long
    Dim nm As String, edges As Double, cost As Double, ppe As Double, usual As Double, x As Double
    Set mToolRow = NewDict()
    Set mInsertOf = NewDict()
    Set mIns = NewDict()
    Set t = SheetNamed("Tools")
    If t Is Nothing Then Exit Sub
    lc = t.Cells(1, t.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        h = LCase$(SafeText(t.Cells(1, c).Value2))
        If h = "tool" And cTool = 0 Then cTool = c
        If h = "insert" And cIns = 0 Then cIns = c
    Next
    If cTool = 0 Then cTool = 1
    r = 2
    Do While SafeText(t.Cells(r, cTool).Value2) <> "" And r < 5000
        num = SafeText(t.Cells(r, cTool).Value2)
        If Not mToolRow.Exists(num) Then
            mToolRow.Add num, r
            If cIns > 0 Then mInsertOf.Add num, SafeText(t.Cells(r, cIns).Value2)
        End If
        r = r + 1
    Loop
    ' The inserts table: the row whose A says "Inserts", then a row per insert.
    last = t.Cells(t.Rows.Count, 2).End(xlUp).Row
    For hr = r To Application.Min(last, r + 20)
        If SafeText(t.Cells(hr, 1).Value2) = "Inserts" Then Exit For
    Next
    If hr > Application.Min(last, r + 20) Then Exit Sub
    lc = t.Cells(hr, t.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        h = LCase$(SafeText(t.Cells(hr, c).Value2))
        If h = "insert" Then cName = c
        If Left$(h, 5) = "edges" Then cEdges = c
        If h = "cost per insert" Then cCost = c
        If Left$(h, 14) = "parts per edge" Then cPpe = c
        If Left$(h, 15) = "usual edge time" Then cUsual = c
    Next
    If cName = 0 Then cName = 2
    For r = hr + 1 To last
        nm = SafeText(t.Cells(r, cName).Value2)
        If nm <> "" And Not mIns.Exists(LCase$(nm)) Then
            edges = 0: cost = -1: ppe = 0: usual = 0
            If cEdges > 0 Then
                If NumOf(t.Cells(r, cEdges).Value2, x) Then edges = x
            End If
            If cCost > 0 Then
                If NumOf(t.Cells(r, cCost).Value2, x) Then cost = x
            End If
            If cPpe > 0 Then
                If NumOf(t.Cells(r, cPpe).Value2, x) Then ppe = x
            End If
            If cUsual > 0 Then
                usual = Planner.SecondsOfText(t.Cells(r, cUsual).Value)
                If usual < 0 Then usual = 0
            End If
            mIns.Add LCase$(nm), Array(nm, edges, cost, ppe, usual, r)
        End If
    Next
End Sub

' "T5  RPGV 1204" - a tool and the insert it carries.
Private Function ToolText(ByVal tool As String) As String
    If tool = "" Then Exit Function
    ToolText = "T" & tool
    If mInsertOf.Exists(tool) Then
        If mInsertOf(tool) <> "" Then ToolText = ToolText & "  " & mInsertOf(tool)
    End If
End Function

' ============================================================ ignored findings

' The hidden list: key -> what it said.
Private Function IgnoredList() As Object
    Dim ws As Worksheet, r As Long, k As String, d As Object
    Set d = NewDict()
    Set IgnoredList = d
    Set ws = SheetNamed(IGNORE_SHEET)
    If ws Is Nothing Then Exit Function
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        k = SafeText(ws.Cells(r, 1).Value2)
        If k <> "" And Not d.Exists(k) Then d.Add k, SafeText(ws.Cells(r, 2).Value2)
    Next
End Function

Private Sub SaveIgnoredList(ByVal d As Object)
    Dim ws As Worksheet, k As Variant, r As Long, act As Worksheet
    Set ws = SheetNamed(IGNORE_SHEET)
    If ws Is Nothing Then
        If d.Count = 0 Then Exit Sub
        Set act = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = IGNORE_SHEET
        ws.Visible = xlSheetHidden
        On Error Resume Next
        act.Activate
        On Error GoTo 0
    End If
    ws.Cells.Clear
    ws.Columns(1).NumberFormat = "@"
    ws.Range("A1:B1").Value = Array("key", "finding")
    r = 2
    For Each k In d.Keys
        ws.Cells(r, 1).Value = CStr(k)
        ws.Cells(r, 2).Value = d(k)
        r = r + 1
    Next
End Sub

' The Program check sheet's Ignore column as it stands - a yes or a blank set since the
' last check - into the hidden list. Only rows that are findings (they carry a key).
Private Sub SyncIgnores()
    Dim ws As Worksheet, r As Long, c As Long, hr As Long, kc As Long, ic As Long, wc As Long, last As Long
    Dim k As String, d As Object, h As String
    Set ws = SheetNamed(CHECK_SHEET)
    If ws Is Nothing Then Exit Sub
    For r = 1 To 20
        kc = 0: ic = 0: wc = 0
        For c = 1 To 12
            h = LCase$(SafeText(ws.Cells(r, c).Value2))
            If h = "key" Then kc = c
            If h = "ignore" Then ic = c
            If h = "what was found" Then wc = c
        Next
        If kc > 0 And ic > 0 Then hr = r: Exit For
    Next
    If hr = 0 Then Exit Sub
    If wc = 0 Then wc = kc
    Set d = IgnoredList()
    last = ws.Cells(ws.Rows.Count, kc).End(xlUp).Row
    For r = hr + 1 To last
        k = SafeText(ws.Cells(r, kc).Value2)
        If k <> "" Then
            If LCase$(SafeText(ws.Cells(r, ic).Value2)) = "yes" Then
                If Not d.Exists(k) Then d.Add k, SafeText(ws.Cells(r, wc).Value2)
            ElseIf d.Exists(k) Then
                d.Remove k
            End If
        End If
    Next
    SaveIgnoredList d
End Sub

' For the checks: the ignored keys, one per line.
Public Function IgnoredKeys() As String
    Dim k As Variant
    For Each k In IgnoredList().Keys
        IgnoredKeys = IgnoredKeys & IIf(IgnoredKeys = "", "", vbLf) & k
    Next
End Function

' A key the .ptconfig can keep: no "=" (its separator), one line.
Private Function CleanKey(ByVal k As String) As String
    k = Replace(Replace(Replace(k, "=", "-"), vbCr, " "), vbLf, " ")
    CleanKey = Trim$(k)
End Function

' ============================================================ the rules

Private Sub AddFinding(ByVal sev As Long, ByVal rank As Long, ByVal impact As Double, ByVal r As Long, ByVal col As Long, _
                       ByVal ops As String, ByVal tool As String, ByVal what As String, ByVal why As String, _
                       ByVal howTo As String, ByVal key As String)
    found.Add Array(sev, rank, impact, r, col, ops, ToolText(tool), tool, what, why, howTo, CleanKey(key), False)
End Sub

' The rules one op at a time.
Private Sub OpRules(ByVal batch As Double)
    Dim r As Long, op As String, tool As String, kind As String, s As String, how As String, col As String
    Dim air As Double, cut As Double, sec As Double, x As Double, cutting As Boolean
    For r = FIRST_ROW To mLast
        op = OpOf(r)
        If op <> "" Then
            tool = ToolOf(r)
            kind = UCase$(Tx(r, "type"))
            cutting = (tool <> "" And kind <> "MANUAL" And kind <> "TRANSFORM")

            ' The toolpath is out of date.
            If LCase$(Tx(r, "needs_regen")) = "yes" Then
                AddFinding LOOK_AT, 1, 0, r, LinkCol("needs_regen"), "op " & op, tool, _
                    "op " & op & "'s toolpath is out of date - it needs regenerating.", _
                    "What posts is the old toolpath, not what its parameters say - and this sheet's time, flips and material for it come from the old toolpath too.", _
                    "Regenerate op " & op & " in Mastercam (load this sheet's changes first, if any), then dump again for fresh figures.", _
                    "needs-regen|op " & op
            End If

            ' ParamTable.CheckRow's own tests: a feed in the wrong unit, CSS with no cap.
            s = ParamTable.CheckRow(r)
            If InStr(s, "meant per rev?") > 0 Then
                AddFinding LOOK_AT, 2, 0, r, LinkCol("feed_mode|feed"), "op " & op, tool, _
                    "op " & op & "'s feed is " & Tx(r, "feed") & " per min - very slow for a per-minute feed. It looks like a per-rev figure.", _
                    "If the number is meant per rev, the unit is wrong and the post puts out a crawl - or the op really does crawl, wasting time every part.", _
                    "Check feed_mode on op " & op & ": set it to per rev if " & Tx(r, "feed") & " is a per-rev feed, or type the right per-minute feed.", _
                    "feed-unit|op " & op
            ElseIf InStr(s, "meant per min?") > 0 Then
                AddFinding LOOK_AT, 2, 0, r, LinkCol("feed_mode|feed"), "op " & op, tool, _
                    "op " & op & "'s feed is " & Tx(r, "feed") & " per rev - very high for a per-rev feed. It looks like a per-minute figure.", _
                    "Run per rev, a feed that size would crash the tool or break the insert on the first move.", _
                    "Check feed_mode on op " & op & ": set it to per min if " & Tx(r, "feed") & " is a per-minute feed, or type the right per-rev feed.", _
                    "feed-unit|op " & op
            End If
            If InStr(s, "CSS with no max_ss") > 0 Then
                AddFinding LOOK_AT, 3, 0, r, LinkCol("max_ss|speed"), "op " & op, tool, _
                    "op " & op & " runs CSS (speed " & Tx(r, "speed") & ") with no max spindle speed - max_ss is " & IIf(Tx(r, "max_ss") = "", "blank", Tx(r, "max_ss")) & ".", _
                    "With constant surface speed the spindle speeds up as the tool moves toward center. With no cap it can run to the machine's top speed - unsafe for the part in the chuck, the bar and the setup.", _
                    "Type max_ss on op " & op & ": the most RPM this setup should run (the chuck's and the bar's limit).", _
                    "css-no-max|op " & op
            End If

            ' Air: feed time spent cutting nothing.
            If Nm(r, "air_pct", air) And Nm(r, "cut_seconds_est", cut) Then
                sec = air / 100 * cut
                If air > AIR_PCT And sec > AIR_MIN_SEC Then
                    AddFinding LOOK_AT, 5, sec, r, LinkCol("air_pct|cut_seconds_est"), "op " & op, tool, _
                        "op " & op & " cuts air " & Format$(air, "0") & "% of its cutting time - " & Tm(sec) & " per part" & _
                        IIf(batch > 1, ", " & Tm(sec * batch) & " per batch of " & batch, "") & ".", _
                        "That is feed time spent cutting nothing, on every part - the easiest time to take out, because the cut itself does not change.", _
                        "Start the cut where the stock really is: check the stock op " & op & " starts from (use the stock the op before it left), trim the chain or the entry and exit amounts, or drop passes that cut nothing. Regenerate, dump and check again.", _
                        "air|op " & op
                End If
            End If

            ' Coolant off on a cutting op.
            If cutting Then
                If CoolantOff(r, how, col) Then
                    AddFinding FYI, 8, 0, r, LinkCol(col), "op " & op, tool, _
                        "op " & op & " cuts with no coolant on (" & how & ").", _
                        "Cutting dry runs the insert and the part hot: shorter edge life, size drift, chips that weld on - unless this material or insert is meant to run dry.", _
                        "Turn the coolant on for op " & op & " (Coolant on the ribbon) - or Ignore this if it runs dry on purpose.", _
                        "coolant-off|op " & op
                End If
            End If

            ' No comment.
            If ColN("comment") > 0 Then
                If Tx(r, "comment") = "" Then
                    AddFinding FYI, 9, 0, r, LinkCol("comment"), "op " & op, tool, _
                        "op " & op & " has no comment.", _
                        "The operator sees the comment in the program and on the setup sheet. A blank one leaves them guessing which op this is when they stop or restart.", _
                        "Type a short comment on op " & op & " (double-click its comment cell) - what it cuts, e.g. ROUGH OD or GROOVE 3MM.", _
                        "no-comment|op " & op
                End If
            End If
        End If
    Next
End Sub

' Whether a row's coolant is off, and how the sheet says so. A V9 machine's one setting
' ("Off"); else X-style: nothing turned on before or with the move.
Private Function CoolantOff(ByVal r As Long, ByRef how As String, ByRef col As String) As Boolean
    Dim v As String, b As String, w As String
    v = Tx(r, "coolant")
    If v <> "" Then
        col = "coolant"
        how = "coolant: " & v
        CoolantOff = (LCase$(v) = "off")
        Exit Function
    End If
    If ColN("coolant_before") = 0 And ColN("coolant_with") = 0 Then Exit Function
    b = Tx(r, "coolant_before")
    w = Tx(r, "coolant_with")
    If b = "" And w = "" Then Exit Function            ' not read for this op
    If TurnsOn(b) Or TurnsOn(w) Then Exit Function
    col = "coolant_with|coolant_before"
    how = "nothing turned on before or with the move"
    CoolantOff = True
End Function

' "10BAR + 70BAR off" - whether any coolant is turned on.
Private Function TurnsOn(ByVal s As String) As Boolean
    Dim p As Variant, t As String
    For Each p In Split(s, "+")
        t = LCase$(Trim$(p))
        If t <> "" And t <> "none" And Right$(t, 4) <> " off" And Left$(t, 5) <> "code " Then TurnsOn = True
    Next
End Function

' "SFM" / "m/min" for CSS, "RPM" for RPM.
Private Function SpeedUnit(ByVal r As Long) As String
    If UCase$(Tx(r, "speed_mode")) = "CSS" Then
        SpeedUnit = IIf(Tx(r, "units") = "mm", "m/min", "SFM")
    Else
        SpeedUnit = "RPM"
    End If
End Function

' One tool, one kind of op: the speeds (and feeds) it runs, compared within the same
' mode (CSS with CSS, per rev with per rev).
Private Sub MixedSpeeds()
    Dim groups As Object, r As Long, k As Variant, rows As Collection, tool As String, kind As String
    Dim sTxt As String, fTxt As String, slow As Long, kindName As String
    Set groups = NewDict()
    For r = FIRST_ROW To mLast
        tool = ToolOf(r)
        kind = UCase$(Tx(r, "type"))
        If OpOf(r) <> "" And tool <> "" And kind <> "" And kind <> "MANUAL" And kind <> "TRANSFORM" Then
            k = tool & "|" & kind
            If Not groups.Exists(k) Then groups.Add k, New Collection
            Set rows = groups(k)
            rows.Add r
        End If
    Next
    For Each k In groups.Keys
        Set rows = groups(k)
        If rows.Count >= 2 Then
            slow = 0
            sTxt = SpreadText(rows, "speed", "speed_mode", True, slow)
            fTxt = SpreadText(rows, "feed", "feed_mode", False, slow)
            If sTxt <> "" Or fTxt <> "" Then
                tool = Split(k, "|")(0)
                kindName = Tx(rows(1), "type")
                AddFinding FYI, 6, 0, slow, LinkCol(IIf(sTxt <> "", "speed|feed", "feed|speed")), Panel.RowsText(rows), tool, _
                    "T" & tool & " runs its " & kindName & " ops at different " & IIf(sTxt <> "" And fTxt <> "", "speeds and feeds", IIf(sTxt <> "", "speeds", "feeds")) & ": " & _
                    sTxt & IIf(sTxt <> "" And fTxt <> "", ";  ", "") & fTxt & ".", _
                    "The same insert doing the same kind of cut usually wants the same speed and feed. The slower op may be leaving time on the table - or the faster one wearing the insert early.", _
                    "Pick the speed and feed that work and use them on every " & kindName & " op of T" & tool & ": click a cell in the good op's row, select those columns, then Copy from op on the ribbon, Into: every other op of its tool. Ignore this if the difference is on purpose (another diameter, an interrupted cut).", _
                    "mixed-speeds|T" & tool & "|" & kindName
            End If
        End If
    Next
End Sub

' "speed 500 - 600 SFM (op 4 500, op 1 600)" when one mode's values are more than SPREAD_MAX
' apart, else "". slowRow: the row with the lowest speed (or feed), for the link.
Private Function SpreadText(ByVal rows As Collection, ByVal name As String, ByVal modeName As String, ByVal isSpeed As Boolean, _
                        ByRef slowRow As Long) As String
    Dim modes As Object, r As Variant, m As Variant, x As Double, lo As Double, hi As Double, loRow As Long
    Dim s As String, list As String, n As Long, unit As String, these As Collection
    Set modes = NewDict()
    For Each r In rows
        If Nm(CLng(r), name, x) Then
            If x > 0 Then
                m = Tx(CLng(r), modeName)
                If Not modes.Exists(m) Then modes.Add m, New Collection
                Set these = modes(m)
                these.Add CLng(r)
            End If
        End If
    Next
    For Each m In modes.Keys
        Set these = modes(m)
        If these.Count >= 2 Then
            lo = 0: hi = 0: loRow = 0: list = "": n = 0
            For Each r In these
                x = 0
                Nm CLng(r), name, x
                If loRow = 0 Or x < lo Then
                    lo = x
                    loRow = CLng(r)
                End If
                If x > hi Then hi = x
                n = n + 1
                If n <= 4 Then list = list & IIf(list = "", "", ", ") & "op " & OpOf(CLng(r)) & " " & Planner.NumText(x)
            Next
            If hi > 0 Then
                If (hi - lo) / hi > SPREAD_MAX Then
                    If isSpeed Then unit = SpeedUnit(CLng(these(1))) Else unit = CStr(m)
                    s = s & IIf(s = "", "", ";  ") & name & " " & Planner.NumText(lo) & " - " & Planner.NumText(hi) & " " & unit & _
                        " (" & list & IIf(n > 4, ", ...", "") & ")"
                    If slowRow = 0 Then slowRow = loRow
                End If
            End If
        End If
    Next
    SpreadText = s
End Function

' A tool put in more than once in a toolpath group, in Operation Manager order. A row
' with no tool (a manual entry) does not end a run - no tool change happens there.
Private Sub ToolChanges()
    Dim lastTool As Object, runs As Object, firstOf As Object, r As Variant, g As String, tool As String, k As Variant
    Dim n As Long, txt As String, i As Long, rc As Collection, same As Boolean
    Set lastTool = NewDict()
    Set runs = NewDict()                      ' group|tool -> Collection of runs (each "1, 2" - its ops)
    Set firstOf = NewDict()                   ' group|tool -> the first row of its second run
    For Each r In mOrder
        tool = ToolOf(CLng(r))
        If tool <> "" Then
            g = Tx(CLng(r), "group_name")
            k = g & "|" & tool
            If Not runs.Exists(k) Then runs.Add k, New Collection
            Set rc = runs(k)
            same = False
            If lastTool.Exists(g) Then same = (lastTool(g) = tool)
            If same And rc.Count > 0 Then
                ' The same run: the op joins its last entry.
                txt = rc(rc.Count)
                rc.Remove rc.Count
                rc.Add txt & ", " & OpOf(CLng(r))
            Else
                rc.Add OpOf(CLng(r))
                If rc.Count = 2 Then firstOf(k) = CLng(r)
            End If
            lastTool(g) = tool
        End If
    Next
    For Each k In runs.Keys
        Set rc = runs(k)
        n = rc.Count
        If n >= 2 Then
            g = Left$(k, InStrRev(k, "|") - 1)          ' a group name may hold a "|" itself
            tool = Mid$(k, InStrRev(k, "|") + 1)
            txt = ""
            For i = 1 To n
                txt = txt & IIf(txt = "", "", ";  ") & IIf(InStr(rc(i), ",") > 0, "ops ", "op ") & rc(i)
            Next
            AddFinding FYI, 7, n - 1, CLng(firstOf(k)), LinkCol("tool"), "op " & OpOf(CLng(firstOf(k))), tool, _
                "T" & tool & " is put in " & n & " times" & IIf(g <> "", " in '" & g & "'", "") & " (" & txt & ") - " & _
                (n - 1) & " tool change" & IIf(n > 2, "s", "") & " per part could be saved.", _
                "Every tool change costs an index and an approach, every part - seconds each that add up over a batch.", _
                "If the part allows it, move T" & tool & "'s ops next to each other in the Operation Manager. Often the order has to stay (stock left for a later op, a feature that must be cut first) - then Ignore this.", _
                "tool-changes|T" & tool & IIf(g <> "", "|" & g, "")
        End If
    Next
End Sub

' ---------------------------------------------------------------- edge time

' Per tool, live, as the Tools page sums them (each op x (1 + xf_copies)): cut seconds and
' flips per part; the longest between flips its unedited ops recorded at the dump; its
' first inspected row and its first row. tool -> Array(cut, flips, longest, inspRow, firstRow, rows)
Private Function ToolSums() As Object
    Dim d As Object, r As Long, tool As String, w As Double, x As Double, cut As Double, flips As Double
    Dim lng As Double, a As Variant, rows As Collection, inspRow As Long, edited As Boolean
    Set d = NewDict()
    For r = FIRST_ROW To mLast
        tool = ToolOf(r)
        If OpOf(r) <> "" And tool <> "" Then
            w = 1
            If Nm(r, "xf_copies", x) Then w = 1 + x
            cut = 0: flips = 0: lng = 0
            If Nm(r, "cut_seconds_est", x) Then cut = x * w
            If Nm(r, "flips_part", x) Then flips = x * w
            ' The longest between flips is from the dump: kept only while the op is as dumped.
            edited = False
            If Nm(r, "changes", x) Then edited = (x > 0)
            If Tx(r, "flip_longest") <> "" And Not edited Then
                lng = Planner.SecondsOfText(Tx(r, "flip_longest"))
                If lng < 0 Then lng = 0
            End If
            inspRow = 0
            If Tx(r, "insp_time") <> "" Or Tx(r, "flip_longest") <> "" Then inspRow = r
            If Not d.Exists(tool) Then
                Set rows = New Collection
                rows.Add r
                d.Add tool, Array(cut, flips, lng, inspRow, r, rows)
            Else
                a = d(tool)
                Set rows = a(5)
                rows.Add r
                d(tool) = Array(a(0) + cut, a(1) + flips, IIf(lng > a(2), lng, a(2)), IIf(a(3) = 0, inspRow, a(3)), a(4), rows)
            End If
        End If
    Next
    Set ToolSums = d
End Function

' How long one edge of a tool cuts, in seconds, live; how that was worked out; and
' whether it is a real edge time (it flips, or parts per edge is typed) or only the
' least it can be (an edge that outlasts a part, how long unknown).
Private Function EdgeOf(ByVal a As Variant, ByVal ppe As Double, ByRef basis As String, ByRef known As Boolean) As Double
    Dim cut As Double, flips As Double, e As Double
    cut = a(0): flips = a(1)
    known = False
    If cut <= 0 Then Exit Function
    If flips >= 1 Then
        e = cut / flips
        basis = Tm(cut) & " of cutting per part over " & Planner.NumText(flips) & " flip" & IIf(flips = 1, "", "s")
        If a(2) > e Then
            e = a(2)
            basis = "its longest edge between flips; " & basis
        End If
        known = True
    ElseIf ppe > 0 Then
        e = cut * ppe
        basis = Tm(cut) & " of cutting per part x " & Planner.NumText(ppe) & " parts per edge"
        known = True
    Else
        e = cut
        basis = "it turns no edge in a part - at least its " & Tm(cut) & " of cutting per part"
    End If
    EdgeOf = e
End Function

Private Function Median(ByVal xs As Collection) As Double
    Dim a() As Double, i As Long, j As Long, t As Double, n As Long
    n = xs.Count
    If n = 0 Then Exit Function
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = xs(i)
    Next
    For i = 2 To n
        t = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next
    If n Mod 2 = 1 Then Median = a((n + 1) \ 2) Else Median = (a(n \ 2) + a(n \ 2 + 1)) / 2
End Function

' A tool whose edge runs far longer than its insert usually does: against the usual edge
' time typed on Tools, else against the other tools carrying that insert in this program.
Private Sub EdgeTimes()
    Dim sums As Object, tool As Variant, other As Variant, ins As String, info As Variant, ppe As Double, usual As Double
    Dim e As Double, basis As String, known As Boolean, oe As Double, ob As String, oknown As Boolean
    Dim others As Collection, names As String, med As Double, r As Long, opsTxt As String, key As String
    Dim a As Variant, rows As Collection, firstOther As String
    Set sums = ToolSums()
    For Each tool In sums.Keys
        ins = ""
        If mInsertOf.Exists(tool) Then ins = mInsertOf(tool)
        If ins <> "" Then
            ppe = 0: usual = 0
            If mIns.Exists(LCase$(ins)) Then
                info = mIns(LCase$(ins))
                ppe = info(3)
                usual = info(4)
            End If
            a = sums(tool)
            e = EdgeOf(a, ppe, basis, known)
            r = a(3)
            If r = 0 Then r = a(4)
            Set rows = a(5)
            opsTxt = Panel.RowsText(rows)
            key = "edge-time|T" & tool & "|" & ins
            If e > 0 And usual > 0 Then
                If e > usual * EDGE_OVER_USUAL Then
                    AddFinding LOOK_AT, 4, e - usual, r, LinkCol("insp_time|flips_part"), opsTxt, CStr(tool), _
                        "T" & tool & " cuts about " & Tm(e) & " on each edge - the usual for " & ins & " is " & Tm(usual) & _
                        " (Tools page), so " & Format$((e / usual - 1) * 100, "0") & "% longer. (" & basis & ")", _
                        EdgeWhy(), _
                        "Turn the insert sooner: insp_time about " & Tm(usual) & " on T" & tool & "'s ops - Inspection & inserts on the ribbon works it out and evens the edges. If this edge really lasts, type a longer usual edge time for " & ins & " on Tools.", _
                        key
                End If
            ElseIf e > 0 Then
                ' No usual typed: the other tools with this insert that have a real edge time.
                Set others = New Collection
                names = ""
                For Each other In sums.Keys
                    If other <> tool And mInsertOf.Exists(other) Then
                        If LCase$(mInsertOf(other)) = LCase$(ins) Then
                            oe = EdgeOf(sums(other), ppe, ob, oknown)
                            If oknown And oe > 0 Then
                                others.Add oe
                                If names = "" Then firstOther = other
                                names = names & IIf(names = "", "", ", ") & "T" & other & " " & Tm(oe)
                            End If
                        End If
                    End If
                Next
                If others.Count > 0 Then
                    med = Median(others)
                    If e > med * EDGE_OVER_MEDIAN And e - med >= EDGE_MIN_EXTRA Then
                        AddFinding LOOK_AT, 4, e - med, r, LinkCol("insp_time|flips_part"), opsTxt, CStr(tool), _
                            "T" & tool & " cuts about " & Tm(e) & " on each edge - " & _
                            IIf(others.Count = 1, "T" & firstOther & ", the other " & ins & " tool in this program, runs " & Tm(med), _
                                "the other " & ins & " tools in this program run about " & Tm(med) & " (" & names & ")") & _
                            ". (" & basis & ")", _
                            EdgeWhy(), _
                            "Turn the insert sooner on T" & tool & ": insp_time about " & Tm(med) & " (Inspection & inserts on the ribbon). If its edge really lasts (a lighter cut), type the usual edge time for " & ins & " on Tools - it is then judged against that.", _
                            key
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Function EdgeWhy() As String
    EdgeWhy = "An edge cut far past its usual life wears: size and finish drift, and a worn edge can chip and scrap the part - or break and take the next edge with it."
End Function

' ============================================================ program check

' Look the program over and write the Program check sheet. The number of findings
' (ignored ones too), or -1 (LastError says why).
Public Function ProgramCheck() As Long
    Dim batch As Double, items As Variant, nIgn As Long, upd As Boolean
    lastErr = ""
    upd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fail
    SyncIgnores                                 ' Ignore set on the last check's sheet
    Set found = New Collection
    If Snapshot() Then
        ReadTools
        batch = BatchQty()
        OpRules batch
        MixedSpeeds
        ToolChanges
        EdgeTimes
    End If
    items = Sorted(IgnoredList(), nIgn)
    WriteCheckSheet items, nIgn
    ProgramCheck = found.Count
Done:
    Application.ScreenUpdating = upd
    Exit Function
Fail:
    lastErr = Err.Description
    ProgramCheck = -1
    Resume Done
End Function

Public Function LastError() As String
    LastError = lastErr
End Function

' The check's count line ("9 findings - 1 ignored (shown greyed at the bottom)").
Public Function CheckLine() As String
    CheckLine = mLine
End Function

' The findings, most useful first: not ignored before ignored, Look at before FYI, by
' rule, the bigger impact first, then sheet order. Marks the ignored ones.
Private Function Sorted(ByVal ign As Object, ByRef nIgn As Long) As Variant
    Dim n As Long, i As Long, j As Long, f As Variant, items() As Variant, keys() As String, t As String, tv As Variant
    n = found.Count
    nIgn = 0
    If n = 0 Then Sorted = Array(): Exit Function
    ReDim items(1 To n): ReDim keys(1 To n)
    For i = 1 To n
        f = found(i)
        f(F_IGN) = ign.Exists(f(F_KEY))
        If f(F_IGN) Then nIgn = nIgn + 1
        items(i) = f
        keys(i) = IIf(f(F_IGN), "1", "0") & f(F_SEV) & Format$(f(F_RANK), "00") & _
                  Format$(99999999 - Application.Min(99999999, Int(f(F_IMPACT))), "00000000") & Format$(f(F_ROW), "000000")
    Next
    For i = 2 To n
        t = keys(i): tv = items(i): j = i - 1
        Do While j >= 1
            If keys(j) <= t Then Exit Do
            keys(j + 1) = keys(j): items(j + 1) = items(j): j = j - 1
        Loop
        keys(j + 1) = t: items(j + 1) = tv
    Next
    Sorted = items
End Function

Private Sub WriteCheckSheet(ByVal items As Variant, ByVal nIgn As Long)
    Dim out As Worksheet, n As Long, i As Long, r As Long, f As Variant, data() As Variant, nLook As Long, nFyi As Long
    Dim last As Long, fc As FormatCondition, mainName As String
    Set out = NewSheet(CHECK_SHEET)
    Widths out, Array(4, 9, 12, 16, 48, 44, 52, 8, 10)
    n = 0
    If IsArray(items) Then n = UBound(items) - LBound(items) + 1
    For i = 1 To n
        f = items(i)
        If Not f(F_IGN) Then
            If f(F_SEV) = LOOK_AT Then nLook = nLook + 1 Else nFyi = nFyi + 1
        End If
    Next
    If n = 0 Then
        mLine = "No findings - nothing to look at."
    Else
        mLine = n & " finding" & IIf(n = 1, "", "s") & IIf(nIgn > 0, " - " & nIgn & " ignored (shown greyed at the bottom)", "")
    End If
    out.Range("A1").Value = "Program check"
    out.Range("A1").Font.Size = 16
    out.Range("A1").Font.Bold = True
    HelpLink out.Range("G1"), "check"
    out.Range("A2").Value = mLine
    out.Range("A2").Font.Bold = True
    out.Range("A3").Value = nLook & " to look at, " & nFyi & " FYI  -  checked " & Format$(Now, "yyyy-mm-dd hh:nn") & _
                            ".   Look at: worth fixing before the program goes to the floor.  FYI: worth knowing.  " & _
                            "Click an op to go to it.  Fine as it is? Set Ignore to yes - it stays ignored on the next check and the next dump."
    out.Range("A3").Font.Color = RGB(91, 101, 115)
    out.Range("A5:I5").Value = Array("#", "Severity", "Op", "Tool", "What was found", "Why it matters", "How to fix", "Ignore", "key")
    With out.Range("A5:H5")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    If n > 0 Then
        ReDim data(1 To n, 1 To 9)
        For i = 1 To n
            f = items(i)
            data(i, 1) = i
            data(i, 2) = IIf(f(F_SEV) = LOOK_AT, "Look at", "FYI")
            data(i, 3) = f(F_OPS)
            data(i, 4) = f(F_TOOL)
            data(i, 5) = f(F_WHAT)
            data(i, 6) = f(F_WHY)
            data(i, 7) = f(F_FIX)
            data(i, 8) = IIf(f(F_IGN), "yes", "")
            data(i, 9) = f(F_KEY)
        Next
        last = 5 + n
        out.Range("C6:D" & last).NumberFormat = "@"
        out.Range("I6:I" & last).NumberFormat = "@"
        out.Range("A6").Resize(n, 9).Value = data
        mainName = ParamTable.MainSheet.Name
        For i = 1 To n
            f = items(i)
            r = 5 + i
            If f(F_ROW) >= FIRST_ROW Then
                out.Hyperlinks.Add Anchor:=out.Cells(r, 3), Address:="", _
                    SubAddress:="'" & mainName & "'!" & mMain.Cells(f(F_ROW), f(F_COL)).Address(False, False), _
                    ScreenTip:="Go to op " & OpOf(f(F_ROW)), TextToDisplay:=CStr(f(F_OPS))
            End If
            If f(F_TOOLNO) <> "" Then
                If mToolRow.Exists(f(F_TOOLNO)) Then
                    out.Hyperlinks.Add Anchor:=out.Cells(r, 4), Address:="", _
                        SubAddress:="'Tools'!A" & mToolRow(f(F_TOOLNO)), ScreenTip:="Go to T" & f(F_TOOLNO) & " on the Tools page", _
                        TextToDisplay:=CStr(f(F_TOOL))
                End If
            End If
            If f(F_SEV) = LOOK_AT Then
                out.Cells(r, 2).Interior.Color = RGB(255, 224, 138)
            Else
                out.Cells(r, 2).Interior.Color = RGB(219, 234, 254)
            End If
        Next
        With out.Range("A6:H" & last)
            .VerticalAlignment = xlTop
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
        End With
        out.Range("E6:G" & last).WrapText = True
        ' Ignore: yes or blank. An ignored finding goes grey at once (and to the bottom at the next check).
        With out.Range("H6:H" & last)
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="yes"
            .Validation.IgnoreBlank = True
            .Validation.InputTitle = "Ignore"
            .Validation.InputMessage = "yes = this is fine as it is: the finding goes grey, to the bottom on the next check, and stays ignored on later checks and dumps. Blank = show it again."
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(142, 169, 219)
        End With
        Set fc = out.Range("A6:G" & last).FormatConditions.Add(Type:=xlExpression, Formula1:="=$H6=""yes""")
        fc.Font.Color = RGB(160, 160, 160)
        fc.Interior.Color = RGB(245, 245, 245)
        out.Range("A6:H" & last).Rows.AutoFit
    Else
        out.Range("A6").Value = "Nothing found - the program looks clean."
    End If
    out.Columns("I").Hidden = True
    out.Range("A3").WrapText = False
    PrintSetup out, 5, 5 + n, "Program check", "H"
    FreezeAt out, 6
End Sub

' ============================================================ slowest ops

' Filter the main sheet to the ops that together make SLOW_SHARE of the estimated cycle
' time, rank them on the Slowest ops sheet, and say so in words (the status line).
Public Function SlowestOps() As String
    Dim est As Long, r As Long, n As Long, rws() As Long, secs() As Double, i As Long, j As Long, k As Long
    Dim total As Double, cum As Double, x As Double, tr As Long, ts As Double, crit() As String, list As String
    Dim nOps As Long, upd As Boolean, s As String
    lastErr = ""
    upd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fail
    If Not Snapshot() Then
        SlowestOps = "No operations on the sheet."
        GoTo Done
    End If
    est = ColN("est_seconds")
    If est = 0 Then
        SlowestOps = "This sheet has no est_seconds column - no times to rank."
        GoTo Done
    End If
    ReDim rws(1 To mLast): ReDim secs(1 To mLast)
    For r = FIRST_ROW To mLast
        If OpOf(r) <> "" Then
            nOps = nOps + 1
            If NumOf(mV(r, est), x) Then
                If x > 0 Then
                    n = n + 1
                    rws(n) = r: secs(n) = x
                    total = total + x
                End If
            End If
        End If
    Next
    If n = 0 Then
        SlowestOps = "No op has an estimated time - regenerate the ops and dump again."
        GoTo Done
    End If
    ' Longest first; equal times keep sheet order.
    For i = 2 To n
        tr = rws(i): ts = secs(i): j = i - 1
        Do While j >= 1
            If secs(j) >= ts Then Exit Do
            rws(j + 1) = rws(j): secs(j + 1) = secs(j): j = j - 1
        Loop
        rws(j + 1) = tr: secs(j + 1) = ts
    Next
    For k = 1 To n
        cum = cum + secs(k)
        If cum >= SLOW_SHARE * total - 0.000001 Then Exit For
    Next
    If k > n Then k = n
    ' The sheet's own filter: every row again first, then only these op numbers.
    With mMain
        If .AutoFilterMode Then
            If .FilterMode Then .ShowAllData
        Else
            .Range(.Cells(HEADER_ROW, 1), .Cells(mLast, mLc)).AutoFilter
        End If
        ReDim crit(0 To k - 1)
        For i = 1 To k
            crit(i - 1) = .Cells(rws(i), 1).Text
        Next
        .AutoFilter.Range.AutoFilter Field:=1, Criteria1:=crit, Operator:=xlFilterValues
    End With
    For i = 1 To Application.Min(k, SLOW_NAMED)
        list = list & IIf(list = "", "", ", ") & "op " & OpOf(rws(i)) & " " & Tm(secs(i))
    Next
    If k > SLOW_NAMED Then list = list & ", ..."
    s = k & " of " & nOps & " ops make " & Format$(SLOW_SHARE * 100, "0") & "% of the cycle time (" & Tm(cum) & " of " & Tm(total) & _
        ") - longest first: " & list & ".  Show all brings the rest back."
    WriteSlowSheet rws, secs, k, n, nOps, total
    On Error Resume Next
    mMain.Activate
    ActiveWindow.ScrollRow = FIRST_ROW
    On Error GoTo Fail
    SlowestOps = s
Done:
    Application.ScreenUpdating = upd
    Exit Function
Fail:
    lastErr = Err.Description
    SlowestOps = "Slowest ops stopped: " & lastErr
    Resume Done
End Function

Private Sub WriteSlowSheet(ByRef rws() As Long, ByRef secs() As Double, ByVal k As Long, ByVal n As Long, ByVal nOps As Long, _
                           ByVal total As Double)
    Dim out As Worksheet, i As Long, r As Long, cum As Double, data() As Variant, x As Double, air As Double, cut As Double
    Set out = NewSheet(SLOW_SHEET)
    Widths out, Array(4, 9, 7, 14, 40, 12, 12, 12, 12, 12)
    out.Range("A1").Value = "Slowest ops - where the cycle time goes"
    out.Range("A1").Font.Size = 16
    out.Range("A1").Font.Bold = True
    HelpLink out.Range("H1"), "slowest"
    out.Range("A2").Value = k & " of " & nOps & " ops make " & Format$(SLOW_SHARE * 100, "0") & "% of the " & Tm(total) & _
                            " cycle time (the live estimate, longest first). 'Lathe params' shows only these now - Show all on the ribbon brings the rest back."
    out.Range("A2").Font.Color = RGB(91, 101, 115)
    out.Range("A4:J4").Value = Array("#", "Op", "Tool", "Type", "Comment", "Time", "Share", "Running share", "Cut time", "Air time")
    With out.Range("A4:J4")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ReDim data(1 To k, 1 To 10)
    For i = 1 To k
        r = rws(i)
        cum = cum + secs(i)
        data(i, 1) = i
        data(i, 2) = "op " & OpOf(r)
        data(i, 3) = IIf(ToolOf(r) = "", "", "T" & ToolOf(r))
        data(i, 4) = Tx(r, "type")
        data(i, 5) = Tx(r, "comment")
        data(i, 6) = Tm(secs(i))
        data(i, 7) = secs(i) / total
        data(i, 8) = cum / total
        If Nm(r, "cut_seconds_est", cut) Then data(i, 9) = Tm(cut)
        If Nm(r, "air_pct", air) And cut > 0 Then
            If air > 0 Then data(i, 10) = Tm(air / 100 * cut)
        End If
        cut = 0
    Next
    out.Range("A5").Resize(k, 10).NumberFormat = "@"
    out.Range("A5").Resize(k, 1).NumberFormat = "0"
    out.Range("G5").Resize(k, 2).NumberFormat = "0.0%"
    out.Range("A5").Resize(k, 10).Value = data
    For i = 1 To k
        r = rws(i)
        out.Hyperlinks.Add Anchor:=out.Cells(4 + i, 2), Address:="", _
            SubAddress:="'" & mMain.Name & "'!" & mMain.Cells(r, LinkCol("est_cycle_time|est_seconds")).Address(False, False), _
            ScreenTip:="Go to op " & OpOf(r), TextToDisplay:="op " & OpOf(r)
    Next
    out.Range("F4").Resize(k + 1, 5).HorizontalAlignment = xlRight
    PrintSetup out, 4, 4 + k, "Slowest ops", "J"
    FreezeAt out, 5
End Sub

' ============================================================ change report

' The Change report sheet: before -> after for the part and a batch, then every change by
' op. A Why typed on the last report is kept for a change that is still there. The cells
' changed, or -1 (LastError says why).
Public Function ChangeReport() As Long
    Dim out As Worksheet, whys As Object, upd As Boolean, qty As Double, r As Long, c As Long, dr As Long
    Dim hdr As String, rw As Long, key As String, t0 As Double, t1 As Double
    Dim ins0 As Double, cost0 As Double, batch0 As Double, ins1 As Double, cost1 As Double, batch1 As Double
    Dim anyEdges As Boolean, anyCost As Boolean, head As String, subLine As String, sm As Worksheet, hdrRow As Long
    Dim rv As Variant, chg As Collection, ch As Variant, nOps As Long
    lastErr = ""
    upd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fail
    Set whys = OldWhys()
    If Not Snapshot() Then Err.Raise vbObjectError + 1, , "No operations on the sheet."
    If mDLast = 0 Then Err.Raise vbObjectError + 2, , "This workbook has no Dumped sheet - nothing to compare with."
    ReadTools
    qty = BatchQty()

    ' The part, from the Summary's heading and the line under it.
    head = "Change report"
    Set sm = SheetNamed("Summary")
    If Not sm Is Nothing Then
        If Left$(SafeText(sm.Range("A1").Value2), 10) = "Summary - " Then head = "Change report - " & Mid$(SafeText(sm.Range("A1").Value2), 11)
        subLine = SafeText(sm.Range("A2").Value2)
    End If

    Set out = NewSheet(REPORT_SHEET)
    out.Cells.Font.Size = 10
    Widths out, Array(9, 6, 30, 14, 14, 14, 24, 44, 10)
    out.Range("A1").Value = head
    out.Range("A1").Font.Size = 16
    out.Range("A1").Font.Bold = True
    HelpLink out.Range("H1"), "report"
    out.Range("A2").Value = "Made " & Format$(Now, "yyyy-mm-dd hh:nn") & IIf(subLine <> "", "   -   " & subLine, "")
    out.Range("A2").Font.Color = RGB(91, 101, 115)

    ' ---- Before -> after: the part's totals as the Summary sums them, the batch typed there.
    rw = 4
    Section out, rw, "Before -> after"
    rw = rw + 1
    out.Range(out.Cells(rw, 5), out.Cells(rw, 7)).Value = Array("As dumped", "Now", "Change")
    out.Cells(rw, 1).Value = "What"
    HeadRow out.Range(out.Cells(rw, 1), out.Cells(rw, 8))
    rw = rw + 1
    t0 = DumpedSum("est_seconds"): t1 = NowSum("est_seconds")
    TotalRow out, rw, "Cycle time per part", Tm(t0), Tm(t1), TimeChange(t0, t1), t1 - t0
    If qty > 1 Then TotalRow out, rw, "Cycle time per batch of " & qty, Tm(t0 * qty), Tm(t1 * qty), TimeChange(t0 * qty, t1 * qty), t1 - t0
    If ColN("cut_seconds_est") > 0 Then
        t0 = DumpedSum("cut_seconds_est"): t1 = NowSum("cut_seconds_est")
        TotalRow out, rw, "Cutting time per part", Tm(t0), Tm(t1), TimeChange(t0, t1), t1 - t0
    End If
    If ColN("flips_part") > 0 Then
        t0 = DumpedSum("flips_part"): t1 = NowSum("flips_part")
        TotalRow out, rw, "Insert flips per part", Planner.NumText(Round(t0, 2)), Planner.NumText(Round(t1, 2)), NumChange(t0, t1, False), t1 - t0
        InsertTotals True, qty, ins0, cost0, batch0, anyEdges, anyCost
        InsertTotals False, qty, ins1, cost1, batch1, anyEdges, anyCost
        If anyEdges Then
            TotalRow out, rw, "Inserts used per part (average)", Planner.NumText(Round(ins0, 3)), Planner.NumText(Round(ins1, 3)), _
                     NumChange(ins0, ins1, False), ins1 - ins0
        End If
        If anyCost Then
            TotalRow out, rw, "Insert cost per part (average)", FormatCurrency(cost0, 2), FormatCurrency(cost1, 2), _
                     NumChange(cost0, cost1, True), cost1 - cost0
            TotalRow out, rw, "Insert cost per batch of " & qty & " (whole inserts)", FormatCurrency(batch0, 2), FormatCurrency(batch1, 2), _
                     NumChange(batch0, batch1, True), batch1 - batch0
        ElseIf anyEdges Then
            out.Cells(rw, 1).Value = "Type the cost per insert on the Tools page for the insert cost per part and per batch."
            out.Cells(rw, 1).Font.Italic = True
            rw = rw + 1
        End If
    End If

    ' ---- Every change, by op, in Operation Manager order (what List changes finds).
    Set chg = New Collection
    For Each rv In mOrder
        r = CLng(rv)
        dr = DRow(r)
        If dr > 0 Then
            For c = 1 To mLc
                hdr = SafeText(mV(HEADER_ROW, c))
                If hdr <> "" And Not ParamTable.Untracked(hdr) And c <= UBound(mD, 2) Then
                    If Not ParamTable.SameValue(mV(r, c), mD(dr, c)) Then
                        ' A formula's cell follows another (a linked percent) - not an edit of its own.
                        If Not mMain.Cells(r, c).HasFormula Then chg.Add Array(r, c, hdr, mD(dr, c), mV(r, c), dr)
                    End If
                End If
            Next
        End If
    Next
    nOps = OpsIn(chg)
    out.Cells(rw, 1).Value = chg.Count & " cell" & IIf(chg.Count = 1, "", "s") & " changed" & _
                             IIf(chg.Count > 0, " in " & nOps & " op" & IIf(nOps = 1, "", "s"), "") & "."
    rw = rw + 2
    Section out, rw, "Every change, by op"
    rw = rw + 1
    hdrRow = rw
    out.Range(out.Cells(rw, 1), out.Cells(rw, 9)).Value = Array("Op", "Tool", "Parameter", "column", "As dumped", "Now", _
                                                                 "Op time change", "Why (type the reason)", "key")
    HeadRow out.Range(out.Cells(rw, 1), out.Cells(rw, 8))
    rw = rw + 1
    If chg.Count = 0 Then
        out.Cells(rw, 1).Value = "No changes from the dump yet - edit 'Lathe params' first."
        out.Cells(rw, 1).Font.Italic = True
        rw = rw + 1
    End If
    r = 0
    For Each ch In chg
        If ch(0) <> r Then
            ' The op's line: op, tool, kind and comment, and what its time did.
            r = ch(0)
            out.Range(out.Cells(rw, 1), out.Cells(rw, 8)).Interior.Color = RGB(237, 242, 250)
            out.Range(out.Cells(rw, 1), out.Cells(rw, 8)).Font.Bold = True
            out.Range(out.Cells(rw, 1), out.Cells(rw, 7)).NumberFormat = "@"
            out.Hyperlinks.Add Anchor:=out.Cells(rw, 1), Address:="", _
                SubAddress:="'" & mMain.Name & "'!" & mMain.Cells(r, 1).Address(False, False), _
                ScreenTip:="Go to op " & OpOf(r), TextToDisplay:="op " & OpOf(r)
            out.Cells(rw, 2).Value = IIf(ToolOf(r) = "", "", "T" & ToolOf(r))
            out.Cells(rw, 3).Value = Tx(r, "type") & IIf(Tx(r, "comment") <> "", "  -  " & Tx(r, "comment"), "")
            t0 = 0: t1 = 0
            If NumOf(Dv(ch(5), "est_seconds"), t0) And Nm(r, "est_seconds", t1) Then
                out.Cells(rw, 7).Value = Tm(t0) & " -> " & Tm(t1) & "  (" & IIf(t1 - t0 < 0, "-", "+") & Tm(Abs(t1 - t0)) & ")"
                If Round(t1 - t0, 0) < 0 Then out.Cells(rw, 7).Font.Color = RGB(21, 128, 61)
                If Round(t1 - t0, 0) > 0 Then out.Cells(rw, 7).Font.Color = RGB(185, 28, 28)
            End If
            rw = rw + 1
        End If
        key = OpOf(r) & "|" & ch(2)
        out.Range(out.Cells(rw, 3), out.Cells(rw, 9)).NumberFormat = "@"
        out.Cells(rw, 3).Value = PlainName(ch(2), mMain.Cells(r, ch(1)))
        out.Cells(rw, 4).Value = ch(2)
        out.Cells(rw, 4).Font.Size = 8
        out.Cells(rw, 4).Font.Color = RGB(120, 120, 120)
        out.Cells(rw, 5).Value = Shown(ch(3))
        out.Cells(rw, 6).Value = Shown(ch(4))
        out.Cells(rw, 6).Font.Bold = True
        If whys.Exists(key) Then out.Cells(rw, 8).Value = whys(key)
        out.Cells(rw, 9).Value = key
        With out.Cells(rw, 8)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(142, 169, 219)
            .WrapText = True
        End With
        rw = rw + 1
    Next
    out.Range(out.Cells(hdrRow + 1, 1), out.Cells(Application.Max(hdrRow + 1, rw), 8)).VerticalAlignment = xlTop
    out.Range(out.Cells(5, 5), out.Cells(Application.Max(hdrRow + 1, rw), 6)).HorizontalAlignment = xlRight
    out.Range(out.Cells(hdrRow + 1, 1), out.Cells(Application.Max(hdrRow + 1, rw), 8)).Rows.AutoFit
    out.Columns("I").Hidden = True
    PrintSetup out, hdrRow, rw - 1, head, "H"
    ChangeReport = chg.Count
Done:
    Application.ScreenUpdating = upd
    Exit Function
Fail:
    lastErr = Err.Description
    ChangeReport = -1
    Resume Done
End Function

' The Why texts of the report there now: "op|column" -> what was typed.
Private Function OldWhys() As Object
    Dim ws As Worksheet, r As Long, last As Long, k As String, w As String, d As Object
    Set d = NewDict()
    Set OldWhys = d
    Set ws = SheetNamed(REPORT_SHEET)
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    For r = 1 To last
        k = SafeText(ws.Cells(r, 9).Value2)
        w = SafeText(ws.Cells(r, 8).Value2)
        If k <> "" And k <> "key" And w <> "" And Not d.Exists(k) Then d.Add k, w
    Next
End Function

' How many ops the changes are in (they come grouped by op).
Private Function OpsIn(ByVal chg As Collection) As Long
    Dim ch As Variant, last As Long
    For Each ch In chg
        If ch(0) <> last Then
            OpsIn = OpsIn + 1
            last = ch(0)
        End If
    Next
End Function

Private Function NowSum(ByVal name As String) As Double
    Dim r As Long, x As Double
    If ColN(name) = 0 Then Exit Function
    For r = FIRST_ROW To mLast
        If Nm(r, name, x) Then NowSum = NowSum + x
    Next
End Function

Private Function DumpedSum(ByVal name As String) As Double
    Dim r As Long, x As Double
    For r = FIRST_ROW To mDLast
        If NumOf(Dv(r, name), x) Then DumpedSum = DumpedSum + x
    Next
End Function

' Per part: inserts used and their cost, and a batch's whole-insert cost - as the Tools
' page's inserts table works them out (flips / edges; 1 / parts per edge where the stops
' turn no edge in a part), from the sheet now or as dumped (the Dumped sheet).
Private Sub InsertTotals(ByVal dumped As Boolean, ByVal qty As Double, ByRef inserts As Double, ByRef cost As Double, _
                         ByRef batchCost As Double, ByRef anyEdges As Boolean, ByRef anyCost As Boolean)
    Dim flipsOf As Object, r As Long, tool As String, x As Double, w As Double, k As Variant, t As Variant, info As Variant
    Dim fl As Double, last As Long
    Set flipsOf = NewDict()
    last = IIf(dumped, mDLast, mLast)
    For r = FIRST_ROW To last
        If dumped Then
            tool = SafeText(Dv(r, "tool"))
        Else
            tool = ToolOf(r)
        End If
        If tool <> "" And tool <> "0" Then
            w = 1
            If dumped Then
                If NumOf(Dv(r, "xf_copies"), x) Then w = 1 + x
                If NumOf(Dv(r, "flips_part"), x) Then flipsOf(tool) = flipsOf(tool) + x * w
            Else
                If Nm(r, "xf_copies", x) Then w = 1 + x
                If Nm(r, "flips_part", x) Then flipsOf(tool) = flipsOf(tool) + x * w
            End If
        End If
    Next
    For Each k In mIns.Keys
        info = mIns(k)
        fl = 0
        For Each t In mInsertOf.Keys
            If LCase$(mInsertOf(t)) = k Then
                If flipsOf.Exists(t) Then fl = fl + flipsOf(t)
            End If
        Next
        If info(3) > 0 And fl < 1 Then fl = 1 / info(3)
        If info(1) > 0 Then
            anyEdges = True
            inserts = inserts + fl / info(1)
            If info(2) >= 0 Then
                anyCost = True
                cost = cost + fl / info(1) * info(2)
                batchCost = batchCost + -Int(-Round(fl * qty / info(1), 6)) * info(2)
            End If
        End If
    Next
End Sub

' A column in plain shop words (the column's own name goes beside it, small).
Private Function PlainName(ByVal hdr As String, ByVal c As Range) As String
    Dim s As String, p As Long
    Select Case hdr
        Case "feed": s = "Feed"
        Case "feed_mode": s = "Feed unit (per rev / per min)"
        Case "speed": s = "Speed (SFM, m/min or RPM)"
        Case "speed_mode": s = "Speed mode (CSS / RPM)"
        Case "spindle_dir": s = "Spindle direction"
        Case "max_ss": s = "Max spindle speed (RPM)"
        Case "plunge": s = "Plunge feed"
        Case "plunge_mode": s = "Plunge feed unit"
        Case "retract": s = "Retract feed"
        Case "retract_mode": s = "Retract feed unit"
        Case "step": s = "Depth of cut"
        Case "min_step": s = "Smallest depth of cut"
        Case "stepover": s = "Stepover"
        Case "stepover_percent": s = "Stepover (% of tool radius)"
        Case "radius": s = "Toolpath radius"
        Case "radius_percent": s = "Toolpath radius (% of tool radius)"
        Case "n_cuts": s = "Number of finish passes"
        Case "peck1": s = "First peck depth"
        Case "peck2": s = "Peck depth"
        Case "stock_x": s = "Stock to leave X"
        Case "stock_z": s = "Stock to leave Z"
        Case "fin_stock_x": s = "Semi-finish stock to leave X"
        Case "fin_stock_z": s = "Semi-finish stock to leave Z"
        Case "coolant": s = "Coolant"
        Case "coolant_before": s = "Coolant before the move"
        Case "coolant_with": s = "Coolant with the move"
        Case "coolant_after": s = "Coolant after the move"
        Case "comment": s = "Op comment"
        Case "manual_text": s = "Manual entry text"
        Case "insp_time": s = "Inspection time (cut between stops)"
        Case "insp_comment": s = "Inspection stop comment"
        Case "insp_do_stop": s = "Tool inspection on"
        Case "insp_dist": s = "Inspection distance (cut between stops)"
        Case "finish_feed": s = "Finish pass feed"
        Case "finish_ss": s = "Finish pass speed"
        Case "rough_step": s = "Rough step"
        Case "finish_step": s = "Finish pass cut"
        Case "pt_rough_speed": s = "PrimeTurning rough speed"
        Case "pt_rough_feed_axial": s = "PrimeTurning rough feed along Z"
        Case "pt_rough_feed_radial": s = "PrimeTurning rough feed along X"
        Case "pt_fin_speed": s = "PrimeTurning finish speed"
        Case "pt_fin_feed_axial": s = "PrimeTurning finish feed along Z"
        Case "pt_fin_feed_radial": s = "PrimeTurning finish feed along X"
    End Select
    If s = "" Then
        ' The column's own tooltip, its first sentence.
        On Error Resume Next
        s = c.Validation.InputMessage
        On Error GoTo 0
        p = InStr(s, vbLf)
        If p > 0 Then s = Left$(s, p - 1)
        p = InStr(s, ". ")
        If p > 0 Then s = Left$(s, p - 1)
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
        If Len(s) > 60 Or s = "" Then s = UCase$(Left$(hdr, 1)) & Replace(Mid$(hdr, 2), "_", " ")
    End If
    PlainName = s
End Function

Private Function Shown(ByVal v As Variant) As String
    If IsError(v) Then Shown = "(error)": Exit Function
    If IsEmpty(v) Then Shown = "(blank)": Exit Function
    Shown = CStr(v)
    If Trim$(Shown) = "" Then Shown = "(blank)"
End Function

' "-1:00  (-2.4%)" - a time's change.
Private Function TimeChange(ByVal a As Double, ByVal b As Double) As String
    Dim d As Double
    d = Round(b - a, 0)
    TimeChange = IIf(d < 0, "-", "+") & Tm(Abs(d)) & Pct(a, b)
End Function

' "-2  (-16.7%)" - a number's change; money in the computer's own currency.
Private Function NumChange(ByVal a As Double, ByVal b As Double, ByVal money As Boolean) As String
    Dim d As Double
    d = b - a
    If money Then
        NumChange = IIf(d < 0, "-", "+") & FormatCurrency(Abs(d), 2) & Pct(a, b)
    Else
        NumChange = IIf(d < 0, "-", "+") & Planner.NumText(Round(Abs(d), 3)) & Pct(a, b)
    End If
End Function

Private Function Pct(ByVal a As Double, ByVal b As Double) As String
    If Abs(a) < 0.0000001 Then Exit Function
    Pct = "  (" & Format$((b - a) / a * 100, "+0.0;-0.0;0.0") & "%)"
End Function

' One line of the before -> after table. delta < 0 (less time, fewer flips, less cost) is
' green, more is red.
Private Sub TotalRow(ByVal out As Worksheet, ByRef rw As Long, ByVal what As String, ByVal was As String, ByVal nowText As String, _
                     ByVal change As String, ByVal delta As Double)
    out.Range(out.Cells(rw, 5), out.Cells(rw, 7)).NumberFormat = "@"
    out.Cells(rw, 1).Value = what
    out.Cells(rw, 5).Value = was
    out.Cells(rw, 6).Value = nowText
    out.Cells(rw, 7).Value = change
    out.Cells(rw, 6).Font.Bold = True
    If Round(delta, 3) < 0 Then out.Cells(rw, 7).Font.Color = RGB(21, 128, 61)
    If Round(delta, 3) > 0 Then out.Cells(rw, 7).Font.Color = RGB(185, 28, 28)
    out.Range(out.Cells(rw, 1), out.Cells(rw, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    out.Range(out.Cells(rw, 1), out.Cells(rw, 8)).Borders(xlEdgeBottom).Color = RGB(225, 225, 225)
    rw = rw + 1
End Sub

Private Sub Section(ByVal out As Worksheet, ByVal rw As Long, ByVal what As String)
    out.Cells(rw, 1).Value = what
    out.Cells(rw, 1).Font.Size = 12
    out.Cells(rw, 1).Font.Bold = True
    out.Cells(rw, 1).Font.Color = RGB(31, 56, 100)
End Sub

Private Sub HeadRow(ByVal rg As Range)
    rg.Font.Bold = True
    rg.Interior.Color = RGB(217, 225, 242)
    rg.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

' Landscape, one page wide, the table's heading on every page. (A computer with no
' printer may refuse some of it - the sheet is still made.)
Private Sub PrintSetup(ByVal out As Worksheet, ByVal hdrRow As Long, ByVal lastLine As Long, ByVal head As String, _
                       ByVal lastCol As String)
    On Error Resume Next
    Application.PrintCommunication = False
    With out.PageSetup
        .PrintArea = "$A$1:$" & lastCol & "$" & Application.Max(lastLine, hdrRow + 1)
        .PrintTitleRows = "$" & hdrRow & ":$" & hdrRow
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.6)
        .BottomMargin = Application.InchesToPoints(0.6)
        .CenterFooter = "Page &P of &N"
        .LeftFooter = Replace(head, "&", "&&")
    End With
    Application.PrintCommunication = True
End Sub

' ============================================================ shared by the sheets

' A sheet made again: the old one (if any) deleted, the new one after the main sheet,
' no gridlines.
Private Function NewSheet(ByVal name As String) As Worksheet
    Dim old As Worksheet, ws As Worksheet, alerts As Boolean
    Set old = SheetNamed(name)
    If Not old Is Nothing Then
        alerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        old.Delete
        Application.DisplayAlerts = alerts
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ParamTable.MainSheet)
    ws.Name = name
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set NewSheet = ws
End Function

' Column widths in characters, from A.
Private Sub Widths(ByVal out As Worksheet, ByVal w As Variant)
    Dim i As Long
    For i = 0 To UBound(w)
        out.Columns(i + 1).ColumnWidth = w(i)
    Next
End Sub

' Freeze the rows above `below` (the sheet is the active one, just made).
Private Sub FreezeAt(ByVal ws As Worksheet, ByVal below As Long)
    On Error Resume Next
    If ActiveSheet.Name <> ws.Name Then ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = below - 1
    ActiveWindow.FreezePanes = True
End Sub

' The "?" link to the manual's page for this sheet (help\<topic>.html beside the add-in).
Private Sub HelpLink(ByVal c As Range, ByVal topic As String)
    Dim d As String, p As String
    On Error Resume Next
    d = Panel.HelpFolder()
    If d = "" Then
        c.Value = "?  (the user manual is not installed)"
        c.Font.Color = RGB(120, 120, 120)
        Exit Sub
    End If
    p = d & "\" & topic & ".html"
    If Dir(p) = "" Then p = d & "\index.html"
    c.Worksheet.Hyperlinks.Add Anchor:=c, Address:=p, ScreenTip:="Open the user manual's page for this sheet", _
                               TextToDisplay:="?  How this works"
End Sub

' ============================================================ ribbon

Public Sub RbCheck(control As IRibbonControl)
    If ProgramCheck() < 0 Then
        MsgBox "The program check stopped: " & lastErr, vbExclamation, TITLE
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.Worksheets(CHECK_SHEET).Activate
    ActiveWindow.ScrollRow = 1
    Application.StatusBar = TITLE & ":  " & mLine
End Sub

Public Sub RbSlowest(control As IRibbonControl)
    Dim s As String
    s = SlowestOps()
    If lastErr <> "" Then
        MsgBox s, vbExclamation, TITLE
        Exit Sub
    End If
    Application.StatusBar = TITLE & ":  " & s
End Sub

Public Sub RbReport(control As IRibbonControl)
    If ChangeReport() < 0 Then
        MsgBox "The change report stopped: " & lastErr, vbExclamation, TITLE
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.Worksheets(REPORT_SHEET).Activate
    ActiveWindow.ScrollRow = 1
    Application.StatusBar = TITLE & ":  the change report is ready to print (landscape, one page wide)."
End Sub
