Attribute VB_Name = "Planner"
' Parameter Table Tool - the planning tools on the ribbon: hit a target time, inspection
' and inserts (the edge time that uses every insert edge evenly), what-if scenarios,
' one-click filters, and copy to every op of a tool.
'
' Plain text in the repo (vba\Planner.bas); tools\build_vba.ps1 compiles it with
' ParamTable.bas. The windows (PlanBox, InspectBox, EditBox) call the quiet Functions
' here, and so do the checks (tools\check_macros.ps1) - nothing here shows a message.
'
' HOW A PLAN IS MADE. The sheet's own live formulas (est_seconds, flips_part, mrr_avg) are
' the model: a plan puts trial values into the cells, recalculates, reads the result and
' puts the old values back - so what the window shows is exactly what the sheet will
' say. The trials are never recorded for undo. Nothing is kept until Apply, and then every
' value goes in through TryWrite, so each cell's own limits still have the last word.
'
' PERCENT, everywhere: +10% is 10% more, -10% is 10% less. The sign is always typed - a
' bare 10% is never guessed at (more? less? to 10%?).
Option Explicit

Private Const FIRST_ROW As Long = 3
Private Const HEADER_ROW As Long = 2
' The feed cells the live estimate is tied to ("every feed" - like the feed override knob),
' the main cutting feeds alone, and the speeds.
Private Const FEEDS As String = "feed,finish_feed,plunge,retract,back_feed,retract_feedrate,fr_override,pt_rough_feed_axial,pt_rough_feed_radial,pt_fin_feed_axial,pt_fin_feed_radial"
Private Const MAIN_FEEDS As String = "feed,finish_feed,fr_override,pt_rough_feed_axial,pt_rough_feed_radial,pt_fin_feed_axial,pt_fin_feed_radial"
Private Const SPEEDS As String = "speed,finish_ss,ss_override,pt_rough_speed,pt_fin_speed"
Public Const STORE_SHEET As String = "Scenario store"
Public Const SCEN_SHEET As String = "Scenarios"

' Column numbers, looked up once per plan (0 = the sheet has no such column).
Private cEst As Long, cCut As Long, cFlips As Long, cMrr As Long, cRem As Long, cTool As Long, cInsp As Long

' The plan the window shows and Apply writes (the window takes it with TakePlan).
Private pCells As Collection
Private pVals As Collection
Private pLines As Collection                ' each an Array of 7 strings; the first is the heading
Private pOk As Boolean
Private pTimes(1) As Double                 ' the time of the ops before and with the plan
Private lastAutoSave As String              ' the scenario the last restore kept the sheet's edits in

' The planning windows: one of each, kept while open (WindowGone forgets one).
Private timeWin As PlanBox
Private inspWin As InspectBox

' The inspection window's plan, one entry per tool (see PlanInspect).
Private Type ToolInsp
    tool As String
    insert As String
    first As Long                   ' its rows in aRows: first .. last
    last As Long
    cells As Collection             ' the insp_time cells its flips listen to (editable)
    anyStop As Boolean              ' any of its ops stops for inspection
    cut As Double                   ' cut time per part (s)
    f0 As Double                    ' flips per part now
    L0 As Double                    ' edge time now (s; 0 = none set)
    fMax As Double                  ' flips at the shortest edge time
    fMin As Double                  ' flips at a very long one: what the toolpath fixes
    even As Double                  ' the shortest edge time that keeps f0 (0 = none)
    goal As Double
    L As Double                     ' the edge time suggested (0 = none)
    f1 As Double                    ' flips with it
    change As Boolean               ' Apply would write something
    warn As Boolean                 ' its edges would run longer than now
    note As String                  ' what to know, in a sentence
    stamp As String                 ' its cells and flips as the plan found them (InspectStale)
End Type

Private iT() As ToolInsp
Private nTool As Long
Private aRows() As Long             ' every row of the tools, tool by tool
Private aW() As Double              ' and its weight: 1 + the copies transforms cut of it
Private aN As Long
Private iTotal As Double            ' flips per part of every tool on the sheet now
Private cDo As Long, cTon As Long, cCon As Long, cEnd As Long, cDon As Long, cDist As Long
Private cNon As Long, cNcut As Long, cUnits As Long, cLim As Long, cCopies As Long

' ============================================================ small things

Private Sub LoadCols()
    cEst = ParamTable.ColOf("est_seconds")
    cCut = ParamTable.ColOf("cut_seconds_est")
    cFlips = ParamTable.ColOf("flips_part")
    cMrr = ParamTable.ColOf("mrr_avg")
    cRem = ParamTable.ColOf("removed")
    cTool = ParamTable.ColOf("tool")
    cInsp = ParamTable.ColOf("insp_time")
End Sub

Private Function Num(ByVal v As Variant, ByRef x As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Or Not IsNumeric(v) Then Exit Function
    ElseIf Not IsNumeric(v) Then
        Exit Function
    End If
    x = CDbl(v)
    Num = True
End Function

' A cell of the main sheet as a number, or Empty.
Private Function NumCell(ByVal r As Long, ByVal c As Long) As Variant
    Dim x As Double
    If c = 0 Then Exit Function
    If Num(ParamTable.MainSheet.Cells(r, c).Value, x) Then NumCell = x
End Function

' Seconds from "m:ss", "h:mm:ss" or plain seconds - or a time of day Excel made of a
' typed 9:00 (as the sheet's own formulas read it). -1 when it is not a time.
Public Function SecondsOfText(ByVal s As Variant) As Double
    Dim p() As String, i As Long, t As Double
    SecondsOfText = -1
    If IsError(s) Or IsEmpty(s) Then Exit Function
    If VarType(s) <> vbString Then
        If IsNumeric(s) Then
            t = CDbl(s)
            If t > 0 And t < 1 Then t = t * 86400
            If t >= 0 Then SecondsOfText = t
        End If
        Exit Function
    End If
    s = Replace(Trim$(s), " ", "")
    If s = "" Then Exit Function
    p = Split(s, ":")
    If UBound(p) > 2 Then Exit Function
    For i = 0 To UBound(p)
        If p(i) = "" Or Not IsNumeric(p(i)) Or InStr(p(i), "-") > 0 Or InStr(p(i), "+") > 0 Then Exit Function
        If i > 0 And CDbl(p(i)) >= 60 Then Exit Function
        t = t * 60 + CDbl(p(i))
    Next
    SecondsOfText = t
End Function

' 754 -> "12:34", 3754 -> "1:02:34".
Public Function TimeText(ByVal sec As Double) As String
    Dim s As Long, sign As String
    If sec < 0 Then sign = "-": sec = -sec
    s = CLng(Int(sec + 0.5))
    If s >= 3600 Then
        TimeText = sign & (s \ 3600) & ":" & Format$((s \ 60) Mod 60, "00") & ":" & Format$(s Mod 60, "00")
    Else
        TimeText = sign & (s \ 60) & ":" & Format$(s Mod 60, "00")
    End If
End Function

' A number as a person would write it: 0.0135, 12.5, 220.
Public Function NumText(ByVal x As Double) As String
    Dim s As String
    s = Format$(Sig(x, 5), "0.######")
    If Right$(s, 1) = "." Or Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    NumText = s
End Function

Private Function Sig(ByVal x As Double, ByVal n As Long) As Double
    Dim e As Long
    If x = 0 Then Exit Function
    e = Int(Log(Abs(x)) / Log(10#))
    Sig = Round(x, Application.Max(0, Application.Min(12, n - 1 - e)))
End Function

Private Function Pct(ByVal a As Double, ByVal b As Double) As String
    If a = 0 Then Exit Function
    Pct = Format$((b - a) / a * 100, "+0.0;-0.0;0.0") & "%"
End Function

' A typed change in percent: "+10%" -> 10, "-7.5" -> -7.5 (the % sign may be left off,
' the + or - may not). "" when understood, else what to type instead.
Public Function PercentChange(ByVal s As String, ByRef pc As Double) As String
    Dim t As String
    pc = 0
    t = Replace(Replace(Trim$(s), " ", ""), "%", "")
    If t = "" Then PercentChange = "Type +10% for 10% more, or -10% for 10% less.": Exit Function
    If Left$(t, 1) <> "+" And Left$(t, 1) <> "-" Then
        If IsNumeric(t) Then
            PercentChange = "'" & Trim$(s) & "' - say which way: +" & t & "% for more, -" & t & "% for less."
        Else
            PercentChange = "'" & Trim$(s) & "' is not a percent - type +10% or -10%."
        End If
        Exit Function
    End If
    If Not IsNumeric(Mid$(t, 2)) Or InStr(Mid$(t, 2), "+") > 0 Or InStr(Mid$(t, 2), "-") > 0 Then
        PercentChange = "'" & Trim$(s) & "' is not a percent - type +10% or -10%."
        Exit Function
    End If
    pc = CDbl(t)
    If pc = 0 Then PercentChange = "0% changes nothing.": Exit Function
    If pc <= -100 Then PercentChange = Trim$(s) & " would leave nothing - type a smaller cut, such as -10%.": pc = 0
End Function

' 10 -> "10% more", -10 -> "10% less".
Public Function PercentWords(ByVal pc As Double) As String
    PercentWords = NumText(Abs(pc)) & "% " & IIf(pc < 0, "less", "more")
End Function

Private Function InList(ByVal name As String, ByVal list As String) As Boolean
    InList = (InStr("," & list & ",", "," & name & ",") > 0)
End Function

' A number cell's limits from its rule, and whether it takes whole numbers only.
Private Sub LimitsOf(ByVal c As Range, ByRef lo As Double, ByRef hi As Double, ByRef whole As Boolean)
    Dim t As Long, op As Long, f1 As String, f2 As String
    lo = -1E+300: hi = 1E+300: whole = False
    On Error Resume Next
    t = -1
    t = c.Validation.Type
    If t <> xlValidateDecimal And t <> xlValidateWholeNumber Then Exit Sub
    op = c.Validation.Operator
    f1 = Replace(c.Validation.Formula1, "=", "")
    f2 = Replace(c.Validation.Formula2, "=", "")
    On Error GoTo 0
    whole = (t = xlValidateWholeNumber)
    Select Case op
    Case xlBetween
        If IsNumeric(f1) Then lo = CDbl(f1)
        If IsNumeric(f2) Then hi = CDbl(f2)
    Case xlGreaterEqual, xlGreater
        If IsNumeric(f1) Then lo = CDbl(f1)
    Case xlLessEqual, xlLess
        If IsNumeric(f1) Then hi = CDbl(f1)
    End Select
End Sub

' v0 times k, rounded the way a person would type it, held inside the cell's limits.
' hit says which limit held it ("" = none).
Private Function Scaled(ByVal c As Range, ByVal v0 As Double, ByVal k As Double, ByVal isSpeed As Boolean, _
                        ByRef hit As String) As Double
    Dim lo As Double, hi As Double, whole As Boolean, x As Double
    LimitsOf c, lo, hi, whole
    x = v0 * k
    If whole Or (isSpeed And x >= 50) Then x = Round(x, 0) Else x = Sig(x, 4)
    hit = ""
    If x > hi Then x = hi: hit = "max " & NumText(hi)
    If x < lo Then x = lo: hit = "min " & NumText(lo)
    Scaled = x
End Function

' What an insp_time cell gets for some seconds: "6:15" in a Text cell, else the seconds.
Private Function InspValue(ByVal c As Range, ByVal sec As Double) As Variant
    If c.NumberFormat = "@" Then InspValue = TimeText(sec) Else InspValue = sec
End Function

Private Function Editable(ByVal c As Range) As Boolean
    If c.HasFormula Then Exit Function
    Editable = ParamTable.TakesInput(c)
End Function

Private Function ToolOf(ByVal r As Long) As String
    If cTool > 0 Then ToolOf = Trim$(CStr(ParamTable.MainSheet.Cells(r, cTool).Text))
End Function

Private Function OpOf(ByVal r As Long) As String
    OpOf = CStr(ParamTable.MainSheet.Cells(r, 1).Value)
End Function

' "ops 2, 9, 11" (eight at most).
Private Function OpsText(ByRef rows() As Long, ByVal n As Long) As String
    Dim i As Long, s As String
    For i = 1 To n
        If i <= 8 Then s = s & IIf(s = "", "", ", ") & OpOf(rows(i))
    Next
    If n > 8 Then s = s & ", ..."
    OpsText = IIf(n = 1, "op ", "ops ") & s
End Function

' ============================================================ which operations

' The data rows a range touches, in sheet order (1-based; n says how many).
Public Function RowsIn(ByVal r As Range, ByRef n As Long) As Long()
    Dim out() As Long, c As Range, seen As String, last As Long
    ReDim out(1 To 1)
    n = 0
    If r Is Nothing Then RowsIn = out: Exit Function
    last = ParamTable.LastRow
    For Each c In Intersect(r.EntireRow, ParamTable.MainSheet.Columns(1)).Cells
        If c.Row >= FIRST_ROW And c.Row <= last Then
            If InStr(seen, "|" & c.Row & "|") = 0 Then
                seen = seen & "|" & c.Row & "|"
                n = n + 1
                ReDim Preserve out(1 To n)
                out(n) = c.Row
            End If
        End If
    Next
    SortRows out, n
    RowsIn = out
End Function

Private Sub SortRows(ByRef a() As Long, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 2 To n
        t = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next
End Sub

' Every row with a tool number, as a range of op_idn cells (Nothing when none).
Public Function ToolRange(ByVal tool As String) As Range
    Dim r As Long, out As Range
    LoadCols
    If cTool = 0 Or tool = "" Then Exit Function
    For r = FIRST_ROW To ParamTable.LastRow
        If ToolOf(r) = tool Then
            If out Is Nothing Then Set out = ParamTable.MainSheet.Cells(r, 1) Else Set out = Union(out, ParamTable.MainSheet.Cells(r, 1))
        End If
    Next
    Set ToolRange = out
End Function

' The three ways to say which ops: the selected rows, every op of the first selected row's
' tool, every op. Captions for the window; ScopeRange gives the rows.
Public Function ScopeChoices(ByVal sel As Range) As Variant
    Dim n As Long, rows() As Long, t As String, tr As Range, nt As Long, trows() As Long
    LoadCols
    rows = RowsIn(ParamTable.DataCells(sel), n)
    If n > 0 Then t = ToolOf(rows(1))
    Set tr = ToolRange(t)
    If Not tr Is Nothing Then trows = RowsIn(tr, nt)
    ScopeChoices = Array( _
        IIf(n = 0, "Selected rows (none selected)", "Selected rows: " & OpsText(rows, n)), _
        IIf(nt = 0, "Every op of this row's tool (no tool on this row)", "Every op of tool " & t & ": " & OpsText(trows, nt)), _
        "Every op on the sheet (" & ParamTable.LastRow - FIRST_ROW + 1 & ")")
End Function

Public Function ScopeRange(ByVal sel As Range, ByVal which As Long) As Range
    Dim n As Long, rows() As Long
    Select Case which
    Case 0
        Set ScopeRange = ParamTable.DataCells(sel)
    Case 1
        LoadCols
        rows = RowsIn(ParamTable.DataCells(sel), n)
        If n > 0 Then Set ScopeRange = ToolRange(ToolOf(rows(1)))
    Case Else
        Set ScopeRange = ParamTable.MainSheet.Range(ParamTable.MainSheet.Cells(FIRST_ROW, 1), _
                                                    ParamTable.MainSheet.Cells(ParamTable.LastRow, 1))
    End Select
End Function

' The window's default: one row selected -> its tool, several -> those rows.
Public Function ScopeDefault(ByVal sel As Range) As Long
    Dim n As Long, rows() As Long
    LoadCols
    rows = RowsIn(ParamTable.DataCells(sel), n)
    If n = 1 And cTool > 0 Then
        If ToolOf(rows(1)) <> "" Then ScopeDefault = 1
    End If
End Function

' Op rows (a window's) as a range of their op_idn cells - Nothing when there are none.
Public Function RowsRange(ByVal rows As Collection) As Range
    Dim r As Variant, out As Range
    If rows Is Nothing Then Exit Function
    For Each r In rows
        If out Is Nothing Then Set out = ParamTable.MainSheet.Cells(r, 1) Else Set out = Union(out, ParamTable.MainSheet.Cells(r, 1))
    Next
    Set RowsRange = out
End Function

' ============================================================ trying values

' Put values in, recalculate, read the rows, put the old values back. Returns per row
' (1..n) est | cut | flips | mrr_avg | removed, Empty where the sheet has no number.
Private Function TrialSnap(ByVal cs As Collection, ByVal vs As Collection, ByRef rows() As Long, ByVal n As Long) As Variant
    Dim olds As New Collection, i As Long, c As Range, calc As Long, ev As Boolean, out As Variant
    ev = Application.EnableEvents
    calc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If calc <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    On Error GoTo Back
    If Not cs Is Nothing Then
        For i = 1 To cs.Count
            Set c = cs(i)
            If c.HasFormula Then olds.Add Array(True, c.Formula) Else olds.Add Array(False, c.Value)
            c.Value = vs(i)
        Next
    End If
    ParamTable.MainSheet.Calculate              ' the plan reads only this sheet
    out = Snap(rows, n)
Back:
    If Not cs Is Nothing Then
        For i = olds.Count To 1 Step -1
            Set c = cs(i)
            If olds(i)(0) Then c.Formula = olds(i)(1) Else c.Value = olds(i)(1)
        Next
    End If
    If calc <> xlCalculationManual Then Application.Calculation = calc Else Application.Calculate
    Application.EnableEvents = ev
    Application.ScreenUpdating = True
    TrialSnap = out
End Function

Private Function Snap(ByRef rows() As Long, ByVal n As Long) As Variant
    Dim out() As Variant, i As Long
    ReDim out(1 To Application.Max(1, n), 1 To 5)
    For i = 1 To n
        out(i, 1) = NumCell(rows(i), cEst)
        out(i, 2) = NumCell(rows(i), cCut)
        out(i, 3) = NumCell(rows(i), cFlips)
        out(i, 4) = NumCell(rows(i), cMrr)
        out(i, 5) = NumCell(rows(i), cRem)
    Next
    Snap = out
End Function

Private Function SumOf(ByVal s As Variant, ByVal n As Long, ByVal k As Long) As Double
    Dim i As Long
    For i = 1 To n
        If Not IsEmpty(s(i, k)) Then SumOf = SumOf + s(i, k)
    Next
End Function

' Material removed / time, over the rows that have both - in3 (cm3) per minute.
Private Function MrrOf(ByVal s As Variant, ByVal n As Long) As Double
    Dim i As Long, rem_ As Double, t As Double
    For i = 1 To n
        If Not IsEmpty(s(i, 5)) And Not IsEmpty(s(i, 1)) Then rem_ = rem_ + s(i, 5): t = t + s(i, 1)
    Next
    If t > 0 Then MrrOf = rem_ / t * 60
End Function

Private Function Arrow(ByVal a As Variant, ByVal b As Variant, ByVal asTime As Boolean) As String
    If IsEmpty(a) And IsEmpty(b) Then Exit Function
    If asTime Then
        Arrow = IIf(IsEmpty(a), "", TimeText(a)) & IIf(SameNum(a, b), "", "  ->  " & IIf(IsEmpty(b), "", TimeText(b)))
    Else
        Arrow = IIf(IsEmpty(a), "", NumText(a)) & IIf(SameNum(a, b), "", "  ->  " & IIf(IsEmpty(b), "", NumText(b)))
    End If
End Function

Private Function SameNum(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) Or IsEmpty(b) Then SameNum = (IsEmpty(a) And IsEmpty(b)): Exit Function
    SameNum = Abs(a - b) < 0.0005 * Application.Max(1, Abs(a))
End Function

' ============================================================ the plan

Private Sub ClearPlan()
    Set pCells = New Collection
    Set pVals = New Collection
    Set pLines = New Collection
    pOk = False
End Sub

Private Sub AddLine(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String, _
                    ByVal f As String, Optional ByVal g As String = "")
    pLines.Add Array(a, b, c, d, e, f, g)
End Sub

' The plan's table for the window: a 2-D array, row 0 the heading - or Empty. Column 6
' is the changes in full (column 5 is cut to fit the window).
Public Function PlanList() As Variant
    Dim out() As String, i As Long, j As Long
    If pLines Is Nothing Then Exit Function
    If pLines.Count = 0 Then Exit Function
    ReDim out(0 To pLines.Count - 1, 0 To 6)
    For i = 1 To pLines.Count
        For j = 0 To 6
            out(i - 1, j) = pLines(i)(j)
        Next
    Next
    PlanList = out
End Function

Public Function PlanCount() As Long
    If Not pCells Is Nothing Then PlanCount = pCells.Count
End Function

' The ops' time before and with the plan last worked out (seconds).
Public Function PlanTimes() As Variant
    PlanTimes = Array(pTimes(0), pTimes(1))
End Function

' Hand the plan to a window, which keeps its own copy (two windows may be open) - Nothing
' when there is none to apply.
Public Sub TakePlan(ByRef cells As Collection, ByRef vals As Collection)
    If pOk And Not pCells Is Nothing Then
        Set cells = pCells
        Set vals = pVals
    Else
        Set cells = Nothing
        Set vals = Nothing
    End If
    Set pCells = Nothing
    Set pVals = Nothing
    pOk = False
End Sub

' Write the plan, each value through its cell's rule - "done refused".
Public Function ApplyPlan() As String
    Dim i As Long, ok As Long, bad As Long
    If pCells Is Nothing Or Not pOk Then ApplyPlan = "0 0": Exit Function
    For i = 1 To pCells.Count
        If ParamTable.TryWrite(pCells(i), pVals(i)) Then ok = ok + 1 Else bad = bad + 1
    Next
    ApplyPlan = ok & " " & bad
    ClearPlan
End Function

' ---------------------------------------------------------------- hit a target time

' Scale the feeds (and speeds) of some ops, all by one factor, so their estimated time
' comes to a target: "12:30", "1:02:00", "540" (seconds) or a change - "-10%" for 10%
' less time, "+5%" for 5% more. Each cell stays inside its own limits; the others make up
' for one that is held. "1|summary" when there is a plan, "0|why" when not. PlanList /
' TakePlan (or ApplyPlan) follow.
Public Function PlanTargetTime(ByVal scope As Range, ByVal target As String, ByVal withSpeeds As Boolean, _
                               ByVal mainOnly As Boolean) As String
    Dim rows() As Long, n As Long, i As Long, j As Long, k As Double, T As Double, e0 As Double, fixedT As Double
    Dim c0 As Double, s0 As Variant, s1 As Variant, best As Variant, bestK As Double, bestErr As Double
    Dim cs As Collection, v0 As Collection, isSp As Collection, vs As Collection, hits As Collection, bestVals As Collection, bestHits As Collection
    Dim it As Long, e As Double, ePrev As Double, p As Double, hdr As String, c As Range, x As Double, h As String
    Dim cols As String, what As String, s As String, nh As Long, msg As String, nc As Long, typed As String, pc As Double, why As String

    ClearPlan
    pTimes(0) = 0: pTimes(1) = 0
    LoadCols
    If cEst = 0 Or cCut = 0 Then PlanTargetTime = "0|This sheet has no live time estimate (est_seconds, cut_seconds_est) - dump it again with this version.": Exit Function
    rows = RowsIn(scope, n)
    If n = 0 Then PlanTargetTime = "0|Select one or more ops (or pick a tool).": Exit Function
    s0 = Snap(rows, n)
    e0 = SumOf(s0, n, 1)
    For i = 1 To n
        If Not IsEmpty(s0(i, 1)) Then
            If IsEmpty(s0(i, 2)) Then fixedT = fixedT + s0(i, 1) Else fixedT = fixedT + s0(i, 1) - s0(i, 2): c0 = c0 + s0(i, 2)
        End If
    Next
    If e0 <= 0 Then PlanTargetTime = "0|No time estimate on " & OpsText(rows, n) & ".": Exit Function
    pTimes(0) = e0: pTimes(1) = e0

    typed = Replace(Trim$(target), " ", "")
    If typed = "" Then
        PlanTargetTime = "0|" & OpsText(rows, n) & IIf(n = 1, " takes ", " take ") & TimeText(e0) & " now (" & TimeText(c0) & " of it cutting)." & _
                         vbCrLf & "Type the time you want (" & TimeText(e0 * 0.9) & ") or a change (-10% for 10% less time), then press Enter."
        Exit Function
    End If
    If InStr(typed, "%") > 0 Or Left$(typed, 1) = "+" Or Left$(typed, 1) = "-" Then
        why = PercentChange(typed, pc)
        If why <> "" Then PlanTargetTime = "0|" & why & "  (-10% = 10% less time.)": Exit Function
        T = e0 * (1 + pc / 100)
    Else
        T = SecondsOfText(typed)
        If T <= 0 Then PlanTargetTime = "0|'" & target & "' is not a time - type 12:30, 1:02:00, or a change such as -10%.": Exit Function
    End If
    If Abs(T - e0) < 0.5 Then PlanTargetTime = "0|" & OpsText(rows, n) & IIf(n = 1, " already takes ", " already take ") & TimeText(e0) & ".": Exit Function
    If c0 <= 0 Then PlanTargetTime = "0|No feed time on these ops for the feeds to move.": Exit Function
    If T <= fixedT + 1 Then
        PlanTargetTime = "0|Out of reach: " & TimeText(fixedT) & " of the " & TimeText(e0) & " is rapids, dwells and other time " & _
                         "feeds do not move - the target has to be more than that."
        Exit Function
    End If

    ' The cells: every feed (or the main ones) - and the speeds - of rows with feed time.
    Set cs = New Collection: Set v0 = New Collection: Set isSp = New Collection
    cols = IIf(mainOnly, MAIN_FEEDS, FEEDS) & IIf(withSpeeds, "," & SPEEDS, "")
    For i = 1 To n
        If Not IsEmpty(s0(i, 2)) Then
            If s0(i, 2) > 0 Then
                For j = 1 To ParamTable.LastCol
                    hdr = CStr(ParamTable.MainSheet.Cells(HEADER_ROW, j).Value)
                    If InList(hdr, cols) Then
                        Set c = ParamTable.MainSheet.Cells(rows(i), j)
                        If Num(c.Value, x) And VarType(c.Value) <> vbString Then
                            If x > 0 And Editable(c) Then
                                cs.Add c: v0.Add x: isSp.Add InList(hdr, SPEEDS)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    If cs.Count = 0 Then PlanTargetTime = "0|No feed cells on these ops that can be changed.": Exit Function

    ' One factor for all; the estimate is feed time / k (/ k twice with speeds) plus the
    ' rest - start from that, then correct on the sheet's own figure.
    p = IIf(withSpeeds, 2, 1)
    k = (c0 / (T - fixedT)) ^ (1 / p)
    bestErr = 1E+300
    For it = 1 To 14
        Set vs = New Collection: Set hits = New Collection
        For i = 1 To cs.Count
            vs.Add Scaled(cs(i), v0(i), k, isSp(i), h)
            hits.Add h
        Next
        s1 = TrialSnap(cs, vs, rows, n)
        e = SumOf(s1, n, 1)
        If Abs(e - T) < bestErr Then
            bestErr = Abs(e - T): best = s1: bestK = k: Set bestVals = vs: Set bestHits = hits
        End If
        If Abs(e - T) < 0.5 Then Exit For
        If it > 1 And Abs(e - ePrev) < 0.05 Then Exit For        ' held at the limits: it goes no further
        ePrev = e
        If e - fixedT <= 0 Then Exit For
        k = k * ((e - fixedT) / (T - fixedT)) ^ (1 / p)
        If k < 0.01 Then k = 0.01
        If k > 100 Then k = 100
    Next

    ' The plan: changed cells, a line per op, the totals.
    For i = 1 To cs.Count
        If Abs(bestVals(i) - v0(i)) > 1E-12 Then
            pCells.Add cs(i): pVals.Add bestVals(i)
            nc = nc + 1
        End If
        If bestHits(i) <> "" Then
            nh = nh + 1
            If nh <= 4 Then msg = msg & IIf(msg = "", "", ", ") & "op " & OpOf(cs(i).Row) & " " & _
                                  ParamTable.MainSheet.Cells(HEADER_ROW, cs(i).Column).Value & " (" & bestHits(i) & ")"
        End If
    Next
    AddLine "op", "tool", "time", "flips / part", "MRR avg", "changes"
    For i = 1 To n
        what = ""
        For j = 1 To cs.Count
            If cs(j).Row = rows(i) And Abs(bestVals(j) - v0(j)) > 1E-12 Then
                what = what & IIf(what = "", "", ";  ") & ParamTable.MainSheet.Cells(HEADER_ROW, cs(j).Column).Value & " " & _
                       NumText(v0(j)) & " -> " & NumText(bestVals(j)) & IIf(bestHits(j) <> "", " (" & bestHits(j) & ")", "")
            End If
        Next
        ' The list shows what fits beside the numbers; the window shows the rest for the
        ' op picked.
        AddLine OpOf(rows(i)), ToolOf(rows(i)), Tight(Arrow(s0(i, 1), best(i, 1), True)), Tight(Arrow(s0(i, 3), best(i, 3), False)), _
                Tight(Arrow(s0(i, 4), best(i, 4), False)), IIf(Len(what) > 62, Left$(what, 59) & "...", what), what
    Next
    e = SumOf(best, n, 1)
    pTimes(1) = e
    s = OpsText(rows, n) & ":  " & TimeText(e0) & "  ->  " & TimeText(e) & "  (" & Pct(e0, e) & ")"
    If Abs(e - T) >= 1 Then s = s & "  -  the target " & TimeText(T) & " is out of reach"
    s = s & vbCrLf & IIf(mainOnly, "Main feeds ", "Feeds ") & Format$((bestK - 1) * 100, "+0.0;-0.0;0.0") & "%" & _
        IIf(withSpeeds, ", speeds " & Format$((bestK - 1) * 100, "+0.0;-0.0;0.0") & "%", "") & _
        ".   Flips / part " & Arrow(SumOf(s0, n, 3), SumOf(best, n, 3), False) & _
        IIf(cRem > 0 And MrrOf(s0, n) > 0, ".   MRR avg " & Arrow(MrrOf(s0, n), MrrOf(best, n), False), "") & "."
    If nh > 0 Then s = s & vbCrLf & nh & " cell(s) held at a limit: " & msg & IIf(nh > 4, ", ...", "") & "."
    pOk = (nc > 0)
    If Not pOk Then s = s & vbCrLf & "Nothing to change."
    PlanTargetTime = IIf(pOk, "1|", "0|") & nc & " cell(s) will change.  " & s
End Function

' "11:00  ->  10:22" as the list shows it: "11:00 -> 10:22".
Private Function Tight(ByVal s As String) As String
    Tight = Replace(s, "  ->  ", " -> ")
End Function

' ---------------------------------------------------------------- inspection & inserts
'
' Per tool: its flips per part (flips_part over ALL its ops - an edge carries on from op
' to op), the edge time now, and the edge time that uses every edge evenly - the SHORTEST
' insp_time that still gives the flips wanted, so each edge cuts about the same and the
' last one in a part is not left a stub. Worked out on the sheet's own flips, with every
' tool's trial in the same recalculation (a tool's flips hang on its own ops only).
'
' Stops the toolpath places itself - every so much cut, every few cuts, at the end of an
' op, a comment stop that does not go by time - do not move with insp_time or the feeds
' on the sheet; they change when the toolpath is made again in Mastercam. A tool whose
' flips do not move at any edge time says so, and gets no suggestion.

' The goals the window offers.
Public Function InspectGoals() As Variant
    InspectGoals = Array("Same flips - use every edge evenly", "One flip fewer per part", "One flip more per part")
End Function

' The two ways to say which tools: those of the ops selected, or every tool.
Public Function InspectScopes(ByVal rows As Collection) As Variant
    Dim t As Variant, s As String, n As Long
    LoadCols
    For Each t In ToolsOf(rows, False)
        n = n + 1
        If n <= 8 Then s = s & IIf(s = "", "", ", ") & "T" & t
    Next
    If n > 8 Then s = s & ", ..."
    InspectScopes = Array(IIf(n = 0, "Tools of the ops selected (none selected)", "Tools of the ops selected: " & s), _
                          "Every tool on the sheet (" & ToolsOf(Nothing, True).Count & ")")
End Function

' "|1|12|" - the tools of some rows, to tell whether other rows are about the same tools.
Public Function ToolsKey(ByVal rows As Collection) As String
    Dim t As Variant
    LoadCols
    ToolsKey = "|"
    For Each t In ToolsOf(rows, False)
        ToolsKey = ToolsKey & t & "|"
    Next
End Function

Private Sub LoadInspCols()
    cDo = ParamTable.ColOf("insp_do_stop")
    cTon = ParamTable.ColOf("insp_time_on")
    cCon = ParamTable.ColOf("insp_comment_on")
    cEnd = ParamTable.ColOf("insp_at_end")
    cDon = ParamTable.ColOf("insp_dist_on")
    cDist = ParamTable.ColOf("insp_dist")
    cNon = ParamTable.ColOf("insp_n_cuts_on")
    cNcut = ParamTable.ColOf("insp_n_cuts")
    cUnits = ParamTable.ColOf("units")
    cLim = ParamTable.ColOf("edge_limit")
    cCopies = ParamTable.ColOf("xf_copies")
End Sub

' The tools of some rows (or of every row), in sheet order, each once.
Private Function ToolsOf(ByVal rows As Collection, ByVal allTools As Boolean) As Collection
    Dim out As New Collection, r As Variant, t As String, seen As String, last As Long, i As Long
    Set ToolsOf = out
    If cTool = 0 Then Exit Function
    If allTools Then
        last = ParamTable.LastRow
        For i = FIRST_ROW To last
            t = ToolOf(i)
            If t <> "" And InStr(seen & "|", "|" & t & "|") = 0 Then seen = seen & "|" & t: out.Add t
        Next
    ElseIf Not rows Is Nothing Then
        For Each r In rows
            t = ToolOf(CLng(r))
            If t <> "" And InStr(seen & "|", "|" & t & "|") = 0 Then seen = seen & "|" & t: out.Add t
        Next
    End If
End Function

' 1 or 0 for an on/off cell, -1 when the sheet has no such column.
Private Function Flag(ByVal r As Long, ByVal c As Long) As Long
    Dim v As Variant
    If c = 0 Then Flag = -1: Exit Function
    v = ParamTable.MainSheet.Cells(r, c).Value
    If IsError(v) Then Exit Function
    If CStr(v) = "1" Or UCase$(CStr(v)) = "TRUE" Then Flag = 1
End Function

' Whether an op's insp_time counts for its flips: it stops for inspection and times its
' stops - or, with no comment to name the stop, the edge time decides whether its stop at
' the end turns the insert. (A sheet without the switches: an op that sets a time.)
Private Function TimeMatters(ByVal r As Long) As Boolean
    If Flag(r, cDo) = 0 Then Exit Function
    Select Case Flag(r, cTon)
    Case 1: TimeMatters = True
    Case 0: TimeMatters = (Flag(r, cCon) <> 1 And Flag(r, cEnd) = 1)
    Case Else: TimeMatters = (SecondsOfText(ParamTable.MainSheet.Cells(r, cInsp).Value) > 0)
    End Select
End Function

' Whether an op stops for inspection at all.
Private Function Stops(ByVal r As Long) As Boolean
    Select Case Flag(r, cDo)
    Case 1: Stops = True
    Case -1: Stops = SecondsOfText(ParamTable.MainSheet.Cells(r, cInsp).Value) > 0 Or Flag(r, cDon) = 1 Or _
                     Flag(r, cNon) = 1 Or Flag(r, cEnd) = 1
    End Select
End Function

' A cell of the Tools page by tool number and heading ("Insert"), as shown - "" when none.
Private Function ToolsText(ByVal tool As String, ByVal head As String) As String
    Dim ws As Worksheet, f As Range, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Tools")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set f = ws.Rows(1).Find(What:=head, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Exit Function
    r = 2
    Do While Trim$(CStr(ws.Cells(r, 1).Text)) <> "" And r < 2000
        If Trim$(CStr(ws.Cells(r, 1).Text)) = tool Then ToolsText = Trim$(CStr(ws.Cells(r, f.Column).Text)): Exit Function
        r = r + 1
    Loop
End Function

' What a tool's stops placed by the toolpath are, in words - "every 5 in of cut, at the
' end of the op" - from its ops' own settings, else the Tools page's Inspection column.
Private Function StopsText(ByVal k As Long) As String
    Dim j As Long, r As Long, s As String, u As String, x As Double
    For j = iT(k).first To iT(k).last
        r = aRows(j)
        If Flag(r, cDo) <> 0 Then
            u = ""
            If cUnits > 0 Then u = Trim$(CStr(ParamTable.MainSheet.Cells(r, cUnits).Value))
            If Flag(r, cDon) = 1 And cDist > 0 Then
                If Num(ParamTable.MainSheet.Cells(r, cDist).Value, x) Then AddWords s, "every " & NumText(x) & IIf(u = "", "", " " & u) & " of cut"
            End If
            If Flag(r, cNon) = 1 And cNcut > 0 Then
                If Num(ParamTable.MainSheet.Cells(r, cNcut).Value, x) Then AddWords s, "every " & NumText(x) & " cuts"
            End If
            If Flag(r, cEnd) = 1 Then AddWords s, "at the end of the op"
        End If
    Next
    If s = "" Then s = ToolsText(iT(k).tool, "Inspection")
    StopsText = s
End Function

Private Sub AddWords(ByRef s As String, ByVal t As String)
    If InStr(", " & s & ",", ", " & t & ",") = 0 Then s = s & IIf(s = "", "", ", ") & t
End Sub

' Each tool's insp_time cells at its own trial edge time (seconds; 0 leaves the tool as
' it is), all in one recalculation - each tool's flips per part.
Private Function TryEdges(ByRef Lx() As Double) As Double()
    Dim cs As New Collection, vs As New Collection, k As Long, j As Long, c As Variant, s As Variant, out() As Double
    ReDim out(1 To nTool)
    For k = 1 To nTool
        If Lx(k) > 0 Then
            For Each c In iT(k).cells
                cs.Add c
                vs.Add InspValue(c, Lx(k))
            Next
        End If
    Next
    s = TrialSnap(cs, vs, aRows, aN)
    For k = 1 To nTool
        For j = iT(k).first To iT(k).last
            If Not IsEmpty(s(j, 3)) Then out(k) = out(k) + s(j, 3) * aW(j)
        Next
    Next
    TryEdges = out
End Function

' For each tool with a goal (-1 = none): the shortest edge time, in whole seconds, whose
' flips are no more than the goal - halving the gap between 5 s (too many flips) and its
' long end (few enough), every tool in each recalculation.
Private Function Shortest(ByRef goal() As Double, ByRef hiB() As Double) As Double()
    Dim lo() As Double, hi() As Double, Lx() As Double, f() As Double, k As Long, more As Boolean, out() As Double
    ReDim lo(1 To nTool): ReDim hi(1 To nTool): ReDim Lx(1 To nTool): ReDim out(1 To nTool)
    For k = 1 To nTool
        lo(k) = 5
        hi(k) = hiB(k)
    Next
    Do
        more = False
        For k = 1 To nTool
            Lx(k) = 0
            If goal(k) >= 0 And hi(k) - lo(k) > 1 Then
                Lx(k) = Int((lo(k) + hi(k)) / 2)
                more = True
            End If
        Next
        If Not more Then Exit Do
        f = TryEdges(Lx)
        For k = 1 To nTool
            If Lx(k) > 0 Then
                If f(k) <= goal(k) Then hi(k) = Lx(k) Else lo(k) = Lx(k)
            End If
        Next
    Loop
    For k = 1 To nTool
        If goal(k) >= 0 Then out(k) = hi(k)
    Next
    Shortest = out
End Function

' The inspection window's work: a line per tool of the rows given (or of every op) for a
' goal - 0 the same flips with every edge used evenly, 1 one flip fewer per part, 2 one
' more. "1|..." when Apply would change something, "0|why" when not; InspectList,
' InspectTotal and InspectApply follow.
Public Function PlanInspect(ByVal rows As Collection, ByVal allTools As Boolean, ByVal goalIdx As Long) As String
    Dim tools As Collection, t As Variant, r As Long, k As Long, j As Long, c As Range, x As Double, last As Long
    Dim Lx() As Double, f() As Double, goal() As Double, hiB() As Double, res() As Double, nc As Long, s0 As Variant

    nTool = 0: aN = 0: iTotal = 0
    LoadCols
    LoadInspCols
    If cFlips = 0 Or cTool = 0 Or cInsp = 0 Then
        PlanInspect = "0|This sheet has no live insert flips (flips_part, insp_time) - dump it again with this version."
        Exit Function
    End If
    Set tools = ToolsOf(rows, allTools)
    If tools.Count = 0 Then
        PlanInspect = "0|" & IIf(allTools, "No op on this sheet has a tool number.", "Select one or more ops (rows with a tool number).")
        Exit Function
    End If

    ' Every row of the tools, tool by tool; the cells that set each one's edge time.
    last = ParamTable.LastRow
    nTool = tools.Count
    ReDim iT(1 To nTool)
    ReDim aRows(1 To 1): ReDim aW(1 To 1)
    k = 0
    For Each t In tools
        k = k + 1
        iT(k).tool = t
        iT(k).first = aN + 1
        Set iT(k).cells = New Collection
        For r = FIRST_ROW To last
            If ToolOf(r) = t Then
                aN = aN + 1
                ReDim Preserve aRows(1 To aN): ReDim Preserve aW(1 To aN)
                aRows(aN) = r
                aW(aN) = 1
                If cCopies > 0 Then If Num(ParamTable.MainSheet.Cells(r, cCopies).Value, x) Then aW(aN) = 1 + x
                If Num(ParamTable.MainSheet.Cells(r, cCut).Value, x) Then iT(k).cut = iT(k).cut + x * aW(aN)
                If Stops(r) Then iT(k).anyStop = True
                If TimeMatters(r) Then
                    Set c = ParamTable.MainSheet.Cells(r, cInsp)
                    If Editable(c) Then iT(k).cells.Add c
                    x = SecondsOfText(c.Value)
                    If x > iT(k).L0 Then iT(k).L0 = x
                End If
            End If
        Next
        iT(k).last = aN
        iT(k).insert = ToolsText(CStr(t), "Insert")
        If iT(k).L0 = 0 And iT(k).cells.Count > 0 Then
            ' No time typed: its stops go by the tool's edge life (edge_limit).
            For j = iT(k).first To iT(k).last
                If cLim > 0 Then
                    If Num(ParamTable.MainSheet.Cells(aRows(j), cLim).Value, x) Then If x > iT(k).L0 Then iT(k).L0 = x
                End If
            Next
            If iT(k).L0 = 0 Then iT(k).L0 = Application.Max(0, SecondsOfText(ToolsText(CStr(t), "Edge life (fallback)")))
        End If
    Next

    ' Now; then the two ends - the shortest edge time (5 s) and one longer than the tool
    ' cuts in a part: what is left at the long end comes from the toolpath itself.
    ReDim Lx(1 To nTool): ReDim hiB(1 To nTool): ReDim goal(1 To nTool)
    f = TryEdges(Lx)
    For k = 1 To nTool
        iT(k).f0 = f(k)
    Next
    iTotal = SheetFlips()
    For k = 1 To nTool
        If iT(k).cells.Count > 0 Then
            Lx(k) = 5
            hiB(k) = Application.Max(2 * iT(k).cut + 60, 4 * iT(k).L0, 600)
        End If
    Next
    f = TryEdges(Lx)
    For k = 1 To nTool
        iT(k).fMax = IIf(Lx(k) > 0, f(k), iT(k).f0)
        Lx(k) = hiB(k)
    Next
    f = TryEdges(Lx)
    For k = 1 To nTool
        iT(k).fMin = IIf(Lx(k) > 0, f(k), iT(k).f0)
    Next

    ' Even: the shortest edge time that keeps today's flips (also says how much of the
    ' last edge a part uses now).
    For k = 1 To nTool
        goal(k) = -1
        If iT(k).cells.Count > 0 And iT(k).fMax > iT(k).f0 And iT(k).f0 >= iT(k).fMin Then goal(k) = iT(k).f0
    Next
    res = Shortest(goal, hiB)
    For k = 1 To nTool
        If goal(k) >= 0 Then iT(k).even = res(k)
        Select Case goalIdx
        Case 1: iT(k).goal = iT(k).f0 - 1
        Case 2: iT(k).goal = iT(k).f0 + 1
        Case Else: iT(k).goal = iT(k).f0
        End Select
    Next
    If goalIdx <> 0 Then
        For k = 1 To nTool
            goal(k) = -1
            If iT(k).cells.Count > 0 And iT(k).goal >= 0 And iT(k).goal >= iT(k).fMin And iT(k).goal < iT(k).fMax Then goal(k) = iT(k).goal
        Next
        res = Shortest(goal, hiB)
    End If

    ' The suggestion: up to whole 5 seconds; "the same flips" never makes an edge longer.
    For k = 1 To nTool
        x = 0
        If goal(k) >= 0 Then x = res(k)
        If goalIdx = 0 And iT(k).f0 = 0 Then x = 0          ' no flips to spread
        If x > 0 Then
            iT(k).L = -Int(-x / 5) * 5
            If goalIdx = 0 And iT(k).L0 > 0 And iT(k).L > iT(k).L0 Then iT(k).L = iT(k).L0
            For Each c In iT(k).cells
                If Abs(SecondsOfText(c.Value) - iT(k).L) >= 0.5 Then iT(k).change = True
            Next
        End If
        Lx(k) = IIf(iT(k).change, iT(k).L, 0)
    Next
    f = TryEdges(Lx)
    For k = 1 To nTool
        iT(k).f1 = IIf(iT(k).change, f(k), iT(k).f0)
        NoteOf k, goalIdx
        iT(k).stamp = ToolStamp(k)
        If iT(k).change Then nc = nc + 1
    Next
    If nc = 0 Then
        PlanInspect = "0|Nothing to change - " & IIf(goalIdx = 0, "every tool here is even already, or its flips come from the toolpath.", "see each tool's note.")
    Else
        PlanInspect = "1|" & nc & " tool(s) can change."
    End If
End Function

' Flips per part of every tool on the sheet now (an op transforms cut again counts again).
Private Function SheetFlips() As Double
    Dim r As Long, x As Double, w As Double
    For r = FIRST_ROW To ParamTable.LastRow
        If ToolOf(r) <> "" Then
            If Num(ParamTable.MainSheet.Cells(r, cFlips).Value, x) Then
                w = 1
                If cCopies > 0 Then If Num(ParamTable.MainSheet.Cells(r, cCopies).Value, w) Then w = 1 + w Else w = 1
                SheetFlips = SheetFlips + x * w
            End If
        End If
    Next
End Function

' What to know about a tool's line, in plain words.
Private Sub NoteOf(ByVal k As Long, ByVal goalIdx As Long)
    Dim why As String, d As Double, s As String
    With iT(k)
        If .cells.Count = 0 Or .fMax = .fMin Then
            ' Nothing on the sheet moves its flips.
            If .f0 = 0 Then
                s = IIf(.anyStop, "No flips - its stops do not turn the insert.", "No flips - it does not stop to turn the insert.")
            Else
                why = StopsText(k)
                s = "Flips come from stops in the toolpath" & IIf(why = "", "", " (" & why & ")") & " - change them in Mastercam."
            End If
        ElseIf .goal < .fMin Then
            why = StopsText(k)
            s = "It cannot go below " & NumText(.fMin) & " - " & IIf(.fMin = 1, "that flip comes", "those flips come") & _
                " from stops in the toolpath" & IIf(why = "", "", " (" & why & ")") & " - change them in Mastercam."
        ElseIf .goal >= .fMax Then
            s = "It cannot flip more often - it already stops as often as its settings allow."
        ElseIf goalIdx = 0 And .f0 = 0 Then
            s = "No flips to spread - one edge does the whole part."
        ElseIf Not .change Then
            s = IIf(goalIdx = 0, "Already even - every edge cuts about the same.", "Already at that edge time.")
        Else
            d = .L - .L0
            If .L0 > 0 And d >= 0.5 Then
                .warn = True
                s = "Edge runs " & TimeText(d) & " longer than now - check the insert can take it."
            ElseIf .L0 > 0 And d <= -0.5 Then
                s = IIf(.f1 = .f0, "Same flips - each edge cuts " & TimeText(-d) & " less.", "Each edge cuts " & TimeText(-d) & " less.")
            Else
                s = "The same edge time on every op of the tool."
            End If
            If .f1 <> .goal Then s = s & "  It gives " & NumText(.f1) & " - the count jumps past " & NumText(.goal) & " there."
        End If
        .note = s
    End With
End Sub

' How much of its edge time the last edge in a part gets to cut now: today's flips at
' the even edge time stand for the cut time a part has for this tool - (f + 1) x even -
' and the last edge gets what f edges of today's time leave of it.
Private Function LastEdgeText(ByVal k As Long) As String
    Dim x As Double
    With iT(k)
        If .even <= 0 Or .L0 <= 0 Then Exit Function
        x = ((.f0 + 1) * .even - .f0 * .L0) / .L0
    End With
    If x < 0 Then x = 0
    If x > 1 Then x = 1
    LastEdgeText = Format$(x * 100, "0") & "%"
End Function

' The window's list: a row per tool - tool, insert, flips now, edge time now, last edge
' used, new edge time, new flips, note; then (hidden) 1 when Apply changes it, the tool,
' flips now and new as numbers, 1 when its edges run longer. Empty when none.
Public Function InspectList() As Variant
    Dim out() As String, k As Long, e As String
    If nTool = 0 Then Exit Function
    ReDim out(0 To nTool - 1, 0 To 12)
    For k = 1 To nTool
        With iT(k)
            If .cells.Count = 0 Or .fMax = .fMin Then
                e = ToolsText(.tool, "Longest between flips")      ' as the toolpath has it
            ElseIf .L0 > 0 Then
                e = TimeText(.L0)
            Else
                e = ""
            End If
            out(k - 1, 0) = "T" & .tool
            out(k - 1, 1) = .insert
            out(k - 1, 2) = NumText(.f0)
            out(k - 1, 3) = e
            out(k - 1, 4) = LastEdgeText(k)
            out(k - 1, 5) = IIf(.change, TimeText(.L), "")
            out(k - 1, 6) = IIf(.change, NumText(.f1), "")
            out(k - 1, 7) = .note
            out(k - 1, 8) = IIf(.change, "1", "0")
            out(k - 1, 9) = .tool
            out(k - 1, 10) = CStr(.f0)
            out(k - 1, 11) = CStr(.f1)
            out(k - 1, 12) = IIf(.warn, "1", "0")
        End With
    Next
    InspectList = out
End Function

' A tool's edge-time cells, its ops and its flips as they are on the sheet now - the plan
' keeps it, to tell whether the sheet changed under the window since.
Private Function ToolStamp(ByVal k As Long) As String
    Dim c As Variant, j As Long, x As Double, f As Double, s As String
    For Each c In iT(k).cells
        s = s & c.Address(False, False) & "=" & CStr(c.Value) & "|"
    Next
    For j = iT(k).first To iT(k).last
        s = s & ParamTable.MainSheet.Cells(aRows(j), 1).Value & ","
        If Num(ParamTable.MainSheet.Cells(aRows(j), cFlips).Value, x) Then f = f + x * aW(j)
    Next
    ToolStamp = s & "|" & f
End Function

' Whether the sheet changed under the plan (typed over, undone, sorted) since it was worked
' out - Apply would then write values worked out on a sheet that is not there any more.
Public Function InspectStale() As Boolean
    Dim k As Long
    For k = 1 To nTool
        If ToolStamp(k) <> iT(k).stamp Then InspectStale = True: Exit Function
    Next
End Function

' Flips per part of every tool on the sheet, when the plan was last worked out.
Public Function InspectTotal() As Double
    InspectTotal = iTotal
End Function

' Write each ticked tool's edge time ("1|12": tool numbers) to its ops, each through its
' cell's rule - "done|refused|ops|tools".
Public Function InspectApply(ByVal ticked As String) As String
    Dim k As Long, c As Variant, ok As Long, bad As Long, ops As String, tl As String
    For k = 1 To nTool
        If iT(k).change And InStr("|" & ticked & "|", "|" & iT(k).tool & "|") > 0 Then
            tl = tl & IIf(tl = "", "", ", ") & "T" & iT(k).tool
            For Each c In iT(k).cells
                If Abs(SecondsOfText(c.Value) - iT(k).L) >= 0.5 Then
                    If ParamTable.TryWrite(c, InspValue(c, iT(k).L)) Then
                        ok = ok + 1
                        ops = ops & IIf(ops = "", "", ", ") & OpOf(c.Row)
                    Else
                        bad = bad + 1
                    End If
                End If
            Next
        End If
    Next
    InspectApply = ok & "|" & bad & "|" & ops & "|" & tl
End Function

' ============================================================ what-if scenarios

' The hidden sheet the scenarios live in: scenario | op_idn | column | value | saved.
' A scenario's "#totals" row keeps cycle time | cut time | flips | removed | cells as saved.
Private Function Store(ByVal make As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(STORE_SHEET)
    On Error GoTo 0
    If ws Is Nothing And make Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = STORE_SHEET
        ws.Range("A1:E1").Value = Array("scenario", "op_idn", "column", "value", "saved")
        ws.Visible = xlSheetHidden
    End If
    Set Store = ws
End Function

Private Function StoreLast(ByVal ws As Worksheet) As Long
    StoreLast = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

' The saved scenarios' names, in the order saved (an empty array when none).
Public Function ScenarioNames() As Variant
    Dim ws As Worksheet, r As Long, s As String, nm As String
    Set ws = Store(False)
    If ws Is Nothing Then ScenarioNames = Array(): Exit Function
    For r = 2 To StoreLast(ws)
        nm = CStr(ws.Cells(r, 1).Value)
        If nm <> "" And InStr(vbLf & s & vbLf, vbLf & nm & vbLf) = 0 Then s = s & IIf(s = "", "", vbLf) & nm
    Next
    If s = "" Then ScenarioNames = Array() Else ScenarioNames = Split(s, vbLf)
End Function

Public Function ScenarioExists(ByVal name As String) As Boolean
    Dim v As Variant
    For Each v In ScenarioNames()
        If LCase$(v) = LCase$(Trim$(name)) Then ScenarioExists = True
    Next
End Function

' Every edit on the sheet now: a Collection of Array(op_idn, column name, value). Each
' column is compared with the dumped one of the same NAME (DumpedColumns).
Private Function CurrentEdits() As Collection
    Dim ws As Worksheet, out As New Collection, r As Long, c As Long, dr As Long, lc As Long, hdr As String, mc As Range
    Dim dcol() As Long
    Set CurrentEdits = out
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dumped")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lc = ParamTable.LastCol
    dcol = ParamTable.DumpedColumns()
    For r = FIRST_ROW To ParamTable.LastRow
        dr = ParamTable.DumpedRow(ParamTable.MainSheet.Cells(r, 1).Value)
        If dr > 0 Then
            For c = 1 To lc
                hdr = CStr(ParamTable.MainSheet.Cells(HEADER_ROW, c).Value)
                Set mc = ParamTable.MainSheet.Cells(r, c)
                If hdr <> "" And Not ParamTable.Untracked(hdr) And Not mc.HasFormula And dcol(c) > 0 Then
                    If Not ParamTable.SameValue(mc.Value, ws.Cells(dr, dcol(c)).Value) Then
                        out.Add Array(ParamTable.MainSheet.Cells(r, 1).Value, hdr, mc.Value)
                    End If
                End If
            Next
        End If
    Next
End Function

Public Function EditCount() As Long
    EditCount = CurrentEdits().Count
End Function

' A scenario's cells: Array(op_idn, column, value) each.
Private Function ScenarioCells(ByVal name As String) As Collection
    Dim ws As Worksheet, out As New Collection, r As Long
    Set ScenarioCells = out
    Set ws = Store(False)
    If ws Is Nothing Then Exit Function
    For r = 2 To StoreLast(ws)
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(Trim$(name)) And CStr(ws.Cells(r, 3).Value) <> "#totals" Then
            out.Add Array(ws.Cells(r, 2).Value, CStr(ws.Cells(r, 3).Value), ws.Cells(r, 4).Value)
        End If
    Next
End Function

Private Function ScenarioSaved(ByVal name As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = Store(False)
    If ws Is Nothing Then Exit Function
    For r = 2 To StoreLast(ws)
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(Trim$(name)) Then ScenarioSaved = ws.Cells(r, 5).Text: Exit Function
    Next
End Function

' The totals of a sheet (the main one, or Dumped): cycle time | cut time | flips | removed
' (sums). Each column found by its name on that sheet's own column-name row.
Private Function TotalsNow(ByVal ws As Worksheet) As Variant
    Dim out(0 To 3) As Double, names As Variant, i As Long, c As Long, r As Long, x As Double, last As Long, f As Range
    names = Array("est_seconds", "cut_seconds_est", "flips_part", "removed")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 0 To 3
        Set f = ws.Rows(HEADER_ROW).Find(What:=CStr(names(i)), LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=True)
        c = 0
        If Not f Is Nothing Then c = f.Column
        If c > 0 Then
            For r = FIRST_ROW To last
                If Num(ws.Cells(r, c).Value, x) Then out(i) = out(i) + x
            Next
        End If
    Next
    TotalsNow = out
End Function

Private Function Key(ByVal e As Collection) As String
    Dim v As Variant, s As String, parts() As String, i As Long, j As Long, t As String
    If e.Count = 0 Then Exit Function
    ReDim parts(1 To e.Count)
    i = 0
    For Each v In e
        i = i + 1
        parts(i) = LCase$(CStr(v(0)) & "|" & v(1) & "|" & Trim$(CStr(v(2))))
    Next
    For i = 2 To e.Count                                 ' sort, so the order does not matter
        t = parts(i): j = i - 1
        Do While j >= 1
            If parts(j) <= t Then Exit Do
            parts(j + 1) = parts(j): j = j - 1
        Loop
        parts(j + 1) = t
    Next
    For i = 1 To e.Count
        s = s & parts(i) & vbLf
    Next
    Key = s
End Function

' Whether the edits on the sheet now are kept in some scenario (or there are none).
Private Function EditsSaved() As Boolean
    Dim k As String, v As Variant
    k = Key(CurrentEdits())
    If k = "" Then EditsSaved = True: Exit Function
    For Each v In ScenarioNames()
        If Key(ScenarioCells(CStr(v))) = k Then EditsSaved = True: Exit Function
    Next
End Function

' Save every edit on the sheet as a scenario (replacing one of that name). The cells saved.
Public Function SaveScenario(ByVal name As String) As Long
    Dim ws As Worksheet, e As Collection, v As Variant, r As Long, tot As Variant, stamp As String
    name = Trim$(name)
    If name = "" Then Exit Function
    Set e = CurrentEdits()
    If e.Count = 0 Then Exit Function
    DeleteScenario name
    Set ws = Store(True)
    r = StoreLast(ws) + 1
    stamp = Format$(Now, "yyyy-mm-dd hh:nn")
    tot = TotalsNow(ParamTable.MainSheet)
    For Each v In e
        ws.Cells(r, 1).Value = name
        ws.Cells(r, 2).Value = v(0)
        ws.Cells(r, 3).Value = v(1)
        If VarType(v(2)) = vbString Then ws.Cells(r, 4).NumberFormat = "@"
        ws.Cells(r, 4).Value = v(2)
        ws.Cells(r, 5).NumberFormat = "@"
        ws.Cells(r, 5).Value = stamp
        r = r + 1
    Next
    ws.Cells(r, 1).Value = name
    ws.Cells(r, 3).Value = "#totals"
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value = tot(0) & "|" & tot(1) & "|" & tot(2) & "|" & tot(3) & "|" & e.Count
    ws.Cells(r, 5).NumberFormat = "@"
    ws.Cells(r, 5).Value = stamp
    SaveScenario = e.Count
End Function

Public Function DeleteScenario(ByVal name As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = Store(False)
    If ws Is Nothing Then Exit Function
    For r = StoreLast(ws) To 2 Step -1
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(Trim$(name)) Then ws.Rows(r).Delete: DeleteScenario = True
    Next
End Function

' The name the edits on the sheet are kept under when a restore would wipe them.
Private Function BeforeRestoreName() As String
    BeforeRestoreName = "(before restore " & Format$(Now, "mmm d hh:nn") & ")"
End Function

' The scenario the last restore saved the sheet's edits into ("" when none needed it).
Public Function AutoSaved() As String
    AutoSaved = lastAutoSave
End Function

' Put the sheet back to the dump, then the scenario's values in through each cell's rule.
' Edits on the sheet that no scenario keeps are saved first, as "(before restore <when>)",
' so a restore never loses work (AutoSaved says under what). "set refused reverted".
Public Function RestoreScenario(ByVal name As String) As String
    Dim e As Collection, v As Variant, f As Range, c As Long, ok As Long, bad As Long, back As Long
    lastAutoSave = ""
    If Not ScenarioExists(name) Then RestoreScenario = "0 0 0": Exit Function
    Set e = ScenarioCells(name)
    If Not EditsSaved() Then
        lastAutoSave = BeforeRestoreName()
        If SaveScenario(lastAutoSave) = 0 Then lastAutoSave = ""
    End If
    back = ParamTable.RevertCells(ParamTable.AllData(), False)
    For Each v In e
        Set f = ParamTable.OpCell(v(0))
        c = ParamTable.ColOf(CStr(v(1)))
        If f Is Nothing Or c = 0 Then
            bad = bad + 1
        ElseIf ParamTable.TryWrite(ParamTable.MainSheet.Cells(f.Row, c), v(2)) Then
            ok = ok + 1
        Else
            bad = bad + 1
        End If
    Next
    RestoreScenario = ok & " " & bad & " " & back
End Function

' The edit window's line for the scenario modes ("1|..." when OK would do something).
Public Function ScenarioPreview(ByVal mode As String, ByVal name As String) As String
    Dim n As Long, e As Long
    name = Trim$(name)
    Select Case mode
    Case "scensave"
        If name = "" Then ScenarioPreview = "0|Type a name for these edits (e.g. faster roughing).": Exit Function
        n = EditCount()
        If n = 0 Then ScenarioPreview = "0|No edits to save - the sheet is as dumped.": Exit Function
        ScenarioPreview = "1|Saves the " & n & " edited cell(s) on the sheet as '" & name & "'" & _
                          IIf(ScenarioExists(name), " - REPLACES the one saved " & ScenarioSaved(name), "") & "."
    Case "scenload"
        If Not ScenarioExists(name) Then ScenarioPreview = "0|Pick a scenario.": Exit Function
        n = ScenarioCells(name).Count
        e = EditCount()
        ScenarioPreview = "1|Sets " & n & " cell(s) from '" & name & "' (saved " & ScenarioSaved(name) & ")" & _
                          IIf(e > 0, "; the " & e & " edit(s) on the sheet now go back to the dump first", "") & "." & _
                          IIf(EditsSaved(), "", vbCrLf & "Your edits now are not in any scenario - they are kept first, as a " & _
                              "scenario named '" & BeforeRestoreName() & "'.")
    Case "scendel"
        If Not ScenarioExists(name) Then ScenarioPreview = "0|Pick the scenario to delete.": Exit Function
        ScenarioPreview = "1|Deletes '" & name & "' (" & ScenarioCells(name).Count & " cell(s), saved " & ScenarioSaved(name) & "). The sheet is not changed."
    End Select
End Function

' A scenario's cells as lines for the window.
Public Function ScenarioDetail(ByVal name As String) As Variant
    Dim e As Collection, v As Variant, s As String, n As Long
    If Not ScenarioExists(name) Then
        Set e = CurrentEdits()
    Else
        Set e = ScenarioCells(name)
    End If
    For Each v In e
        n = n + 1
        If n <= 300 Then s = s & IIf(s = "", "", vbLf) & "op " & v(0) & "    " & v(1) & "  =  " & v(2)
    Next
    If s = "" Then ScenarioDetail = Array() Else ScenarioDetail = Split(s, vbLf)
End Function

' The Scenarios sheet: totals side by side (as dumped, each scenario as saved, the sheet
' now), then every changed parameter in each. The rows it wrote.
Public Function CompareScenarios() As Long
    Dim out As Worksheet, names As Variant, i As Long, r As Long, tot As Variant, d As Worksheet, dump As Variant
    Dim dict As Object, k As Variant, e As Collection, v As Variant, col As Long, keyOf As String, nm As Variant
    Dim sc As Long, cmt As Long, f As Range, dr As Long, hdrCol As Long, parts() As String, dcol() As Long
    names = ScenarioNames()
    dcol = ParamTable.DumpedColumns()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SCEN_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set out = ThisWorkbook.Worksheets.Add(After:=ParamTable.MainSheet)
    out.Name = SCEN_SHEET
    On Error Resume Next
    Set d = ThisWorkbook.Worksheets("Dumped")
    On Error GoTo 0

    out.Range("A1").Value = "What-if scenarios - totals over every op (scenarios as when saved)"
    out.Range("A1").Font.Bold = True
    out.Range("A2:G2").Value = Array("scenario", "cells changed", "cycle time", "vs dump", "cutting time", "flips / part", "MRR avg")
    out.Range("A2:G2").Font.Bold = True
    r = 3
    If Not d Is Nothing Then
        dump = TotalsNow(d)
        TotalsRow out, r, "As dumped", 0, dump, dump: r = r + 1
    End If
    For i = 0 To UBound(names)
        tot = SavedTotals(CStr(names(i)))
        If IsArray(tot) Then
            TotalsRow out, r, CStr(names(i)), ScenarioCells(CStr(names(i))).Count, tot, IIf(IsArray(dump), dump, tot)
            out.Cells(r, 8).Value = "saved " & ScenarioSaved(CStr(names(i)))
            r = r + 1
        End If
    Next
    tot = TotalsNow(ParamTable.MainSheet)
    TotalsRow out, r, "Now (the sheet)", EditCount(), tot, IIf(IsArray(dump), dump, tot)
    out.Rows(r).Font.Italic = True
    r = r + 2

    ' Every parameter any of them changes, side by side.
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(names)
        For Each v In ScenarioCells(CStr(names(i)))
            keyOf = CStr(v(0)) & "|" & v(1)
            If Not dict.Exists(keyOf) Then dict.Add keyOf, dict.Count
        Next
    Next
    For Each v In CurrentEdits()
        keyOf = CStr(v(0)) & "|" & v(1)
        If Not dict.Exists(keyOf) Then dict.Add keyOf, dict.Count
    Next
    out.Cells(r, 1).Value = "Parameters changed"
    out.Cells(r, 1).Font.Bold = True
    r = r + 1
    out.Cells(r, 1).Value = "op_idn": out.Cells(r, 2).Value = "type": out.Cells(r, 3).Value = "comment"
    out.Cells(r, 4).Value = "parameter": out.Cells(r, 5).Value = "dumped"
    For i = 0 To UBound(names)
        out.Cells(r, 6 + i).Value = names(i)
    Next
    out.Cells(r, 6 + UBound(names) + 1).Value = "now"
    out.Rows(r).Font.Bold = True
    hdrCol = 6 + UBound(names) + 1
    cmt = ParamTable.ColOf("comment")
    sc = r
    For Each k In dict.Keys
        r = r + 1
        parts = Split(k, "|")
        out.Cells(r, 1).Value = parts(0)
        out.Cells(r, 4).Value = parts(1)
        Set f = ParamTable.OpCell(parts(0))
        col = ParamTable.ColOf(parts(1))
        If Not f Is Nothing Then
            out.Cells(r, 2).Value = ParamTable.MainSheet.Cells(f.Row, 2).Value
            If cmt > 0 Then out.Cells(r, 3).Value = ParamTable.MainSheet.Cells(f.Row, cmt).Value
            If col > 0 Then
                out.Cells(r, hdrCol).NumberFormat = "@"
                out.Cells(r, hdrCol).Value = CStr(ParamTable.MainSheet.Cells(f.Row, col).Value)
            End If
        End If
        dr = 0
        If Not d Is Nothing Then dr = ParamTable.DumpedRow(parts(0))
        If dr > 0 And col > 0 Then
            If dcol(col) > 0 Then out.Cells(r, 5).NumberFormat = "@": out.Cells(r, 5).Value = CStr(d.Cells(dr, dcol(col)).Value)
        End If
        For i = 0 To UBound(names)
            out.Cells(r, 6 + i).NumberFormat = "@"
            out.Cells(r, 6 + i).Value = out.Cells(r, 5).Value      ' as dumped unless the scenario sets it
            For Each v In ScenarioCells(CStr(names(i)))
                If CStr(v(0)) = parts(0) And v(1) = parts(1) Then
                    out.Cells(r, 6 + i).Value = CStr(v(2))
                    out.Cells(r, 6 + i).Font.Bold = True
                End If
            Next
        Next
    Next
    If dict.Count = 0 Then out.Cells(r + 1, 1).Value = "No scenario changes anything yet - edit the sheet, then Save scenario."
    out.Range(out.Columns(1), out.Columns(hdrCol + 1)).AutoFit
    out.Columns(1).ColumnWidth = Application.Max(out.Columns(1).ColumnWidth, 16)
    CompareScenarios = r
End Function

Private Function SavedTotals(ByVal name As String) As Variant
    Dim ws As Worksheet, r As Long, p() As String, out(0 To 3) As Double, i As Long
    Set ws = Store(False)
    If ws Is Nothing Then Exit Function
    For r = 2 To StoreLast(ws)
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(Trim$(name)) And CStr(ws.Cells(r, 3).Value) = "#totals" Then
            p = Split(CStr(ws.Cells(r, 4).Value), "|")
            For i = 0 To Application.Min(3, UBound(p))
                If IsNumeric(p(i)) Then out(i) = CDbl(p(i))
            Next
            SavedTotals = out
            Exit Function
        End If
    Next
End Function

Private Sub TotalsRow(ByVal out As Worksheet, ByVal r As Long, ByVal what As String, ByVal cellsChanged As Long, _
                      ByVal tot As Variant, ByVal base As Variant)
    out.Cells(r, 1).Value = what
    out.Cells(r, 2).Value = cellsChanged
    out.Cells(r, 3).NumberFormat = "@": out.Cells(r, 3).Value = TimeText(tot(0))
    out.Cells(r, 4).NumberFormat = "@"
    out.Cells(r, 4).Value = IIf(tot(0) - base(0) < 0, "-", "+") & TimeText(Abs(tot(0) - base(0))) & "  (" & Pct(base(0), tot(0)) & ")"
    out.Cells(r, 5).NumberFormat = "@": out.Cells(r, 5).Value = TimeText(tot(1))
    out.Cells(r, 6).Value = Round(tot(2), 2)
    If tot(0) > 0 And tot(3) > 0 Then out.Cells(r, 7).Value = Round(tot(3) / tot(0) * 60, 3)
End Sub

' ============================================================ filters

' The sheet's own filter (on the column-name row) - made again if someone took it off.
Private Function SheetFilter() As AutoFilter
    Dim ws As Worksheet
    Set ws = ParamTable.MainSheet
    If Not ws.AutoFilterMode Then
        ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(ParamTable.LastRow, ParamTable.LastCol)).AutoFilter
    End If
    Set SheetFilter = ws.AutoFilter
End Function

' How many operation rows are showing.
Public Function VisibleOps() As Long
    Dim r As Long
    For r = FIRST_ROW To ParamTable.LastRow
        If Not ParamTable.MainSheet.Rows(r).Hidden Then VisibleOps = VisibleOps + 1
    Next
End Function

' Only the rows with edits (the changes column above 0). Other column filters stay.
Public Function FilterChanged() As Long
    Dim c As Long, f As AutoFilter
    c = ParamTable.ColOf("changes")
    If c = 0 Then FilterChanged = -1: Exit Function
    ParamTable.MainSheet.Calculate
    Set f = SheetFilter()
    f.Range.AutoFilter Field:=c - f.Range.Column + 1, Criteria1:=">0"
    FilterChanged = VisibleOps()
End Function

' Only the ops of a tool (a tool number; "" = the active row's).
Public Function FilterTool(ByVal tool As String) As Long
    Dim c As Long, f As AutoFilter
    c = ParamTable.ColOf("tool")
    If c = 0 Then FilterTool = -1: Exit Function
    If tool = "" Then
        If ActiveCell.Worksheet.Name <> ParamTable.MainSheet.Name Or ActiveCell.Row < FIRST_ROW Then FilterTool = -1: Exit Function
        tool = Trim$(ParamTable.MainSheet.Cells(ActiveCell.Row, c).Text)
    End If
    If tool = "" Then FilterTool = -1: Exit Function
    Set f = SheetFilter()
    f.Range.AutoFilter Field:=c - f.Range.Column + 1, Criteria1:="=" & tool
    FilterTool = VisibleOps()
End Function

' Every row again - the filter dropdowns stay.
Public Function ShowAllOps() As Long
    Dim ws As Worksheet
    Set ws = ParamTable.MainSheet
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
    Else
        SheetFilter
    End If
    ShowAllOps = VisibleOps()
End Function

' ============================================================ every op of a tool

' The cells to copy into: the selected columns, on every OTHER row with the source op's
' tool (hidden rows too - the tool is the point, not the filter). Nothing when none.
Public Function ToolTargets(ByVal cells As Range, ByVal srcOp As Variant) As Range
    Dim src As Range, tool As String, r As Long, c As Range, cols As String, col As Variant, out As Range
    LoadCols
    If cells Is Nothing Or cTool = 0 Then Exit Function
    Set src = ParamTable.OpCell(srcOp)
    If src Is Nothing Then Exit Function
    tool = ToolOf(src.Row)
    If tool = "" Then Exit Function
    For Each c In cells.Cells
        If InStr(cols & ",", "," & c.Column & ",") = 0 Then cols = cols & "," & c.Column
    Next
    For r = FIRST_ROW To ParamTable.LastRow
        If r <> src.Row And ToolOf(r) = tool Then
            For Each col In Split(Mid$(cols, 2), ",")
                If out Is Nothing Then Set out = ParamTable.MainSheet.Cells(r, CLng(col)) Else Set out = Union(out, ParamTable.MainSheet.Cells(r, CLng(col)))
            Next
        End If
    Next
    Set ToolTargets = out
End Function

' Copy the selected columns of one op into every other op of its tool - "done refused".
Public Function ToolCopyCells(ByVal cells As Range, ByVal srcOp As Variant) As String
    Dim t As Range
    Set t = ToolTargets(cells, srcOp)
    If t Is Nothing Then ToolCopyCells = "0 0": Exit Function
    ToolCopyCells = ParamTable.CopyCells(t, srcOp)
End Function

Public Function ToolOfRow(ByVal r As Long) As String
    LoadCols
    ToolOfRow = ToolOf(r)
End Function

' ============================================================ chip thinning

' The chip is thinner than the feed by the sine of the ENTERING ANGLE - the angle between
' the cutting edge and the feed direction (90 = a square shoulder, no thinning): for a
' chip of h the feed per rev is h / sin.

' A round insert's entering angle goes with the depth: sin = sqrt(1 - ((D - 2ap) / D)^2)
' while ap < D / 2; past half the insert it is a full-width chip (1).
Public Function ChipFactor(ByVal dia As Double, ByVal depth As Double) As Double
    Dim c As Double
    If dia <= 0 Or depth <= 0 Then Exit Function
    If depth >= dia / 2 Then ChipFactor = 1: Exit Function
    c = (dia - 2 * depth) / dia
    ChipFactor = Sqr(1 - c * c)
End Function

' A straight edge's is the holder's own (degrees, 0 - 180; a US lead angle is 90 - it):
' 45 gives a chip 0.71 x the feed. 0 = no angle.
Public Function AngleChipFactor(ByVal entering As Double) As Double
    If entering <= 0 Or entering >= 180 Then Exit Function
    AngleChipFactor = Sin(entering * Application.Pi() / 180)
End Function

' The feed per rev that gives a chip of h - and back.
Public Function ChipFeed(ByVal dia As Double, ByVal depth As Double, ByVal chip As Double) As Double
    Dim f As Double
    f = ChipFactor(dia, depth)
    If f > 0 Then ChipFeed = chip / f
End Function

Public Function ChipThickness(ByVal dia As Double, ByVal depth As Double, ByVal feedPerRev As Double) As Double
    ChipThickness = feedPerRev * ChipFactor(dia, depth)
End Function

' ============================================================ ribbon

' Hit a target time, Inspection & inserts, Copy from op and Restore scenario are in
' Panel.bas.
Public Sub RbScenSave(control As IRibbonControl): ParamTable.ScenarioWindow "scensave": End Sub
Public Sub RbScenDel(control As IRibbonControl): ParamTable.ScenarioWindow "scendel": End Sub
Public Sub RbScenCompare(control As IRibbonControl)
    CompareScenarios
    ThisWorkbook.Worksheets(SCEN_SHEET).Activate
End Sub
Public Sub RbShowChanged(control As IRibbonControl)
    Dim n As Long
    n = FilterChanged()
    If n < 0 Then
        MsgBox "This sheet has no changes column.", vbInformation, "Parameter Table"
    Else
        Application.StatusBar = "Parameter Table:  " & n & " op(s) with edits showing - Show all to see every op."
    End If
End Sub
Public Sub RbShowTool(control As IRibbonControl)
    Dim n As Long
    If ActiveSheet.Name <> ParamTable.MainSheet.Name Then ParamTable.MainSheet.Activate
    n = FilterTool("")
    If n < 0 Then
        MsgBox "Click a cell in an operation's row (one with a tool number) first.", vbInformation, "Parameter Table"
    Else
        Application.StatusBar = "Parameter Table:  " & n & " op(s) of tool " & ParamTable.MainSheet.Cells(ActiveCell.Row, ParamTable.ColOf("tool")).Text & " showing."
    End If
End Sub
Public Sub RbShowAll(control As IRibbonControl)
    Application.StatusBar = "Parameter Table:  all " & ShowAllOps() & " op(s) showing."
End Sub

' ============================================================ the planning windows

' The ops a window starts on: the rows selected on the main sheet, else the active cell's
' row - none when neither is an op (the window then asks for one).
Public Function WindowRows() As Collection
    Dim rows As New Collection, r As Long
    Set WindowRows = rows
    On Error Resume Next
    If TypeName(Selection) = "Range" Then Set WindowRows = Panel.OpRows(Selection)
    If WindowRows.Count > 0 Then Exit Function
    If ActiveSheet.Name <> ParamTable.MainSheet.Name Then Exit Function
    r = ActiveCell.Row
    If r >= FIRST_ROW And r <= ParamTable.LastRow Then
        rows.Add r
        Set WindowRows = rows
    End If
End Function

' Hit a target time. One window: pressed again, it comes to the front on the ops selected
' now - unless it holds a plan not applied yet, which it says. The sheet shows the ops'
' time and feed. Writes only on Apply (undo there), never while it is worked out.
Public Sub ShowTargetWindow()
    Dim rows As Collection
    Set rows = WindowRows()
    If timeWin Is Nothing Then
        Set timeWin = New PlanBox
        timeWin.Setup rows
    Else
        timeWin.FollowRows rows
    End If
    If rows.Count > 0 And Not timeWin.Holding() Then Panel.ShowColumns "est_cycle_time|feed", rows
    timeWin.Show vbModeless
    Panel.Opened timeWin
End Sub

' Inspection & inserts - the same way; the sheet shows the ops' inspection and flips.
Public Sub ShowInspectWindow()
    Dim rows As Collection
    Set rows = WindowRows()
    If inspWin Is Nothing Then
        Set inspWin = New InspectBox
        inspWin.Setup rows
    Else
        inspWin.FollowRows rows
    End If
    If rows.Count > 0 And Not inspWin.Holding() Then Panel.ShowColumns "insp_do_stop|insp_time_on|insp_time|flips_part", rows
    inspWin.Show vbModeless
    Panel.Opened inspWin
End Sub

' A planning window closed: the next ribbon press makes a new one.
Public Sub WindowGone(ByVal f As Object)
    If Not timeWin Is Nothing Then
        If f Is timeWin Then Set timeWin = Nothing
    End If
    If Not inspWin Is Nothing Then
        If f Is inspWin Then Set inspWin = Nothing
    End If
End Sub

' For the self-tests: press a window's button ("apply") - or "set G5=0.02", a value typed
' on the sheet in the meantime.
Private Sub SelfTestStep(ByVal f As Object, ByVal p As String)
    Dim kv() As String
    p = Trim$(p)
    If p = "" Then Exit Sub
    If LCase$(Left$(p, 4)) = "set " Then
        kv = Split(Mid$(p, 5), "=")
        If IsNumeric(kv(1)) Then
            ParamTable.MainSheet.Range(kv(0)).Value = CDbl(kv(1))
        Else
            ParamTable.MainSheet.Range(kv(0)).Value = kv(1)
        End If
    Else
        f.Press p
    End If
End Sub

' For tools\check_macros.ps1 (no window shown): the target-time window on a range's ops.
' Fill it in (scopeIdx -1: type the target only, the rest as the window opens), press its
' buttons ("enter", "preview", "apply", "revert" - comma-separated), then, if given, click
' other ops on the sheet; read back
' "summary|Apply on|list rows|result line|speeds note|heading|options ticked".
Public Function PlanSelfTest(ByVal sel As Range, ByVal scopeIdx As Long, ByVal target As String, ByVal mainOnly As Boolean, _
                             ByVal withSpeeds As Boolean, Optional ByVal presses As String = "", _
                             Optional ByVal followSel As Range) As String
    Dim f As PlanBox, p As Variant
    Set f = New PlanBox
    f.Setup Panel.OpRows(sel)
    If scopeIdx < 0 Then f.TypeTarget target Else f.SetInputs scopeIdx, target, mainOnly, withSpeeds
    For Each p In Split(presses, ",")
        SelfTestStep f, CStr(p)
    Next
    If Not followSel Is Nothing Then f.FollowRows Panel.OpRows(followSel)
    PlanSelfTest = f.Summary() & "|" & f.CanApply() & "|" & f.ListRows() & "|" & f.Result() & "|" & f.NoteText() & "|" & _
                   f.Heading() & "|" & f.Options()
    Unload f
End Function

' The inspection window the same way: the tools of a range's ops (or every tool), a goal
' (0 same flips, 1 one fewer, 2 one more), tools to untick ("3|5"), presses ("apply",
' "revert"), other ops clicked after - "summary|Apply on|list rows|result line|heading".
Public Function InspectSelfTest(ByVal sel As Range, ByVal allTools As Boolean, ByVal goalIdx As Long, _
                                Optional ByVal untick As String = "", Optional ByVal presses As String = "", _
                                Optional ByVal followSel As Range) As String
    Dim f As InspectBox, p As Variant
    Set f = New InspectBox
    f.Setup Panel.OpRows(sel)
    f.SetInputs IIf(allTools, 1, 0), goalIdx, untick
    For Each p In Split(presses, ",")
        SelfTestStep f, CStr(p)
    Next
    If Not followSel Is Nothing Then f.FollowRows Panel.OpRows(followSel)
    InspectSelfTest = f.Summary() & "|" & f.CanApply() & "|" & f.ListRows() & "|" & f.Result() & "|" & f.Heading()
    Unload f
End Function
