Attribute VB_Name = "Planner"
' Parameter Table Tool - the planning tools on the ribbon: hit a target time, even out
' insert flips, what-if scenarios, one-click filters, and copy to every op of a tool.
'
' Plain text in the repo (vba\Planner.bas); tools\build_vba.ps1 compiles it with
' ParamTable.bas. The windows (PlanBox, EditBox) call the quiet Functions here, and so do
' the checks (tools\check_macros.ps1) - nothing here shows a window or a message.
'
' HOW A PLAN IS MADE. The sheet's own live formulas (est_seconds, flips_part, mrr_avg) are
' the model: a plan puts trial values into the cells, recalculates, reads the result and
' puts the old values back - so what the window shows is exactly what the sheet will
' say. Nothing is kept until OK, and then every value goes in through TryWrite, so each
' cell's own limits still have the last word.
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

' The plan the window shows and OK writes.
Private pCells As Collection
Private pVals As Collection
Private pLines As Collection                ' each an Array of 6 strings; the first is the heading
Private pOk As Boolean

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

Private Sub AddLine(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String, ByVal f As String)
    pLines.Add Array(a, b, c, d, e, f)
End Sub

' The plan's table for the window: a 2-D array, row 0 the heading - or Empty.
Public Function PlanList() As Variant
    Dim out() As String, i As Long, j As Long
    If pLines Is Nothing Then Exit Function
    If pLines.Count = 0 Then Exit Function
    ReDim out(0 To pLines.Count - 1, 0 To 5)
    For i = 1 To pLines.Count
        For j = 0 To 5
            out(i - 1, j) = pLines(i)(j)
        Next
    Next
    PlanList = out
End Function

Public Function PlanCount() As Long
    If Not pCells Is Nothing Then PlanCount = pCells.Count
End Function

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

' The window's work: which tool ("time" / "flips"), the ops, what was typed and ticked.
Public Function PlanRun(ByVal mode As String, ByVal sel As Range, ByVal scopeIdx As Long, ByVal target As String, _
                        ByVal howIdx As Long, ByVal opt1 As Boolean, ByVal opt2 As Boolean) As String
    Dim scope As Range
    Set scope = ScopeRange(sel, scopeIdx)
    If mode = "time" Then
        PlanRun = PlanTargetTime(scope, target, opt1, opt2)
    Else
        PlanRun = PlanFlips(scope, target, IIf(howIdx = 1, "feed", "insp_time"))
    End If
End Function

' ---------------------------------------------------------------- hit a target time

' Scale the feeds (and speeds) of some ops, all by one factor, so their estimated time
' comes to a target: "12:30", "1:02:00", "540" (seconds) or "10%" (10% less; -5% = 5%
' more). Each cell stays inside its own limits; the others make up for one that is held.
' "1|summary" when there is a plan, "0|why" when not. PlanList / ApplyPlan follow.
Public Function PlanTargetTime(ByVal scope As Range, ByVal target As String, ByVal withSpeeds As Boolean, _
                               ByVal mainOnly As Boolean) As String
    Dim rows() As Long, n As Long, i As Long, j As Long, k As Double, T As Double, e0 As Double, fixedT As Double
    Dim c0 As Double, s0 As Variant, s1 As Variant, best As Variant, bestK As Double, bestErr As Double
    Dim cs As Collection, v0 As Collection, isSp As Collection, vs As Collection, hits As Collection, bestVals As Collection, bestHits As Collection
    Dim it As Long, e As Double, ePrev As Double, p As Double, hdr As String, c As Range, x As Double, h As String
    Dim cols As String, what As String, s As String, nh As Long, msg As String, nc As Long, typed As String

    ClearPlan
    LoadCols
    If cEst = 0 Or cCut = 0 Then PlanTargetTime = "0|This sheet has no live time estimate (est_seconds, cut_seconds_est) - dump it again with this version.": Exit Function
    rows = RowsIn(scope, n)
    If n = 0 Then PlanTargetTime = "0|Select the operations first (or pick a tool).": Exit Function
    s0 = Snap(rows, n)
    e0 = SumOf(s0, n, 1)
    For i = 1 To n
        If Not IsEmpty(s0(i, 1)) Then
            If IsEmpty(s0(i, 2)) Then fixedT = fixedT + s0(i, 1) Else fixedT = fixedT + s0(i, 1) - s0(i, 2): c0 = c0 + s0(i, 2)
        End If
    Next
    If e0 <= 0 Then PlanTargetTime = "0|No time estimate on " & OpsText(rows, n) & ".": Exit Function

    typed = Replace(Trim$(target), " ", "")
    If typed = "" Then
        PlanTargetTime = "0|" & OpsText(rows, n) & IIf(n = 1, " takes ", " take ") & TimeText(e0) & " now (" & TimeText(c0) & " of it cutting)." & _
                         vbCrLf & "Type the time you want (" & TimeText(e0 * 0.9) & ") or a cut (10%), then Enter."
        Exit Function
    End If
    If Right$(typed, 1) = "%" Then
        If Not IsNumeric(Left$(typed, Len(typed) - 1)) Then PlanTargetTime = "0|'" & target & "' - type a percent such as 10%.": Exit Function
        T = e0 * (1 - CDbl(Left$(typed, Len(typed) - 1)) / 100)
    Else
        T = SecondsOfText(typed)
        If T <= 0 Then PlanTargetTime = "0|'" & target & "' is not a time - type 12:30, 1:02:00, or a cut such as 10%.": Exit Function
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
                If Len(what) < 120 Then what = what & IIf(what = "", "", ";  ") & ParamTable.MainSheet.Cells(HEADER_ROW, cs(j).Column).Value & " " & _
                                                NumText(v0(j)) & " -> " & NumText(bestVals(j)) & IIf(bestHits(j) <> "", " (" & bestHits(j) & ")", "")
            End If
        Next
        AddLine OpOf(rows(i)), ToolOf(rows(i)), Arrow(s0(i, 1), best(i, 1), True), Arrow(s0(i, 3), best(i, 3), False), _
                Arrow(s0(i, 4), best(i, 4), False), what
    Next
    e = SumOf(best, n, 1)
    s = OpsText(rows, n) & ":  " & TimeText(e0) & "  ->  " & TimeText(e) & "  (" & Pct(e0, e) & ")"
    If Abs(e - T) >= 1 Then s = s & "  -  the target " & TimeText(T) & " is out of reach"
    s = s & vbCrLf & "Feeds x" & Format$(bestK, "0.000") & IIf(withSpeeds, " and speeds x" & Format$(bestK, "0.000"), "") & _
        ".   Flips / part " & Arrow(SumOf(s0, n, 3), SumOf(best, n, 3), False) & _
        IIf(cRem > 0 And MrrOf(s0, n) > 0, ".   MRR avg " & Arrow(MrrOf(s0, n), MrrOf(best, n), False), "") & "."
    If nh > 0 Then s = s & vbCrLf & nh & " cell(s) held at a limit: " & msg & IIf(nh > 4, ", ...", "") & "."
    pOk = (nc > 0)
    If Not pOk Then s = s & vbCrLf & "Nothing to change."
    PlanTargetTime = IIf(pOk, "1|", "0|") & nc & " cell(s) will change.  " & s
End Function

' ---------------------------------------------------------------- even out flips

' The targets offered in the window.
Public Function FlipTargets() As Variant
    FlipTargets = Array("Even out - same flips, every edge cut the same time", "One flip fewer per tool", _
                        "Two flips fewer per tool", "One flip more per tool")
End Function

' How many flips a tool should end with: "Even..." keeps the count; "One flip fewer" -1,
' "Two flips fewer" -2, "One flip more" +1; a number is the count itself. -1 = not understood.
Private Function FlipGoal(ByVal target As String, ByVal now As Double) As Double
    Dim t As String
    t = LCase$(Trim$(target))
    FlipGoal = -1
    If t = "" Or Left$(t, 4) = "even" Then
        FlipGoal = now
    ElseIf Left$(t, 14) = "one flip fewer" Then
        FlipGoal = now - 1
    ElseIf Left$(t, 15) = "two flips fewer" Then
        FlipGoal = now - 2
    ElseIf Left$(t, 13) = "one flip more" Then
        FlipGoal = now + 1
    ElseIf IsNumeric(t) Then
        If CDbl(t) >= 0 Then FlipGoal = Int(CDbl(t))
    End If
    If FlipGoal < -0.5 Then Exit Function
    If FlipGoal < 0 Then FlipGoal = 0
End Function

' Change insp_time (one edge time for each tool) or the feeds of some ops so each tool's
' flips per part come to a goal - by default the same count with every edge used evenly
' (the last edge is not left a stub): the SHORTEST edge time, or the slowest feeds, that
' still gives that count. The flips are the sheet's own (flips_part, live), summed over
' every op of the tool, since an edge carries on from one op of a tool to its next.
Public Function PlanFlips(ByVal scope As Range, ByVal target As String, ByVal how As String) As String
    Dim rows() As Long, n As Long, tools As String, tl As Variant, i As Long, j As Long
    Dim trows() As Long, tn As Long, cs As Collection, v0 As Collection, vs As Collection, c As Range, x As Double
    Dim f0 As Double, goal As Double, lo As Double, hi As Double, mid_ As Double, fl As Double, fh As Double, fm As Double
    Dim L0 As Double, L As Double, cut As Double, s As String, note As String, nc As Long, it As Long, h As String
    Dim allCells As New Collection, allVals As New Collection, s0 As Variant, s1 As Variant, allRows() As Long, an As Long
    Dim before As Variant, what As String, feedCols As String, kBest As Double, fBest As Double, inspCol As Long

    ClearPlan
    LoadCols
    If cFlips = 0 Or cTool = 0 Then PlanFlips = "0|This sheet has no live flips (flips_part) - dump it again with this version.": Exit Function
    If how = "insp_time" And cInsp = 0 Then PlanFlips = "0|This sheet has no insp_time column.": Exit Function
    rows = RowsIn(scope, n)
    If n = 0 Then PlanFlips = "0|Select the operations first (or pick a tool).": Exit Function

    ' The tools of the selected ops, in order.
    For i = 1 To n
        If ToolOf(rows(i)) <> "" And InStr(tools & "|", "|" & ToolOf(rows(i)) & "|") = 0 Then tools = tools & "|" & ToolOf(rows(i))
    Next
    If tools = "" Then PlanFlips = "0|No tool numbers on these rows.": Exit Function

    For Each tl In Split(Mid$(tools, 2), "|")
        trows = RowsIn(ToolRange(CStr(tl)), tn)
        before = Snap(trows, tn)
        f0 = SumOf(before, tn, 3)
        goal = FlipGoal(target, f0)
        If goal < 0 Then PlanFlips = "0|'" & target & "' - pick from the list or type how many flips per part.": ClearPlan: Exit Function
        cut = SumOf(before, tn, 2)

        ' The cells this tool's selected ops give to change.
        Set cs = New Collection: Set v0 = New Collection
        L0 = -1
        For i = 1 To n
            If ToolOf(rows(i)) = CStr(tl) Then
                If how = "insp_time" Then
                    ' Only ops that inspect (an insp_time is set): a blank one stays blank.
                    Set c = ParamTable.MainSheet.Cells(rows(i), cInsp)
                    x = SecondsOfText(c.Value)
                    If x > 0 And Editable(c) Then
                        cs.Add c
                        v0.Add x
                        If L0 < 0 Then L0 = x
                    End If
                Else
                    For j = 1 To ParamTable.LastCol
                        If InList(CStr(ParamTable.MainSheet.Cells(HEADER_ROW, j).Value), FEEDS) Then
                            Set c = ParamTable.MainSheet.Cells(rows(i), j)
                            If Num(c.Value, x) And VarType(c.Value) <> vbString Then
                                If x > 0 And Editable(c) Then cs.Add c: v0.Add x
                            End If
                        End If
                    Next
                End If
            End If
        Next

        If cs.Count = 0 Then
            note = note & vbCrLf & "Tool " & tl & ": no " & IIf(how = "insp_time", "insp_time", "feed") & " to change on the selected ops."
        ElseIf goal = f0 And goal = 0 Then
            note = note & vbCrLf & "Tool " & tl & ": no flips to even out."
        ElseIf how = "insp_time" Then
            ' Flips fall as the edge time grows: find the shortest that gives the goal.
            lo = 5
            hi = Application.Max(2 * cut + 60, 4 * L0, 600)
            fl = FlipsWithInsp(cs, lo, trows, tn)
            fh = FlipsWithInsp(cs, hi, trows, tn)
            If fl = fh Then
                note = note & vbCrLf & "Tool " & tl & ": insp_time does not move its flips (is the timer, insp_time_on, off?)."
            ElseIf fh > goal Then
                note = note & vbCrLf & "Tool " & tl & ": no edge time gives " & goal & " flips - the fewest is " & fh & "."
            ElseIf fl <= goal Then
                note = note & vbCrLf & "Tool " & tl & ": " & goal & " flips at any edge time - nothing to even out."
            Else
                Do While hi - lo > 1
                    mid_ = Int((lo + hi) / 2)
                    If FlipsWithInsp(cs, mid_, trows, tn) <= goal Then hi = mid_ Else lo = mid_
                Loop
                L = -Int(-hi / 5) * 5                      ' up to whole 5 seconds
                If L0 > 0 And goal = f0 And L > L0 Then L = L0
                If FlipsWithInsp(cs, L, trows, tn) > goal Then L = hi
                If L0 > 0 And Abs(L - L0) < 0.5 Then
                    note = note & vbCrLf & "Tool " & tl & ": already even - " & goal & " flips at " & TimeText(L0) & "."
                Else
                    For i = 1 To cs.Count
                        allCells.Add cs(i): allVals.Add InspValue(cs(i), L)
                    Next
                    note = note & vbCrLf & "Tool " & tl & ": flips " & NumText(f0) & " -> " & NumText(goal) & ",  insp_time " & _
                           IIf(L0 > 0, TimeText(L0) & " -> ", "") & TimeText(L) & IIf(L0 > 0, "  (" & Pct(L0, L) & " edge time)", "") & "."
                End If
            End If
        Else
            ' Flips fall as the feeds rise: the slowest feeds that give the goal.
            lo = Log(0.5): hi = Log(2)
            fl = FlipsWithFeeds(cs, v0, Exp(lo), trows, tn)
            fh = FlipsWithFeeds(cs, v0, Exp(hi), trows, tn)
            If fl = fh Then
                note = note & vbCrLf & "Tool " & tl & ": feeds from 50% to 200% do not move its flips."
            ElseIf fh > goal Then
                note = note & vbCrLf & "Tool " & tl & ": even at double feed it flips " & fh & " times - change insp_time instead."
            Else
                If fl <= goal Then
                    hi = lo                                ' slower than half: stop at half
                Else
                    For it = 1 To 16
                        mid_ = (lo + hi) / 2
                        If FlipsWithFeeds(cs, v0, Exp(mid_), trows, tn) <= goal Then hi = mid_ Else lo = mid_
                    Next
                End If
                kBest = Exp(hi)
                ' Rounding the feeds can tip it back: nudge up until it holds.
                For it = 1 To 6
                    If FlipsWithFeeds(cs, v0, kBest, trows, tn) <= goal Then Exit For
                    kBest = kBest * 1.003
                Next
                If Abs(kBest - 1) < 0.002 Then
                    note = note & vbCrLf & "Tool " & tl & ": already even - " & goal & " flips."
                Else
                    For i = 1 To cs.Count
                        x = Scaled(cs(i), v0(i), kBest, False, h)
                        If Abs(x - v0(i)) > 1E-12 Then allCells.Add cs(i): allVals.Add x
                    Next
                    note = note & vbCrLf & "Tool " & tl & ": flips " & NumText(f0) & " -> " & NumText(goal) & ",  feeds x" & _
                           Format$(kBest, "0.000") & IIf(fl <= goal And goal = f0, " (stopped at half feed)", "") & "."
                End If
            End If
        End If
    Next

    ' Every op of the tools, before and after all of it together.
    an = 0
    ReDim allRows(1 To 1)
    For Each tl In Split(Mid$(tools, 2), "|")
        trows = RowsIn(ToolRange(CStr(tl)), tn)
        For i = 1 To tn
            an = an + 1
            ReDim Preserve allRows(1 To an)
            allRows(an) = trows(i)
        Next
    Next
    s0 = Snap(allRows, an)
    s1 = TrialSnap(allCells, allVals, allRows, an)
    AddLine "op", "tool", IIf(how = "insp_time", "insp_time", "feed"), "flips / part", "time", "MRR avg"
    For i = 1 To an
        what = ""
        For j = 1 To allCells.Count
            If allCells(j).Row = allRows(i) Then
                If how = "insp_time" Then
                    what = InspText(allCells(j).Value) & "  ->  " & InspText(allVals(j))
                ElseIf what = "" Then
                    what = NumText(allCells(j).Value) & "  ->  " & NumText(allVals(j))
                End If
            End If
        Next
        If what = "" And how = "insp_time" Then
            what = InspText(ParamTable.MainSheet.Cells(allRows(i), cInsp).Value)
        End If
        AddLine OpOf(allRows(i)), ToolOf(allRows(i)), what, Arrow(s0(i, 3), s1(i, 3), False), _
                Arrow(s0(i, 1), s1(i, 1), True), Arrow(s0(i, 4), s1(i, 4), False)
    Next
    For i = 1 To allCells.Count
        pCells.Add allCells(i): pVals.Add allVals(i)
    Next
    nc = pCells.Count
    pOk = (nc > 0)
    s = "Flips / part " & Arrow(SumOf(s0, an, 3), SumOf(s1, an, 3), False) & ",  time " & _
        Arrow(SumOf(s0, an, 1), SumOf(s1, an, 1), True) & "." & note
    PlanFlips = IIf(pOk, "1|", "0|") & nc & " cell(s) will change.  " & s
End Function

Private Function InspText(ByVal v As Variant) As String
    Dim x As Double
    x = SecondsOfText(v)
    If x >= 0 Then InspText = TimeText(x)
End Function

Private Function FlipsWithInsp(ByVal cs As Collection, ByVal sec As Double, ByRef trows() As Long, ByVal tn As Long) As Double
    Dim vs As New Collection, i As Long
    For i = 1 To cs.Count
        vs.Add InspValue(cs(i), sec)
    Next
    FlipsWithInsp = SumOf(TrialSnap(cs, vs, trows, tn), tn, 3)
End Function

Private Function FlipsWithFeeds(ByVal cs As Collection, ByVal v0 As Collection, ByVal k As Double, ByRef trows() As Long, ByVal tn As Long) As Double
    Dim vs As New Collection, i As Long, h As String
    For i = 1 To cs.Count
        vs.Add Scaled(cs(i), v0(i), k, False, h)
    Next
    FlipsWithFeeds = SumOf(TrialSnap(cs, vs, trows, tn), tn, 3)
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

' Every edit on the sheet now: a Collection of Array(op_idn, column name, value).
Private Function CurrentEdits() As Collection
    Dim ws As Worksheet, out As New Collection, r As Long, c As Long, dr As Long, lc As Long, hdr As String, mc As Range
    Set CurrentEdits = out
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dumped")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lc = ParamTable.LastCol
    For r = FIRST_ROW To ParamTable.LastRow
        dr = ParamTable.DumpedRow(ParamTable.MainSheet.Cells(r, 1).Value)
        If dr > 0 Then
            For c = 1 To lc
                hdr = CStr(ParamTable.MainSheet.Cells(HEADER_ROW, c).Value)
                Set mc = ParamTable.MainSheet.Cells(r, c)
                If hdr <> "" And Not ParamTable.Untracked(hdr) And Not mc.HasFormula Then
                    If Not ParamTable.SameValue(mc.Value, ws.Cells(dr, c).Value) Then
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

' The totals of the sheet now: cycle time | cut time | flips | removed (sums).
Private Function TotalsNow(ByVal ws As Worksheet) As Variant
    Dim out(0 To 3) As Double, names As Variant, i As Long, c As Long, r As Long, x As Double, last As Long
    names = Array("est_seconds", "cut_seconds_est", "flips_part", "removed")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 0 To 3
        c = ParamTable.ColOf(CStr(names(i)))
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

' Put the sheet back to the dump, then the scenario's values in through each cell's rule.
' "set refused reverted".
Public Function RestoreScenario(ByVal name As String) As String
    Dim e As Collection, v As Variant, f As Range, c As Long, ok As Long, bad As Long, back As Long
    If Not ScenarioExists(name) Then RestoreScenario = "0 0 0": Exit Function
    Set e = ScenarioCells(name)
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
                          IIf(EditsSaved(), "", vbCrLf & "Your edits now are not in any scenario - Cancel and Save them first to keep them.")
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
    Dim sc As Long, cmt As Long, f As Range, dr As Long, hdrCol As Long, parts() As String
    names = ScenarioNames()
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
        If dr > 0 And col > 0 Then out.Cells(r, 5).NumberFormat = "@": out.Cells(r, 5).Value = CStr(d.Cells(dr, col).Value)
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

' The ops of a row's tool (keys and list captions) for the window.
Public Function ToolOps(ByVal r As Long, ByRef keys As Variant) As Variant
    Dim tool As String, rr As Long, items As String, k As String, cmt As Long
    LoadCols
    tool = ToolOf(r)
    cmt = ParamTable.ColOf("comment")
    For rr = FIRST_ROW To ParamTable.LastRow
        If ToolOf(rr) = tool And tool <> "" Then
            k = k & IIf(k = "", "", vbLf) & OpOf(rr)
            items = items & IIf(items = "", "", vbLf) & "op " & OpOf(rr) & "    " & ParamTable.MainSheet.Cells(rr, 2).Value & _
                    IIf(cmt > 0, "    " & ParamTable.MainSheet.Cells(rr, IIf(cmt > 0, cmt, 1)).Value, "")
        End If
    Next
    If k = "" Then keys = Empty: ToolOps = Empty: Exit Function
    keys = Split(k, vbLf)
    ToolOps = Split(items, vbLf)
End Function

Public Function ToolOfRow(ByVal r As Long) As String
    LoadCols
    ToolOfRow = ToolOf(r)
End Function

' ============================================================ round-insert chip thinning

' A round insert's entering angle at a depth: the chip is thinner than the feed by sin of
' it, so for a chip of h the feed per rev is h / sin.  sin = sqrt(1 - ((D - 2ap) / D)^2)
' while ap < D / 2; past half the insert it is a full-width chip (1).
Public Function ChipFactor(ByVal dia As Double, ByVal depth As Double) As Double
    Dim c As Double
    If dia <= 0 Or depth <= 0 Then Exit Function
    If depth >= dia / 2 Then ChipFactor = 1: Exit Function
    c = (dia - 2 * depth) / dia
    ChipFactor = Sqr(1 - c * c)
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

' Hit a target time, the inspection window, To all ops of tool and Restore scenario are
' in Panel.bas (undo around each).
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

' ============================================================ the plan window

Public Function PlanWindow(ByVal mode As String, ByVal sel As Range) As PlanBox
    Dim f As PlanBox
    Set f = New PlanBox
    f.Setup mode, sel, ScopeChoices(sel), ScopeDefault(sel)
    Set PlanWindow = f
End Function

Public Sub ShowPlanWindow(ByVal mode As String)
    Dim f As PlanBox, r As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Worksheet.Name <> ParamTable.MainSheet.Name Then
        MsgBox "Select operations (cells in their rows) on the '" & ParamTable.MainSheet.Name & "' sheet first.", vbInformation, "Parameter Table"
        Exit Sub
    End If
    Set f = PlanWindow(mode, Selection)
    f.Show
    If f.Accepted Then
        r = ApplyPlan()
        Application.StatusBar = "Parameter Table:  " & Split(r, " ")(0) & " cell(s) changed" & _
                                IIf(Split(r, " ")(1) <> "0", ", " & Split(r, " ")(1) & " refused by their cell's rule", "") & "."
    End If
    Unload f
End Sub

' For tools\check_macros.ps1 (no window): set the window's inputs, read back
' "summary|OK on|list rows".
Public Function PlanSelfTest(ByVal mode As String, ByVal sel As Range, ByVal scopeIdx As Long, ByVal target As String, _
                             ByVal howIdx As Long, ByVal opt1 As Boolean, ByVal opt2 As Boolean) As String
    Dim f As PlanBox
    Set f = PlanWindow(mode, sel)
    f.SetInputs scopeIdx, target, howIdx, opt1, opt2
    PlanSelfTest = f.Summary() & "|" & f.CanAccept() & "|" & f.ListRows()
    Unload f
End Function
