Attribute VB_Name = "Checks"
' Parameter Table Tool - the continuous-improvement tools: Program check (what to look
' at before the program goes to the floor), Slowest ops (where the time is), and the
' Change report (before -> after, for the record).
'
' Plain text in the repo (vba\Checks.bas); tools\build_vba.ps1 compiles every vba\*.bas.
Option Explicit

Private Const TITLE As String = "Parameter Table"

Public Sub RbCheck(control As IRibbonControl)
    MsgBox "Program check is coming in the next version.", vbInformation, TITLE
End Sub

Public Sub RbSlowest(control As IRibbonControl)
    MsgBox "Slowest ops is coming in the next version.", vbInformation, TITLE
End Sub

Public Sub RbReport(control As IRibbonControl)
    MsgBox "The change report is coming in the next version.", vbInformation, TITLE
End Sub
