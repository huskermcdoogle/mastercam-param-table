' Parameter Table Tool - the "Lathe params" sheet's own events (its code name is Sheet1).
' tools\build_vba.ps1 puts this into the sheet's module; the work is in ParamTable.bas
' and Panel.bas.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Done
    Application.EnableEvents = False
    ParamTable.OnCellsChanged Target
Done:
    Application.EnableEvents = True
End Sub

' The row mark follows the cursor; open op windows follow the ops selected.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Panel.SelectionMoved Target
End Sub

' Double-click a comment, inspection comment or manual entry text cell: the Comments &
' text window on that tab instead of in-cell editing (its Apply is what Undo takes back).
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Cells.Count <> 1 Or Target.Row < 3 Then Exit Sub
    If Not ParamTable.IsTextColumn(CStr(Me.Cells(2, Target.Column).Value)) Then Exit Sub
    Cancel = True
    ParamTable.EditText Target
End Sub
