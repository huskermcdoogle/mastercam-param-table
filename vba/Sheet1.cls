' Parameter Table Tool - the "Lathe params" sheet's own events (its code name is Sheet1).
' tools\build_vba.ps1 puts this into the sheet's module; the work is in ParamTable.bas.
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Done
    Application.EnableEvents = False
    ParamTable.OnCellsChanged Target
Done:
    Application.EnableEvents = True
End Sub
