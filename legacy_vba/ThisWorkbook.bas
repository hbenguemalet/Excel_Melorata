Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Sheets("Sommaire").Activate
    On Error GoTo 0
End Sub
