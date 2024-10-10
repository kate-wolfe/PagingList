Attribute VB_Name = "ClearModule"
Option Explicit

Sub Clear()

Dim WS_Count As Integer
Dim i As Integer

WS_Count = ThisWorkbook.Worksheets.Count

For i = 1 To WS_Count
    ThisWorkbook.Worksheets(i).Columns("A:L").Delete
Next i

End Sub
