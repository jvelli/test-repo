Attribute VB_Name = "Custom"
Option Explicit

Sub Change_Reference_Style()
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

Function COUNT_SHEETS()
    COUNT_SHEETS = ThisWorkbook.Sheets.Count
End Function

Function COUNT_SELECTED_SHEETS()
    COUNT_SELECTED_SHEETS = ActiveWindow.SelectedSheets.Count
End Function

Function COUNT_WORKSHEETS()
    COUNT_WORKSHEETS = ThisWorkbook.Worksheets.Count
End Function

