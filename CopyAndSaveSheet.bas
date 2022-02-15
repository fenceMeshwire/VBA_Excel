Option Explicit

Sub CopyAndSaveSheet()

Dim strFilename As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

strFilename = "Filename"

Sheet1.Copy 'Sheet is copied as a new worksheet (*.xlsx)
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & Year(Date) & Format(Month(Date), "00") & _
    Format(Day(Date), "00") & "_" & strFilename & ".xlsx"
ActiveWorkbook.Close
Application.DisplayAlerts = True

End Sub
