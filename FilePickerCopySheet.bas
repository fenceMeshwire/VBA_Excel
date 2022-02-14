Option Explicit

Sub FilePickerCopySheet()

Dim strFile As String
Dim lngColumn As Long

With Application.FileDialog(msoFileDialogFilePicker)
  .AllowMultiSelect = False
  .Title = "Please select the desired workbook."
  .InitialFileName = ThisWorkbook.Path
  .Filters.Add "Worksheets", "*.xls*", 1
  If .Show = -1 Then
      strFile = .SelectedItems(1)
  End If
End With

If strFile <> "" Then
  Workbooks.Open strFile
Else
  Exit Sub
End If

' Copy Sheet1 from the opened workbook.
ActiveWorkbook.Worksheets(1).UsedRange.Copy Destination:=Sheet1.Range("A1")
ActiveWorkbook.Close

End Sub
