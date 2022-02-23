Option Explicit

Sub FindCellDeleteColumnsNotIncluded()

' Purpose: Find a search term in a specific cell and delete all columns
' which do not have any values stored regarding to the search term.
Dim lngRow, lngRowMax As Long
Dim lngColumn, lngColumnMax As Long
Dim lngColumnSearchTerm, lngRowSearchTerm As Integer
Dim strSearchTerm As String

strSearchTerm = "search term"

With Sheet1
  ' Measure the limits of the used range of the spreadsheet
  lngColumnMax = .UsedRange.Columns.Count
  lngRowMax = .UsedRange.Rows.Count
  
  ' Determine row and column with the designation of strSearchTerm
  ' in order to get the cell address.
  For lngRow = 2 To lngRowMax
    For lngColumn = 1 To lngColumnMax
      If .Cells(lngRow, lngColumn).Value = strSearchTerm Then
        lngColumnSearchTerm = lngColumn
        lngRowSearchTerm = lngRow
      End If
    Next lngColumn
  Next lngRow
  If lngColumnSearchTerm = 0 Then
    MsgBox ("No match with the search term found.")
    Exit Sub
  End If
  
  ' Delete columns which are not including any value.
  For lngColumn = lngColumnMax To 1 Step -1
    If .Cells(lngRowSearchTerm, lngColumn) = "" Then
      .Columns(lngColumn).Delete
    End If
  Next lngColumn
  
End With

End Sub

