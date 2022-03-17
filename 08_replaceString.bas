Option Explicit

Sub replaceString()  
  
Dim lngRow, lngRowMax, lngColumn, lngColumnMax As Long
Dim intColumnTerm As Integer
Dim strKeyword As String

With Sheet1
  
  ' Find keyword in the relevant column
  lngColumnMax = .UsedRange.Columns.Count
  For lngColumn = 1 To lngColumnMax
    strKeyword = .Cells(1, lngColumn).Value
    If strKeyword = "Keyword" Then
      intColumnTerm = lngColumn
    End If
  Next lngColumn
  
  ' Replace the desired string in every row of the document's used range:
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strKeyword = .Cells(lngRow, intColumnTerm).Value
    strKeyword = Replace(strKeyword, "_", " ")
    .Cells(lngRow, intColumnTerm).Value = strKeyword
  Next lngRow
  
End With

End Sub
