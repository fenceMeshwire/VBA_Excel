Option Explicit

Sub BasicForLoop()

Dim lngRow, lngRowMax As Long
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

With wksSheet
  lngRowMax = wksSheet.UsedRange.Rows.Count
  For lngRow = 1 To lngRowMax
    ' Perform any action
  Next lngRow
End With

End Sub
