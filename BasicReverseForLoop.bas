Option Explicit

Sub BasicReverseForLoop()

Dim lngRow, lngRowMax As Long
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

With wksSheet
  lngRowMax = wksSheet.UsedRange.Rows.Count
  For lngRow = lngRowMax To 1 Step -1
     ' Perform any action, usually delete operation of cells or rows
  Next lngRow
End With

End Sub
