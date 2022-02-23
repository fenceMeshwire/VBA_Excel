Option Explicit

Sub frameRange()

Dim rngRange As Range
Dim lngRowMax As Long

lngRowMax = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row

Set rngRange = Sheet1.Range("A1:D" & lngRowMax)

With rngRange
  .Borders(xlInsideHorizontal).LineStyle = xlContinuous
  .Borders(xlInsideVertical).LineStyle = xlContinuous
  .BorderAround Weight:=xlThin, ColorIndex:=1
End With

End Sub
