Option Explicit

Sub compareArrayForDuplicates()

Dim lngRow, lngRowMax As Long
Dim varArray, varArrayCompare As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

With wksSheet

  .Range("C:C").ClearContents
  ' Determine the last used row of Sheet1
  lngRowMax = .Range("A" & .Rows.Count).End(xlUp).Row
  ' Read data from Sheet1 into varArray
  varArray = .Range(.Cells(1, "A"), .Cells(lngRowMax, "B")).Value
  ReDim varArrayCompare(LBound(varArray, 1) To UBound(varArray, 1))
  
  For lngRow = LBound(varArray, 1) To UBound(varArray, 1)
    If varArray(lngRow, 1) = varArray(lngRow, 2) Then
      varArrayCompare(lngRow) = "Duplicate"
    End If
  Next lngRow

  .Cells(1, 3).Resize(UBound(varArrayCompare, 1)).Value = _
  Application.Transpose(varArrayCompare)

End With

End Sub
