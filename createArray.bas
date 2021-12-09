Option Explicit

Sub createArray()

Dim lngCell, lngCellMax, lngCounter As Long
Dim strString as String
Dim varArray as Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

With wksSheet 
    lngCellMax = .Cells(.Rows.Count, 1).End(xlUp).Row
    lngCounter = 0

    ReDim varArray(lngCounter)
    For lngCell = 1 To lngCellMax
    strString = .Cells(lngCell, 1).Value
    If Not IsNumeric(Application.Match(strString, varArray, 0)) Then ' No duplicates
        varArray(lngCounter) = strString
        lngCounter = lngCounter + 1
        ReDim Preserve varArray(lngCounter)
    End If
    Next lngCell

    ReDim Preserve varArray(UBound(varArray) - 1)
End With

End Sub
