Option Explicit

Sub concatenate_cells()

Dim lngRow, lngRowMax As Long
Dim intLBound, intUBound, intCheckSum As Integer
Dim strKey, strKeyResult As String
Dim varDat As Variant

lngRowMax = Sheet1.UsedRange.Rows.Count

For lngRow = 1 To lngRowMax
  strKey = Tabelle4.Cells(lngRow, 1).Value
  If lngRow < lngRowMax Then
    strKeyResult = strKeyResult + strKey + ", "
  Else
    strKeyResult = strKeyResult + strKey
  End If
Next lngRow

Debug.Print strKeyResult

varDat = Split(strKeyResult, ",")

intLBound = LBound(varDat)
intUBound = UBound(varDat)
intCheckSum = intUBound - intLBound + 1

Debug.Print intCheckSum

End Sub
