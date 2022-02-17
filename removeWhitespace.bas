Option Explicit

Sub removeWhitespaceCharacters()

'Trim removes unwanted whitespace characters on the left and right margins
'LTrim removes unwanted whitespace characters on the left margins
'RTrim removes unwanted whitespace characters on the right margins

Dim lngRow, lngRowMax As Long

With Sheet1
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 1 To lngRowMax
    'Remove whitespace characters for Range("A" & lngRow)
    .Range("A" & lngRow).Value = Trim(.Range("A" & lngRow).Value)
  Next lngRow
End With

End Sub
