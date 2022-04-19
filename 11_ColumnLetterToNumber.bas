Option Explicit

Sub ColumnLetterToNumber()

Dim strFeld As String

strFeld = InputBox("Please enter the column letter, like the default value", _
  "Convert a column letter to a column number", "A")
  
If strFeld = "" Then
  Exit Sub
End If

If Len(strFeld) > 0 Then
    If IsNumeric(strFeld) = False Then
        MsgBox "Spalte: " & Columns(strFeld).Column
    End If
End If

End Sub
