Option Explicit

Sub CheckDateInput()

Dim bleChkDate As Boolean
Dim dteDate As Date
Dim strDate As String

strDate = InputBox("Please enter a date", "Date collection", "01.05.2022")

If Len(strDate) = 0 Then ' Abort criterion when pressing cancel
  MsgBox "Action was cancelled by the user."
  Exit Sub
Else
  bleChkDate = IsDate(strDate) 'Plausibility check for an entered date
  If bleChkDate = True Then
    dteDate = CDate(strDate)
  Else
    MsgBox "Please enter a valid date (DD.MM.YYYY)."
    Exit Sub
  End If
End If

MsgBox ("The following date was entered: " & dteDate)

End Sub
