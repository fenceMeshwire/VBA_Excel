Option Explicit

Sub testVariables()

Dim dteDate As Date
Dim intCounter As Integer

Dim objEmpty As Object
Dim objFSO As Object
Dim objNull As Object

Dim strString As String

Dim varDat As Variant
Dim varEmpty As Variant
Dim varNull As Variant

dteDate = "01.01.2024"
intCounter = 14

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNull = Nothing

strString = "ab:3k:23:34:45:rc:de"

varDat = Split(strString, ":")
varEmpty = Empty
varNull = Null

If IsDate(dteDate) = True Then
  Debug.Print "dteDate is a date."
End If

If IsNumeric(intCounter) = True Then
  Debug.Print "intCounter is numeric."
End If

If IsObject(objFSO) Then
  Debug.Print "objFSO is an object."
End If

If IsError(varDat) = False Then
  Debug.Print "varDat is correct."
End If

If IsEmpty(varEmpty) = True Then
  Debug.Print "varEmpty is Empty."
End If

If IsNull(varNull) Then
  Debug.Print "varNull is Null."
End If

Debug.Print VarType(varDat)
Debug.Print TypeName(varDat)

End Sub
