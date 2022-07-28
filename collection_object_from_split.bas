Sub collection_object_from_split()

Dim colObject As New Collection
Dim intCounter As Integer
Dim strInput As String
Dim strOutput As String
Dim varDat As Variant

' Split a given string into a Variant
strInput = "This is a collection"
varDat = Split(strInput, " ")

' Transform the data into a Collection
For intCounter = LBound(varDat) To UBound(varDat)
  colObject.Add varDat(intCounter)
Next intCounter

' Transform the data back to a String
For intCounter = 1 To colObject.Count
  strOutput = strOutput & " " & colObject(intCounter)
Next intCounter

strOutput = strOutput & "."

' Show the output
Debug.Print strOutput

End Sub
