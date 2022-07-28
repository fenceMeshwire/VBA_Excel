Option Explicit

Sub collection_object()

Dim colObject As New Collection
Dim intCounter As Integer
Dim strOutput As String

' Elements of a Collection do not need to have the same data type.

colObject.Add "This"
colObject.Add "is"
colObject.Add "a"
colObject.Add "collection"

For intCounter = 1 To colObject.Count
  strOutput = strOutput & " " & colObject(intCounter)
Next intCounter

strOutput = strOutput & "."

Debug.Print strOutput

End Sub
