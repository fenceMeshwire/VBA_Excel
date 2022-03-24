Option Explicit

Sub createDictionary()

Dim dictDictionary As Object
Dim lngCellFree as Long
Dim objKeysDictionary, objValuesDictionary
Dim wsf As WorksheetFunction
Dim wksSheet As Worksheet

' Set a worksheet to visualize the dictionary's output
Set wksSheet = Tabelle1

Set wsf = Application.WorksheetFunction

Set dictDictionary = CreateObject("Scripting.Dictionary")
' Add keys and values to a dictionary.
dictDictionary.Add "key1", "value1"
dictDictionary.Add "key2", "value2"

' Separate keys and values for output
objKeysDictionary = dictDictionary.keys
objValuesDictionary = dictDictionary.items

' Visualize the output with the "Transpose" function
With wksSheet
  lngCellFree = 2 + Ubound(objKeysDictionary)
  .Range("A1").Value = "Keys"
  .Range("B1").Value = "Values"
  .Range("A2:A" & lngCellFree = wsf.Transpose(dictDictionary.keys)
  .Range("B2:B" & lngCellFree = wsf.Transpose(dictDictionary.items)
End With

End Sub
