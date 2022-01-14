Option Explicit

Sub filterNames()

Sheet1.Range("A:D").Sort Key1:=Sheet1.Range("A1"), order1:=xlAscending, Header:=xlYes

End Sub
