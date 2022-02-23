Option Explicit

Sub DeleteOlderFiles()

' Date, integer and string variables
Dim dteDateFile As Date
Dim dteDateDelete As Date
Dim intDays As Integer
Dim strFileName As String
Dim strDirectory As String

' Object variables
Dim oFSO As Object
Dim oFile As Object
Dim oDirectory As Object

strDirectory = ThisWorkbook.Path
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oDirectory = oFSO.getfolder(strDirectory)

For Each oFile In oDirectory.Files
    strFileName = oFile
    If Dir(strFileName) <> "" Then
        intDays = 365 ' Files which are older than a given period of time in days will be deleted.
        dteDateFile = Format(FileDateTime(strFileName), "00000")
        dteDateDelete = Format(Now - intDays, "00000")
        If dteDateFile < dteDateDelete Then
            Kill strFileName
        End If
    End If
Next oFile

End Sub
