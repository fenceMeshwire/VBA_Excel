Option Explicit

' Purpose:  1.) Add leading zeroes to a number create a five digit number. 
'           2.) Create an output file for each term with numbers from lower
'               to upper boundary.
' An output folder has to be created in the same directory level as the worksheet.

' Examples: 1 -> 00001, 10 -> 00010, 100 -> 00100, etc.

' Data structure main document:
' columns ->
'   1   2   3   4   5   6   7
'       |   |               |
'      min max            term
' 
' Example:
' min: BA02345, max: BA65432, term: KA
' output: BA02345, BA02346, ..., BA65432 

Sub RunningNumber()

Dim lngRow, lngRowMax as Long
Dim lngRunMin, lngRunMax as Long
Dim strRunMin, StrRunMax, StrRun, StrRunPre as String
Dim strTerm as String

With tbl_table
    lngRowMax = .Cells(.Rows.Count, 3).End(xlUp).Row
    For lngRow = 1 to lngRowMax
        strRunMin = .Cells(lngRow, 3).Value 
        strRunPre = Left(strRunMin, 2) ' Two character prefix
        strRunMin = Right(strRunMin, 5) ' Minimum five digit number
        strRunMax = .Cells(lngRow, 4).Value
        strRunMax = Right(strRunMax, 5) ' Maximum five digit number
        strTerm = .Cells(lngRow, 7).Value ' Term
        lngRunMin = CLng(strRunMin) ' Lower boundary
        lngRunMax = CLng(strRunMax) ' Upper boundary

        ' File output:
        Open ThisWorkbook.Path & "\Output\" & CStr(lngRow) & "_" & strTerm & ".txt" For Output as #1
        For lngRunMin = lngRunMin to lngRunMax Step 1
            Select Case Len(lngRunMin)
                Case 1: strRunMin = "0000" & CStr(lngRunMin)
                Case 2: strRunMin = "000" & CStr(lngRunMin)
                Case 3: strRunMin = "00" & CStr(lngRunMin)
                Case 4: strRunMin = "0" & CStr(lngRunMin)
                Case 5: strRunMin = Cstr(lngRunMin)
            End Select
            strRun = strRunPre & strRunMin ' Prefix + Running Number
            Print #1, strRun
        Next lngRunMin
        Close #1
    Next lngRow
End With

End Sub
