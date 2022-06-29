Option Explicit

Sub execute_recursion() ' Procedure to execute the recursive function.

Debug.Print recursion(5)

End Sub

Function recursion(n As Double) As Double ' Recursive function.

If n <= 1 Then
  recursion = 1
Else
  recursion = n * recursion(n - 1)
End If

End Function
