Option Explicit

Type Employee
  name As String
  account As Integer
  phone() As String
  email(1 To 2) As String
  wage As Currency
End Type

Sub user_defined_datatype()
  Dim entity As Employee
  Dim group(1 To 5) As Employee
  
  ' Assign values to variables
  entity.name = "John Smith"
  entity.account = 1234
  entity.email(1) = "john.smith@company.org"
  entity.email(2) = "john.smith@private.com"
  entity.wage = 4232.42
  
  ' Sizing an undefined variable
  ReDim entity.phone(1 To 2)
  entity.phone(1) = "401-12394"
  entity.phone(2) = "401-12395"
  
  group(1) = entity
  
  Debug.Print group(1).name & "; Phone: " & group(1).phone(1)
  
End Sub
