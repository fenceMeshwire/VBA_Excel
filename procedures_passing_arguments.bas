Option Explicit

Sub get_data()

Dim str_name, str_address, str_phone_no As String
Dim dte_birthdate As Date
Dim str_email As String

str_name = "Username"
str_address = "1445 West Norwood Avenue"
str_phone_no = "+49-123-459"
dte_birthdate = "23.05.2019"
str_email = "user.name@domain.com"

Call transfer_data(str_name, str_address, str_phone_no, _
  dte_birthdate, str_email)

End Sub

' Note keyword "optional"
Sub transfer_data(name, address, Optional phone_no, _
Optional birthdate, Optional email)

Debug.Print name
Debug.Print address
Debug.Print phone_no
Debug.Print birthdate
Debug.Print email

End Sub
