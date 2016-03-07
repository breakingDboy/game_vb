Attribute VB_Name = "Module1"
Dim strnum As String

   
Public Function tpl(a As Integer, b As Integer) As Integer
If a > b Then
   strnum = Str(a) + Str(b)
Else
   strnum = Str(b) + Str(a)
End If
   tpl = Val(strnum)
End Function

