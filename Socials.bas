Attribute VB_Name = "Socials"
Public SocName(20) As String
Public SocSelf(20) As String
Public SocOther(20) As String

Function Init()
SocName(0) = "nod"
SocSelf(0) = "You nod."
SocOther(0) = "<name> nods quickly."

SocName(1) = "sigh"
SocSelf(1) = "You sigh."
SocOther(1) = "<name> sighs."

SocName(2) = "hrm"
SocSelf(2) = "You make a 'Hrm' sound."
SocOther(2) = "<name> makes a 'Hrm' like sound."

End Function

Function Check(szIn As String, Socket As Integer) As Boolean
Debug.Print szIn
For I = 0 To 3
If LCase(Left(SocName(I), Len(szIn))) = LCase(szIn) Then
nul = Send(SocSelf(I), Socket)
nul = EchoAround(Replace(SocOther(I), "<name>", Users.Users(Socket).Username), Socket, -1)
Check = True
Exit Function
End If
Next

Check = False

End Function
