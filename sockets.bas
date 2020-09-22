Attribute VB_Name = "Sockets"
Global Port As Integer
Global Const MaxUsers As Integer = 20
Global Const DevMode As Boolean = True
Global Running As Boolean
Public SocketState(MaxUsers) As Integer
Public InitDone As Boolean

Function Init()
If InitDone = True Then Exit Function
Port = 2002
For I = 1 To MaxUsers
Load frmMain.WS(I)
Next
End Function

Function StartServer()
On Error GoTo Failure
Errors = 0
nul = Logit("Attempting to start the server...", 2)
frmMain.WSListen.LocalPort = Port
frmMain.WSListen.Listen
Running = True
If Errors > 0 Then
nul = Logit("Server May Have Been Started, " & Errors & " occured while attempting.", 2)
Else
nul = Logit("Server Started, No Errors.", 2)
End If
Exit Function
Failure:
Errors = Errors + 1
End Function

Function StopServer()
On Error GoTo Failure
Errors = 0
nul = Logit("Stopping the Server...", 2)
frmMain.WSListen.Close
For I = 0 To MaxUsers
If SocketState(I) = 1 Then
frmMain.WS(I).SendData "Server shut down by the system administrator."
frmMain.WS(I).Close
ElseIf SocketState(I) = 2 Then
frmMain.WS(I).SendData "The server has been shut down by the system administrator, your player file has not been saved."
End If
Next
If Errors > 0 Then
nul = Logit("The server may have been stopped, " & Errors & " occured while attempting to stop it.", 2)
Else
nul = Logit("The server was stopped with no errors.", 2)
End If
Running = False
Exit Function
Failure:
Errors = Errors + 1
End Function

Function Connect(requestID As Long)
Dim I As Integer

For I = 0 To MaxUsers
If SocketState(I) = 0 Then
frmMain.WS(I).Accept requestID
nul = Logit("Comm: " & frmMain.WS(I).RemoteHostIP & " has connected.", 3, "comm")
nul = SendSpec("greeting", I)
SocketState(I) = 1
Exit Function
End If
Next

nul = Logit("All sockets are full. A connection was turned down.", 3, "comm")

End Function

Function CloseSock(Socket As Integer)
nul = Logit("Comm: " & Users.Users(Socket).Username & " has lost his\her connection.", 3, "comm")
End Function

Function SendSpec(SendWhat As String, Socket As Integer)
Select Case SendWhat

Case "greeting"
'frmMain.WS(Socket).SendData Replace(SAnsi("white") & Tables.HelpContents(0) & vbCrLf, vbCrLf, Chr(13))
frmMain.WS(Socket).SendData SAnsi("white") & Tables.HelpContents(0) & vbCrLf

End Select

End Function

Function Send(Message As String, Socket As Integer, Optional DoLook As Boolean = False)
On Error Resume Next
If Users.Users(Socket).Level > 800 Then
Prompt = Ansi(7) & "<" & Users.Users(Socket).HP & "/" & Users.Users(Socket).HPMax & " hp><" & Users.Users(Socket).Mana & "/" & Users.Users(Socket).ManaMax & " mana><" & Users.Users(Socket).Moves & "/" & Users.Users(Socket).MovesMax & " moves><vnum: " & Users.Users(Socket).Vnum & ">" & vbCrLf
Else
Prompt = Ansi(7) & "<" & Users.Users(Socket).HP & "/" & Users.Users(Socket).HPMax & " hp><" & Users.Users(Socket).Mana & "/" & Users.Users(Socket).ManaMax & " mana><" & Users.Users(Socket).Moves & "/" & Users.Users(Socket).MovesMax & " moves>" & vbCrLf
End If
'frmMain.WS(Socket).SendData Replace(Message & vbCrLf & vbCrLf & Prompt, vbCrLf, Chr(13))
frmMain.WS(Socket).SendData Message & vbCrLf & vbCrLf & Prompt
'If DoLook = True Then Commands.Do_Look (Socket)
End Function
