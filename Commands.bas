Attribute VB_Name = "Commands"
Function Do_Quit(Socket As Integer)
        nul = Logit("Comm: " & Users.Users(Socket).Username & " has quit.", 3, "comm")
        nul = Send("Good bye!", Socket)
        frmMain.WS(Socket).Close
End Function

Function Do_Look(Socket As Integer)
On Error GoTo Err

Dim ToSend As String

ToSend = Ansi(14) & World.RoomName(Users.Users(Socket).Vnum) & Ansi(7) & vbCrLf
ToSend = ToSend & World.RoomDesc(Users.Users(Socket).Vnum) & vbCrLf
Exits = Ansi(3) & "[Exits: "

If World.RoomExits(Users.Users(Socket).Vnum, 0) > 0 Then
Exits = Exits & " north "
End If
If World.RoomExits(Users.Users(Socket).Vnum, 1) > 0 Then
Exits = Exits & " east "
End If
If World.RoomExits(Users.Users(Socket).Vnum, 2) > 0 Then
Exits = Exits & " south "
End If
If World.RoomExits(Users.Users(Socket).Vnum, 3) > 0 Then
Exits = Exits & " west "
End If

Exits = Exits & "]" & Ansi(7)
ToSend = ToSend & Exits & vbCrLf

Buf = Ansi(14)

For I = 0 To MaxUsers
If Users.Users(I).Vnum = Users.Users(Socket).Vnum Then
If Not I = Socket Then
Buf = Buf & vbCrLf & Users.Users(I).Username & " " & Users.Users(I).Title
End If
End If
Next
ToSend = ToSend & Buf & Ansi(7)

nul = Send(ToSend, Socket)

Exit Function
Err:
Users.Users(Socket).Vnum = 1
Resume Next
End Function

Function Do_Score(Socket As Integer)
Dim Buf As String

Buf = Ansi(10) & "Score Sheet for " & Ansi(12) & Users.Users(Socket).Username & Ansi(10) & vbCrLf & "------------------------------------------------------------------------" & vbCrLf

nul = Send(Buf, Socket)

End Function

Function Do_Help(Socket As Integer)
Dim Buf As String
For I = 1 To ParseMod.Count
Buf = Buf & " " & Results(I)
Next

If Buf = "" Then
nul = Send("Find a help file on what?", Socket)
ElseIf Buf = " " Then
nul = Send("Find a help file on what?", Socket)
ElseIf Buf = "  " Then
nul = Send("Find a help file on what?", Socket)
End If

Buf = Right$(Buf, Len(Buf) - 1)

For I = 0 To Tables.HelpCount
If Left(Tables.Help(I), Len(Buf)) = Buf Then
If Not Tables.HelpLvl(I) > Users.Users(Socket).Level Then
nul = Send(Ansi(15) & Tables.HelpContents(I) & Ansi(7), Socket)
Exit Function
Else
nul = Send("No help file found.", Socket)
End If
End If
Next

nul = Send("No help file found.", Socket)

End Function

Function Do_North(Socket As Integer)
If World.RoomExits(Users.Users(Socket).Vnum, 0) > 0 Then
Users.Users(Socket).Vnum = World.RoomExits(Users.Users(Socket).Vnum, 0)
nul = Send("You leave north.", Socket)
Do_Look (Socket)
Else
nul = EchoAround(PCase(Users.Users(Socket).Username) & " attempts to leave north, but relized only after " & GetHeShe(Socket) & " has walked into the wall, that there is no exit there.", Socket, -1)
nul = Send("You attempt to leave north and relize, only after you've walked into the wall, that there is no exit there.", Socket)
End If
End Function

Function Do_East(Socket As Integer)
If World.RoomExits(Users.Users(Socket).Vnum, 1) > 0 Then
Users.Users(Socket).Vnum = World.RoomExits(Users.Users(Socket).Vnum, 1)
nul = Send("You leave east.", Socket)
Do_Look (Socket)
Else
nul = EchoAround(PCase(Users.Users(Socket).Username) & " attempts to leave east, but relized only after " & GetHeShe(Socket) & " has walked into the wall, that there is no exit there.", Socket, -1)
nul = Send("You attempt to leave east and relize, only after you've walked into the wall, that there is no exit there.", Socket)
End If
End Function

Function Do_South(Socket As Integer)
If World.RoomExits(Users.Users(Socket).Vnum, 2) > 0 Then
Users.Users(Socket).Vnum = World.RoomExits(Users.Users(Socket).Vnum, 2)
nul = Send("You leave south.", Socket)
Do_Look (Socket)
Else
nul = EchoAround(PCase(Users.Users(Socket).Username) & " attempts to leave south, but relized only after " & GetHeShe(Socket) & " has walked into the wall, that there is no exit there.", Socket, -1)
nul = Send("You attempt to leave south and relize, only after you've walked into the wall, that there is no exit there.", Socket)
End If
End Function

Function Do_West(Socket As Integer)
If World.RoomExits(Users.Users(Socket).Vnum, 3) > 0 Then
Users.Users(Socket).Vnum = World.RoomExits(Users.Users(Socket).Vnum, 3)
nul = Send("You leave west.", Socket)
Do_Look (Socket)
Else
nul = EchoAround(PCase(Users.Users(Socket).Username) & " attempts to leave west, but relized only after " & GetHeShe(Socket) & " has walked into the wall, that there is no exit there.", Socket, -1)
nul = Send("You attempt to leave west and relize, only after you've walked into the wall, that there is no exit there.", Socket)
End If
End Function

Function Do_Recall(Socket As Integer, Arguement As String)
Select Case Arguement
Case "reset"
Users.Users(Socket).Recall = Users.Users(Socket).RecallDefault
nul = Send("Recall Reset.", Socket)
Case "set"
Users.Users(Socket).Recall = Users.Users(Socket).Vnum
Case ""
If Not Users.Users(Socket).Recall = Users.Users(Socket).Vnum Then
Users.Users(Socket).Vnum = Users.Users(Socket).Recall
Else
nul = Send("You're already there!", Socket)
End If
End Select

End Function

Function Do_UserState(Socket As Integer)
If Users.Users(Socket).Level >= 800 Then
nul = Send("Your userstate is: " & UserState(Socket), Socket)
Else
nul = Send("Huh?", Socket)
End If
End Function

Function Do_SocketState(Socket As Integer)
If Users.Users(Socket).Level > 800 Then
nul = Send("Your socketstate is: " & SocketState(Socket), Socket)
Else
nul = Send("Huh?", Socket)
End If
End Function

Function Do_Say(Socket As Integer)
For I = 1 To ParseMod.Count
Buf = Buf & " " & Results(I)
Next
nul = EchoAround(Ansi(9) & Users.Users(Socket).Username & " says '" & Ansi(15) & Buf & Ansi(9) & "'" & Ansi(7), Socket, -1)
nul = Send(Ansi(9) & "You say '" & Ansi(15) & Buf & Ansi(9) & "'" & Ansi(7), Socket)
End Function

Function Do_Chat(Socket As Integer)
Dim I As Integer

For I = 1 To ParseMod.Count
Buf = Buf & " " & Results(I)
Next

For I = 0 To MaxUsers
If UserState(I) = 4 Then
nul = Send(Ansi(14) & Users.Users(Socket).Username & " chats '" & Buf & "'" & Ansi(7), I, True)
End If
Next

nul = Send(Ansi(14) & "You chat '" & Buf & "'" & Ansi(7), Socket)
End Function

'Immortal Only, need to add support for the command table

Function Do_Reload(Socket As Integer, Argument As String)
If Users.Users(Socket).Level < 980 Then
nul = Send("Huh?", Socket)
Else
Select Case LCase(Argument)
Case "world"
nul = Send("Reloading World...", Socket)
Mud.WizLocked = True
World.LoadWorld (World.WorldFileName)
frmMain.Wait (3)
Mud.WizLocked = False
nul = Send("Complete!", Socket)
End Select
End If
End Function

Function SEcho(Message As String, Source As Integer)
Dim I As Integer

For I = 0 To MaxUsers
If UserState(I) = 4 Then
If Not Users.Users(Source).Username = Users.Users(I).Username Then
nul = Send(Message, I)
End If
End If
Next

End Function

Function EchoAround(Message As String, TargetSock As Integer, Optional Source As Integer = -1)
If Not Source = -1 Then
If Users.Users(Source).Level > 840 Then
nul = Logit(Users.Users(Source).Username & ": echoaround " & Message & " " & TargetSock, 3, "cmds")
Else
nul = Send("Huh?", Source)
Exit Function
End If
End If

Dim I As Integer
For I = 0 To MaxUsers
If UserState(I) = 4 Then
If Users.Users(I).Vnum = Users.Users(TargetSock).Vnum Then
If Not Users.Users(I).Username = Users.Users(TargetSock).Username Then
nul = Send(Message, I, True)
End If
End If
End If
Next

End Function

Function Do_Mud(Socket As Integer, Arg As String)
Select Case LCase(Arg)
Case "name"
nul = Send(Ansi(15) & "Your are playing on: " & Ansi(12) & Mud.Name & Ansi(7), Socket)
Case "version"
nul = Send(Ansi(15) & "Mud is currently version: " & Ansi(12) & Mud.Version & Ansi(7), Socket)
Case Else
nul = Send(Ansi(15) & "Valid arguments are: name\version" & Ansi(7), Socket)
End Select
End Function
