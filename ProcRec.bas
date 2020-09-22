Attribute VB_Name = "ProcRec"
Global UserState(MaxUsers) As Integer

Function Proccess(strData As String, Socket As Integer)
On Error Resume Next
If Mud.WizLocked = True Then
nul = Send("The whole mud is covered in ice, and you cannot move.", Socket)
Exit Function
End If

Dim ClearAfter As Boolean

    If strData = vbCrLf Then
        nul = Parse(Buffer(Socket), " ")
        If Users.Users(Socket).LastCommand = Buffer(Socket) Then
            If Users.Users(Socket).LastCommandCount > 18 Then
                nul = Logit("TheASP was spamming the following command: " & Buffer(Socket) & ". He\She was disconnected.", 3, "spam")
                frmMain.WS(Socket).SendData "PUT A LID ON IT!!!"
                frmMain.WS(Socket).Close
            Else
            Users.Users(Socket).LastCommandCount = Users.Users(Socket).LastCommandCount + 1
            End If
        Else
            Users.Users(Socket).LastCommand = Buffer(Socket)
        End If
        ClearAfter = True
    ElseIf strData = Chr(8) Then
        Buffer(Socket) = Left$(Buffer(Socket), Len(Buffer(Socket)) - 1)
        Exit Function
    Else
        Buffer(Socket) = Buffer(Socket) & strData
        Exit Function
    End If

If UserState(Socket) = 0 Then
nul = Logit("Comm: " & Buffer(Socket) & " has entered their name, requesting password...", 3, "comm")
Users.Users(Socket).strBuf1 = Buffer(Socket)
UserState(Socket) = 1
frmMain.WS(Socket).SendData "What is thy word of passage? " & vbCrLf
ElseIf UserState(Socket) = 1 Then
Users.Users(Socket).strBuf2 = Buffer(Socket)
UserState(Socket) = 3
nul = Users.Auth(Socket, Users.Users(Socket).strBuf1, Users.Users(Socket).strBuf2)
ElseIf UserState(Socket) = -1 Then
frmMain.WS(Socket).Close
ElseIf UserState(Socket) = 3 Then
frmMain.WS(Socket).SendData "Please wait."
ElseIf UserState(Socket) = 4 Then

Select Case LCase(Results(0))

    Case "quit"
        Do_Quit (Socket)

    Case "reload"
        If Not IsEmpty(Results(1)) Then
        nul = Do_Reload(Socket, Results(1))
        Else
        nul = Send("Reload what?", Socket)
        End If
        
    Case "mud"
        If Not IsEmpty(Results(1)) Then
        nul = Do_Mud(Socket, Results(1))
        Else
        nul = Send("The following arguments are supported by the mud command:" & vbCrLf & "name" & vbCrLf & "version", Socket)
        End If

    Case "look"
        Do_Look (Socket)
        
    Case "n"
        Do_North (Socket)
        
    Case "north"
        Do_North (Socket)
        
    Case "userstate"
        Do_UserState (Socket)
        
    Case "socketstate"
        Do_SocketState (Socket)
        
    Case "help"
        If Not IsEmpty(Results(1)) Then
            Do_Help (Socket)
        Else
            nul = Send("What would you like help on?", Socket)
        End If
        
    Case "e"
        Do_East (Socket)
        
    Case "score"
        Do_Score (Socket)
    
    Case "sc"
        Do_Score (Socket)
        
    Case "chat"
        Do_Chat (Socket)
        
    Case "ch"
        Do_Chat (Socket)
        
    Case "."
        Do_Chat (Socket)
            
    Case "say"
        Do_Say (Socket)
        
    Case "'"
        Do_Say (Socket)
        
    Case "east"
        Do_East (Socket)
        
    Case "s"
        Do_South (Socket)
        
    Case "south"
        Do_South (Socket)
        
    Case "w"
        Do_West (Socket)
        
    Case "west"
        Do_West (Socket)
        
    Case "l"
        Do_Look (Socket)

    Case ""
        nul = Send(" ", Socket)
        Buffer(Socket) = ""
        
    Case "recall"
    If Not ParseMod.Count > 0 Then
    nul = Do_Recall(Socket, "")
    Else
    nul = Do_Recall(Socket, Results(1))
    End If
    
    Case "rec"
    If Not ParseMod.Count > 1 Then
    nul = Do_Recall(Socket, "")
    Else
    nul = Do_Recall(Socket, Results(1))
    End If

    Case Else
        If Socials.Check(Results(0), Socket) = False Then
        nul = Send("Huh?", Socket)
        End If
        
End Select

End If
    
If ClearAfter = True Then
Buffer(Socket) = ""
End If

End Function
