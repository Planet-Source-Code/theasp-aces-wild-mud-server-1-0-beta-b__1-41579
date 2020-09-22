Attribute VB_Name = "Logs"

Function Logit(Message As String, Optional LogTo As Integer = 1, Optional LogFile As String = "system")

Select Case LogTo
Case 1
nul = LogToFile(Message, LogFile)

Case 2
nul = LogToFile(Message, LogFile)
LogToScreen (Message)

Case 3
nul = LogToFile(Message, LogFile)
LogToScreen (Message)
LogToMud (Message)

End Select

End Function

Public Function LogToFile(Message As String, LogTo As String)

Open App.Path & "\logs\" & LogTo & ".log" For Append As #2
Print #2, "{" & Date & " : " & Time & "} " & Message
Close #2

End Function

Public Function LogToScreen(Message As String)
frmMain.lstLog.AddItem ("{" & Date & " : " & Time & "} " & Message)
End Function

Public Function LogToMud(Message As String)
Dim I As Integer
For I = 0 To MaxUsers
If IsImmort(I) = True Then
nul = Send(Message, I)
End If
Next
End Function
