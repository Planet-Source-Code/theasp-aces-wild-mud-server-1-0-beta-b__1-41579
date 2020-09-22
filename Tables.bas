Attribute VB_Name = "Tables"
Public CmdLoaded As Boolean
Public SklLoaded As Boolean
Public HelpLoaded As Boolean

Public Commands() As String
Public CmdLvls() As Integer

Public Skill() As String
Public SklLvl() As Integer

Public Help() As String
Public HelpContents() As String
Public HelpLvl() As Integer
Public HelpCount As Integer


Function LoadCmd()
Dim Count As Integer
Count = -1
Open App.Path & "\commands.dat" For Input As #1
Do While Not EOF(1)
Count = Count + 1
Line Input #1, Commands(Count)
Line Input #1, CmdLvl
CmdLvls(Count) = Int(CmdLvl)
Loop
Close #1
CmdLoaded = True
End Function

Function LoadSkls()
Dim Count As Integer
Count = -1
Open App.Path & "\skills.dat" For Input As #1
Do While Not EOF(1)
Count = Count + 1
Line Input #1, Skill(Count)
Line Input #1, CmdLvl
SklLvls(Count) = Int(CmdLvl)
Loop
Close #1
SklLoaded = True
End Function

Function LoadHelp()
HelpCount = -1
Dim Count As Integer
Count = -1
Open App.Path & "\helpfiles.dat" For Input As #1
Do While Not EOF(1)
Count = Count + 1
ReDim Preserve Help(Count)
ReDim Preserve HelpContents(Count)
ReDim Preserve HelpLvl(Count)
Line Input #1, Help(Count)
Line Input #1, HelpCont
HelpContents(Count) = Replace(HelpCont, "~^~", Chr(13))
Line Input #1, CmdLvl
HelpLvl(Count) = Int(CmdLvl)
HelpCount = HelpCount + 1
Loop
Close #1
HelpLoaded = True
End Function
