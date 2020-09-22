Attribute VB_Name = "Users"
Type Money
    Diamond As Integer
    Gold As Integer
    Silver As Integer
    Bronze As Integer
    BankDiamond As Integer
    BankGold As Integer
    BankSilver As Integer
    BankBronze As Integer
End Type

Type User

    Username As String
    Password As String
    MaleFemale As Integer
    Race As Integer
    Class As Integer
    Level As Integer
    Experience As Long
    Tnl As Integer
    Money As Money
    Title As String
    Align As Integer
    Killed As Long
    Died As Integer
    Age As Integer
    HP As Long
    HPMax As Long
    Mana As Long
    ManaMax As Long
    Moves As Long
    MovesMax As Long
    Pracs As Integer
    Comments() As String
    strBuf1 As String
    strBuf2 As String
    strBuf3 As String
    intBuf1 As Integer
    intBuf2 As Integer
    intBuf3 As Integer
    LonBuf1 As Long
    LonBuf2 As Long
    Vnum As Integer
    Recall As Integer
    RecallDefault As Integer
    Position As Integer
    LastCommand As String
    LastCommandCount As Integer
    
End Type
    
Global Buffer(MaxUsers) As String
Public Users(MaxUsers) As User

Function Auth(Socket As Integer, Uname As String, PWord As String)
On Error GoTo NoFile
Dim ComCount As Integer
ComCount = -1

Open App.Path & "\users\" & Left(Uname, 1) & "\" & Uname & ".pfile" For Input As #1
Do While Not EOF(1)
Line Input #1, strBuffer
If Left(strBuffer, 1) = "#" Then
Select Case LCase(Right(strBuffer, Len(strBuffer) - 1))

Case "age"
Line Input #1, IntBuf
Users(Socket).Age = Int(IntBuf)
Case "align"
Line Input #1, IntBuf
Users(Socket).Align = Int(IntBuf)
Case "class"
Line Input #1, IntBuf
Users(Socket).Class = Int(IntBuf)
Case "died"
Line Input #1, IntBuf
Users(Socket).Died = Int(IntBuf)
Case "experience"
Line Input #1, IntBuf
Users(Socket).Experience = Int(IntBuf)
Case "hp"
Line Input #1, IntBuf
Users(Socket).HP = Int(IntBuf)
Case "hpmax"
Line Input #1, IntBuf
Users(Socket).HPMax = Int(IntBuf)
Case "killed"
Line Input #1, IntBuf
Users(Socket).Killed = Int(IntBuf)
Case "level"
Line Input #1, IntBuf
Users(Socket).Level = Int(IntBuf)
Case "malefemale"
Line Input #1, IntBuf
Users(Socket).MaleFemale = Int(IntBuf)
Case "mana"
Line Input #1, IntBuf
Users(Socket).Mana = Int(IntBuf)
Case "manamax"
Line Input #1, IntBuf
Users(Socket).ManaMax = Int(IntBuf)
Case "bankbronze"
Line Input #1, IntBuf
Users(Socket).Money.BankBronze = Int(IntBuf)
Case "bankdiamond"
Line Input #1, IntBuf
Users(Socket).Money.BankDiamond = Int(IntBuf)
Case "bankgold"
Line Input #1, IntBuf
Users(Socket).Money.BankGold = Int(IntBuf)
Case "banksilver"
Line Input #1, IntBuf
Users(Socket).Money.BankSilver = Int(IntBuf)
Case "bronze"
Line Input #1, IntBuf
Users(Socket).Money.Bronze = Int(IntBuf)
Case "diamond"
Line Input #1, IntBuf
Users(Socket).Money.Diamond = Int(IntBuf)
Case "gold"
Line Input #1, IntBuf
Users(Socket).Money.Gold = Int(IntBuf)
Case "silver"
Line Input #1, IntBuf
Users(Socket).Money.Silver = Int(IntBuf)
Case "moves"
Line Input #1, IntBuf
Users(Socket).Moves = Int(IntBuf)
Case "movesmax"
Line Input #1, IntBuf
Users(Socket).MovesMax = Int(IntBuf)
Case "password"
Line Input #1, Users(Socket).Password
Case "pracs"
Line Input #1, IntBuf
Users(Socket).Pracs = Int(IntBuf)
Case "race"
Line Input #1, IntBuf
Users(Socket).Race = Int(IntBuf)
Case "title"
Line Input #1, Users(Socket).Title
Case "tnl"
Line Input #1, IntBuf
Users(Socket).Tnl = Int(IntBuf)
Case "vnum"
Line Input #1, IntBuf
Users(Socket).Vnum = Int(IntBuf)
Case "recall"
Line Input #1, IntBuf
Users(Socket).Recall = Int(IntBuf)
Case "recalldefault"
Line Input #1, IntBuf
Users(Socket).RecallDefault = Int(IntBuf)

Case "object"

End Select

ElseIf Left(strBuffer, 1) = "$" Then
ComCount = ComCount + 1
ReDim Users(Socket).Comments(ComCount)
Users(Socket).Comments(ComCount) = Right$(strBuffer, Len(strBuffer) - 1)

'Else
'nul = Logit(Users(Socket).Username & "'s pfile is corrupt. (Offending Line: " & strBuffer & ")", 3, "pfile")
'frmMain.WS(Socket).SendData "Your pfile is corrupt, please create a new account and talk to an immortal under that account."
'frmMain.WS(Socket).Close
'Exit Function
End If

Loop

If Users(Socket).Password = PWord Then
ProcRec.UserState(Socket) = 4
Users(Socket).Username = Users(Socket).strBuf1
nul = Send("Welcome back " & Users(Socket).Username & "!", Socket)
nul = Logit("Comm: " & Users(Socket).Username & " has entered the mud.", 3, "comm")
Else
nul = Logit("Comm: " & Users(Socket).Username & " has entered an incorrect password, disconnected.", 3, "comm")
frmMain.WS(Socket).SendData "Incorrect Password."
frmMain.WS(Socket).Close
End If

Close #1

Exit Function
NoFile:
Close #1
frmMain.WS(Socket).SendData "No such user or corrupt pfile."
nul = Logit("***BUG ALERT***: " & Users(Socket).strBuf1 & " has a corrupt or non-existant pfile.", 3, "bugs")

End Function

Function IsImmort(Socket As Integer, Optional Username As String = "") As Boolean

If Not Len(Username) > 0 Then
    If Users(Socket).Level > 800 Then
        IsImmort = True
    Else
        IsImmort = False
    End If
Else
    For I = 0 To MaxUsers
        If Users(I).Username = Username Then
            If Users(I).Level > 800 Then
                IsImmort = True
            Else
                IsImmort = False
            End If
            Exit Function
        End If
    Next
End If

End Function


Function GetHeShe(Socket As Integer) As String
If Users(Socket).MaleFemale = 0 Then
GetHeShe = "he"
Else
GetHeShe = "she"
End If
End Function

