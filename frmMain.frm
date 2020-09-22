VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACES WILD Mud Server"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConn 
      Caption         =   "Connection              Manager"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Settings"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock WS2 
      Index           =   0
      Left            =   4200
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogs 
      Caption         =   "Logs"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Left            =   4320
      Top             =   2520
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ListBox lstLog 
      Height          =   2010
      ItemData        =   "frmMain.frx":0000
      Left            =   0
      List            =   "frmMain.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock WSListen 
      Left            =   4320
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   4080
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TimerDone As Boolean

Private Sub cmdConn_Click()
frmConn.Show
End Sub

Private Sub cmdLogs_Click()
frmLog.Show
End Sub

Private Sub cmdSet_Click()
frmSettings.Show
End Sub

Private Sub cmdStart_Click()
World.LoadWorld (World.WorldFileName)
Sockets.StartServer
Tables.LoadHelp
cmdStop.Enabled = True
cmdStart.Enabled = False
Running = True
End Sub

Function DoStart()
World.LoadWorld (World.WorldFileName)
Sockets.StartServer
Tables.LoadHelp
cmdStop.Enabled = True
cmdStart.Enabled = False
Running = True
End Function

Private Sub cmdStop_Click()
Sockets.StopServer
cmdStop.Enabled = False
cmdStart.Enabled = True
Running = False
End Sub

Function DoStop()
Sockets.StopServer
cmdStop.Enabled = False
cmdStart.Enabled = True
Running = False
End Function

Private Sub Form_Load()
On Error Resume Next

Open App.Path & "\settings.dat" For Input As #1
Line Input #1, Mud.Name
Line Input #1, World.WorldFileName
Line Input #1, IntBuf
Port = Int(IntBuf)
If Port = 0 Then
Port = 4000
End If
If Mud.Name = "" Then
Mud.Name = "Development Mud"
End If
If World.WorldFileName = "" Then
World.WorldFileName = "default"
End If
Close #1

Sockets.Init
Mud_Main.Init
Socials.Init
nul = Logit("Application Started", 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Running = True Then
nul = Logit("Server Shutdown By System Administrator, the mud was still running.", 2)
Else
nul = Logit("Server Shutdown By System Administrator, the mud was not running.", 2)
End If
End
End Sub

Private Sub tmrWait_Timer()
TimerDone = True
End Sub

Private Sub WS_Close(Index As Integer)
Sockets.SocketState(Index) = 0
Sockets.CloseSock (Index)
End Sub

Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Buffer As String
WS(Index).GetData Buffer, vbString
nul = Proccess(Buffer, Index)
End Sub

Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
nul = Logit("Error in WS socket " & Index & ", deactivating. It reported (Number\Description): " & Number & "\" & Description, 3)
WS(Index).Close
Sockets.SocketState(Index) = -1
CancelDisplay = True
End Sub

Private Sub WSListen_ConnectionRequest(ByVal requestID As Long)
Sockets.Connect (requestID)
End Sub

Private Sub WSListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
nul = Logit("WSListen has had an error. This is the number\description it gave: " & Number & "\" & Description, 2)
CancelDisplay = True
End Sub

Function Wait(Seconds As Integer)
tmrWait.Interval = Seconds & "000"
TimerDone = False
tmrWait.Enabled = True
Do While TimerDone = False
DoEvents
Loop
tmrWait.Enabled = False
End Function
