VERSION 5.00
Begin VB.Form frmConn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Manager"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton cmdTran 
      Caption         =   "Transfer"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton cmdDis 
      Caption         =   "Disconnect User"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
   Begin VB.ListBox lstConn 
      Height          =   3765
      ItemData        =   "frmConn.frx":0000
      Left            =   0
      List            =   "frmConn.frx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.ComboBox cmbConn 
      Height          =   315
      ItemData        =   "frmConn.frx":0004
      Left            =   2640
      List            =   "frmConn.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Data"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   4575
   End
End
Attribute VB_Name = "frmConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_-\|/-_

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDis_Click()
For i = 0 To cmbConn.ListCount
If cmbConn.List(i) = cmbConn.Text Then
frmMain.WS(cmbConn.ItemData(i)).Close
End If
Next
Call Form_Load
End Sub

Private Sub cmdRefresh_Click()
Call Form_Load
End Sub

Private Sub cmdSend_Click()
On Error Resume Next
For i = 0 To cmbConn.ListCount
If cmbConn.List(i) = cmbConn.Text Then
Message = InputBox("What would you like to send to this user?")
If Not Message = "" Then
If ProcRec.UserState(cmbConn.ItemData(i)) = 4 Then
nul = Send(Ansi(11) & Message & Ansi(7) & vbCrLf, cmbConn.ItemData(i))
Else
frmMain.WS(cmbConn.ItemData(i)).SendData (Ansi(11) & Message & Ansi(7) & vbCrLf)
End If
End If
End If
Next
End Sub

Private Sub cmdTran_Click()
For i = 0 To cmbConn.ListCount
If cmbConn.List(i) = cmbConn.Text Then
NEwV = InputBox("Transfer to what vnum?")
If Not NEwV = "" Then
Users.Users(cmbConn.ItemData(i)).Vnum = Int(NEwV)
Commands.Do_Look (cmbConn.ItemData(i))
End If
End If
Next
End Sub

Private Sub Form_Load()
cmbConn.Clear
lstConn.Clear
For i = 0 To MaxUsers
    If Sockets.SocketState(i) = 1 Then
        If ProcRec.UserState(i) = 4 Then
            lstConn.AddItem (frmMain.WS(i).RemoteHostIP & " - " & Users.Users(i).Username)
            cmbConn.AddItem (frmMain.WS(i).RemoteHostIP & " - " & Users.Users(i).Username)
            cmbConn.ItemData(cmbConn.ListCount - 1) = i
        Else
            lstConn.AddItem (frmMain.WS(i).RemoteHostIP & " - N\A")
            cmbConn.AddItem (frmMain.WS(i).RemoteHostIP & " - N\A")
            cmbConn.ItemData(cmbConn.ListCount - 1) = i
        End If
    End If
Next
End Sub
