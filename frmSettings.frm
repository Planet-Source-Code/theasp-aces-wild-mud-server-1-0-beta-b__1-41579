VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACES Wild - Settings"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVersion 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtWorld 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   4335
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label5 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "HogwartsWizard"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Mud Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "MUD Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "World:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
Dim DoReStart As Boolean

    
    If Running = True Then
    frmMain.DoStop
    DoReStart = True
    End If
    
    Mud.Name = txtName.Text
    Port = Int(txtPort.Text)
    World.WorldFileName = txtWorld.Text
    
    If txtPort.Text = "" Then
    Port = 0
    GoTo SkipRe
    End If
    
    If DoReStart = True Then
    frmMain.DoStart
    End If
    
SkipRe:
    
    SaveSettings
    
    Call Form_Load
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim DoReStart As Boolean

    If Running = True Then
    frmMain.DoStop
    DoReStart = True
    End If
    
    Mud.Name = txtName.Text
    Port = Int(txtPort.Text)
    World.WorldFileName = txtWorld.Text
    
    
    If txtPort.Text = "" Then
    Port = 0
    GoTo SkipRe
    End If
    
    If DoReStart = True Then
    frmMain.DoStart
    End If
    
SkipRe:
    
    SaveSettings

    Unload Me

End Sub

Private Sub Form_Load()
    txtName.Text = Mud.Name
    txtPort.Text = Port
    txtVersion.Text = Mud.Version
    txtWorld.Text = World.WorldFileName
    txtVersion.Locked = True
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Function SaveSettings()
Dim nul As String
If Mud.Name = "" Then
nul = MsgBox("Please enter a name for the mud.", , "Missing, Empty, or Invalid Entry")
Exit Function
End If

If World.WorldFileName = "" Then
nul = MsgBox("Please enter a world file name, or default for the default file.", , "Missing, Empty, or Invalid Entry")
Exit Function
End If

If Port = 0 Then
nul = MsgBox("You MUST enter a port number between 1 and 32000.", , "Missing, Empty, or Invalid Entry")
Exit Function
End If

If Port > 32000 Then
nul = MsgBox("You MUST enter a port number between 1 and 32000.", , "Missing, Empty, or Invalid Entry")
Exit Function
End If

If Port < 1 Then
nul = MsgBox("You MUST enter a port number between 1 and 32000.", , "Missing, Empty, or Invalid Entry")
Exit Function
End If

Open App.Path & "\settings.dat" For Output As #1
Print #1, Mud.Name
Print #1, World.WorldFileName
Print #1, Port
Close #1
End Function
