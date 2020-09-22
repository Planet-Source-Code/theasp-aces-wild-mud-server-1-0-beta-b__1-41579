VERSION 5.00
Begin VB.Form frmMB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For your information..."
   ClientHeight    =   2835
   ClientLeft      =   405
   ClientTop       =   6990
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Left            =   240
      Top             =   2280
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub tmrAuto_Timer()
Me.Hide
tmrAuto.Enabled = False
End Sub
