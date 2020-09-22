VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logs"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Delete Log"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Append To Screen"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "system.log"
      Top             =   0
      Width           =   3255
   End
   Begin VB.ListBox lstLog 
      Height          =   4740
      ItemData        =   "frmLog.frx":0000
      Left            =   0
      List            =   "frmLog.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
On Error GoTo NoFile

If txtFile.Text = "master.log" Then
MB ("File not found.")
Exit Sub
End If

nul = Logit("Deleting " & txtFile.Text & "...", 1, "master")
Kill (App.Path & "\logs\" & txtFile.Text)
nul = Logit("Log deleted.", 1, "master")

Exit Sub
NoFile:
MB ("File not found.")
End Sub

Private Sub cmdOpen_Click()
On Error GoTo NoFile

If txtFile.Text = "master.log" Then
MB ("File not found.")
Exit Sub
End If

Open App.Path & "\logs\" & txtFile.Text For Input As #1
lstLog.AddItem (" ")
lstLog.AddItem ("----------" & txtFile.Text & "----------")
lstLog.AddItem (" ")
Do While Not EOF(1)
Line Input #1, strBuffer
lstLog.AddItem (strBuffer)
Loop
Close #1
lstLog.AddItem (" ")
lstLog.AddItem ("----------End " & txtFile.Text & "----------")
lstLog.AddItem (" ")

Exit Sub
NoFile:
MB ("File not found.")
End Sub

