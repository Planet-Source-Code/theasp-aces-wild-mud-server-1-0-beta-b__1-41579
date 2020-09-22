Attribute VB_Name = "GUI"

Function MB(Message As String, Optional DisplayFor As Integer = 0)
frmMB.lblMessage.Caption = Message
If DisplayFor > 0 Then
frmMB.tmrAuto.Interval = DisplayFor & "000"
frmMB.tmrAuto.Enabled = True
End If
frmMB.Show
End Function

Function PCase(szIn As String) As String
PCase = UCase(Left(szIn, 1)) & Right(szIn, Len(szIn) - 1)
End Function
