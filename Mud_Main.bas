Attribute VB_Name = "Mud_Main"
Type Mud
    Name As String
    WizLocked As Boolean
    Version As String
End Type
Global Mud As Mud

Function Init()
Mud.Name = frmSettings.txtName.Text
Mud.WizLocked = False
Mud.Version = "BETA 1.0 B"
End Function
