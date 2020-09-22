Attribute VB_Name = "World"
Public WorldTitle As String
Public WorldAuthor As String
Public WorldFileName As String

Const MaxRooms As Integer = 200
Const MaxObjs As Integer = 50
Const MaxMobs As Integer = 50
Public RoomName(MaxRooms) As String
Public RoomDesc(MaxRooms) As String
Public RoomExits(MaxRooms, 3) As Integer

Public ObjShort(MaxObjs) As String
Public ObjLong(MaxObjs) As String
Public ObjType(MaxObjs) As String

Public MobShort(MaxMobs) As String
Public MobLong(MaxMobs) As String
Public MobFunc(MaxMobs) As String

Function LoadWorld(WorldFile As String)
nul = Logit("Loading World File: " & WorldFile & "...", 2, "worlds")
Dim RCount As Integer
RCount = -1
Dim MCount As Integer
MCount = -1
Dim OCount As Integer
OCount = -1

WorldFileName = WorldFile

Open App.Path & "\worlds\" & WorldFile & ".world" For Input As #1

Do While Not EOF(1)

Line Input #1, strBuffer

Select Case strBuffer

Case "#header"
Line Input #1, WorldTitle
Line Input #1, WorldAuthor

Case "#room"
Line Input #1, VN
Vnum = Int(VN)
Line Input #1, RoomName(Vnum)
Line Input #1, BufVar
RoomDesc(Vnum) = Replace(BufVar, "~^~", Chr$(13))
Line Input #1, Ex
RoomExits(Vnum, 0) = Int(Ex)
Line Input #1, Ex
RoomExits(Vnum, 1) = Int(Ex)
Line Input #1, Ex
RoomExits(Vnum, 2) = Int(Ex)
Line Input #1, Ex
RoomExits(Vnum, 3) = Int(Ex)

Case "#DPC"
Line Input #1, strBuffer
Debug.Print strBuffer

End Select

Loop

Close #1
nul = Logit("Area loaded without a problem.", 2, "worlds")
End Function
