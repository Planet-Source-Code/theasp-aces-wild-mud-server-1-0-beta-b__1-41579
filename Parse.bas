Attribute VB_Name = "ParseMod"
Public Results() As String
Public Count As Integer

Function Parse(strData As String, Delim As String)
Dim Buffer As String
Count = -1

For I = 0 To Len(strData)
Digit = Left$(strData, I)
Digit = Right$(Digit, 1)

If Digit = Delim Then
Count = Count + 1
ReDim Preserve Results(Count)
Results(Count) = LCase(Buffer)
Buffer = ""
Else
Buffer = Buffer & Digit
End If
Next

If Len(Buffer) > 0 Then
Count = Count + 1
ReDim Preserve Results(Count)
Results(Count) = LCase(Buffer)
Buffer = ""
End If

End Function
