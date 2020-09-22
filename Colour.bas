Attribute VB_Name = "Colour"
'ANSI Name  Code Sequence   Action
'
'*******************************************************************************
'
'     ED    ESC[ Pn J       Erases all or part of a display.
'                           Pn=0: erases from active position to end of display.
'                           Pn=1: erases from the beginning of display to
'                                 active position.
'                           Pn=2: erases entire display.
'
'     EL    ESC[ Pn K       Erases all or part of a line.
'                           Pn=0: erases from active position to end of line.
'                           Pn=1: erases from beginning of line to active
'                                 position.
'                           Pn=2: erases entire line.
'
'     ECH   ESC[ Pn X       Erases Pn characters.
'
'     CBT   ESC[ Pn Z       Moves active position back Pn tab stops.
'
'     SU    ESC[ Pn S       Scroll screen up Pn lines, introducing new blank
'                           lines at bottom.
'
'     SD    ESC[ Pn T       Scroll screen down Pn lines, introducing new blank
'                           lines at top.
'
'     CUP   ESC[ P1;P2 H    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     HVP   ESC[ P1;P2 f    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     CUU   ESC[ Pn A       Moves cursor up Pn number of lines.
'
'     CUD   ESC[ Pn B       Moves cursor down Pn number of lines.
'
'     CUF   ESC[ Pn C       Moves cursor Pn spaces to the right.
'
'     CUB   ESC[ Pn D       Moves cursor Pn spaces backward.
'
'     HPA   ESC[ Pn '       Moves cursor to column given by Pn.
'
'     HPR   ESC[ Pn a       Moves cursor Pn characters to the right.
'
'     VPA   ESC[ Pn d       Moves cursor to line given by Pn.
'
'     VPR   ESC[ Pn e       Moves cursor down Pn number of lines.
'
'     IL    ESC[ Pn L       Inserts Pn new, blank lines.
'
'     ICH   ESC[ Pn @       Inserts Pn blank places for Pn characters.
'
'     DL    ESC[ Pn M       Deletes Pn lines.
'
'     DCH   ESC[ Pn P       Deletes Pn number of characters.
'
'     CPL   ESC[ Pn F       Moves cursor to beginning of line, Pn lines up.
'
'     CNL   ESC[ Pn E       Moves cursor to beginning of line, Pn lines down.
'
'     SGR   ESC[ Pn m       Changes display mode.
'                           Pn=0: Resets bold, blink, blank, underscore, and
'                           reverse.
'                           Pn=1: Sets bold (light_color).
'                           Pn=4: Sets underscore.
'                           Pn=5: Sets blink.
'                           Pn=7: Sets reverse video.
'                           Pn=8: Sets blank (no display).
'                           Pn=10: Select primary font.
'                           Pn=11: Select first alternate font.
'                           Pn=12: Select second alternate font.
'
'           ESC[ 2h         Lock keyboard. Ignores keyboard input until
'                           unlocked.
'
'           ESC[ 2i         Send screen to host.
'
'           ESC[ 2l         Unlock keyboard.
'
'           ESC[ 3 C m      Selects foreground colour C.
'
'           ESC[ 4 C m      Selects background colour C.
'
'                           C=0  Black
'                           C=1  Red
'                           C=2  Green
'                           C=3  Yellow
'                           C=4  Dark Blue
'                           C=5  Magenta
'                           C=6  Light Blue
'                           C=7  White
'
'****************************************************************************

Function SAnsi(Color As String, Optional Bright As Boolean = False) As String

Select Case LCase(Color)

Case "black"
If Bright = True Then
SSAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "0" & "m"
Else
SSAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "0" & "m"
End If

Case "red"
If Bright = True Then
SSAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "1" & "m"
Else
SSAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "1" & "m"
End If

Case "green"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "2" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "2" & "m"
End If

Case "yellow"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "3" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "3" & "m"
End If

Case "blue"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "4" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "4" & "m"
End If

Case "purple"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "5" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "5" & "m"
End If

Case "cyan"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "6" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "6" & "m"
End If

Case "white"
If Bright = True Then
SAnsi = Chr(27) & "[1m" & Chr(27) & "[3" & "7" & "m"
Else
SAnsi = Chr(27) & "[0m" & Chr(27) & "[3" & "7" & "m"
End If

Case Else
nul = Logit("Unknown Colour: " & Color, 3, "bugs")

End Select

    'Select Case Color
    'Case Is <= 7
    '    Ansi = Chr(27) & "[0m" & Chr(27) & "[3" & Color & "m"
    'Case Is >= 8
    '    Ansi = Chr(27) & "[1m" & Chr(27) & "[3" & Color - 8 & "m"
    'End Select
    
End Function


Function Ansi(Color As Integer) As String
    Select Case Color
    Case Is <= 7
        Ansi = Chr(27) & "[0m" & Chr(27) & "[3" & Color & "m"
    Case Is >= 8
        Ansi = Chr(27) & "[1m" & Chr(27) & "[3" & Color - 8 & "m"
    End Select
End Function

