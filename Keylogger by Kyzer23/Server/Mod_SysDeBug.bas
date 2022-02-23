Attribute VB_Name = "Mod_SysDeBug"

Option Explicit

Private Declare Function AdjustTokenPrivileges _
                Lib "advapi32.dll" (ByVal TokenHandle As Long, _
                                    ByVal DisableAllPriv As Long, _
                                    ByRef NewState As TOKEN_PRIVILEGES, _
                                    ByVal BufferLength As Long, _
                                    ByRef PreviousState As TOKEN_PRIVILEGES, _
                                    ByRef pReturnLength As Long) As Long
Private Declare Function GetCurrentProcess _
                Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue _
                Lib "advapi32.dll" _
                Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, _
                                               ByVal lpName As String, _
                                               lpLuid As LUID) As Long
Private Declare Function OpenProcessToken _
                Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
                                    ByVal DesiredAccess As Long, _
                                    TokenHandle As Long) As Long

Private Type MEMORY_CHUNKS
        Address As Long
        pData As Long
        Length As Long
End Type
Private Type LUID
        UsedPart As Long
        IgnoredForNowHigh32BitPart As Long
End Type '

Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
End Type
Public Const SE_CREATE_TOKEN = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT = "SeMachineAccountPrivilege"
Public Const SE_TCB = "SeTcbPrivilege"
Public Const SE_SECURITY = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP = "SeBackupPrivilege"
Public Const SE_RESTORE = "SeRestorePrivilege"
Public Const SE_SHUTDOWN = "SeShutdownPrivilege"
Public Const SE_DEBUG = "SeDebugPrivilege"
Public Const SE_AUDIT = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN = "SeRemoteShutdownPrivilege"
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20

'提权相关常数
Public Const SE_MIN_WELL_KNOWN_PRIVILEGE = 2
Public Const SE_CREATE_TOKEN_PRIVILEGE = 2
Public Const SE_ASSIGNPRIMARYTOKEN_PRIVILEGE = 3
Public Const SE_LOCK_MEMORY_PRIVILEGE = 4
Public Const SE_INCREASE_QUOTA_PRIVILEGE = 5

' end_wdm
'
' Unsolicited Input is obsolete and unused.
'

Public Const SE_UNSOLICITED_INPUT_PRIVILEGE = 6

' begin_wdm
Public Const SE_MACHINE_ACCOUNT_PRIVILEGE = 6
Public Const SE_TCB_PRIVILEGE = 7
Public Const SE_SECURITY_PRIVILEGE = 8
Public Const SE_TAKE_OWNERSHIP_PRIVILEGE = 9
Public Const SE_LOAD_DRIVER_PRIVILEGE = 10
Public Const SE_SYSTEM_PROFILE_PRIVILEGE = 11
Public Const SE_SYSTEMTIME_PRIVILEGE = 12
Public Const SE_PROF_SINGLE_PROCESS_PRIVILEGE = 13
Public Const SE_INC_BASE_PRIORITY_PRIVILEGE = 14
Public Const SE_CREATE_PAGEFILE_PRIVILEGE = 15
Public Const SE_CREATE_PERMANENT_PRIVILEGE = 16
Public Const SE_BACKUP_PRIVILEGE = 17
Public Const SE_RESTORE_PRIVILEGE = 18
Public Const SE_SHUTDOWN_PRIVILEGE = 19
Public Const SE_DEBUG_PRIVILEGE = 20
Public Const SE_AUDIT_PRIVILEGE = 21
Public Const SE_SYSTEM_ENVIRONMENT_PRIVILEGE = 22
Public Const SE_CHANGE_NOTIFY_PRIVILEGE = 23
Public Const SE_REMOTE_SHUTDOWN_PRIVILEGE = 24
Public Const SE_UNDOCK_PRIVILEGE = 25
Public Const SE_SYNC_AGENT_PRIVILEGE = 26
Public Const SE_ENABLE_DELEGATION_PRIVILEGE = 27
Public Const SE_MANAGE_VOLUME_PRIVILEGE = 28
Public Const SE_IMPERSONATE_PRIVILEGE = 29
Public Const SE_CREATE_GLOBAL_PRIVILEGE = 30
Public Const SE_MAX_WELL_KNOWN_PRIVILEGE = SE_CREATE_GLOBAL_PRIVILEGE
Public Const STATUS_NO_TOKEN = &HC000007C
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Boolean, ByVal Client As Boolean, WasEnabled As Long) As Long

Public Function EnablePrivilege1(ByVal Privilege As Long, Enable As Boolean) As Boolean
    Dim ntStatus As Long
    Dim WasEnabled As Long
    ntStatus = RtlAdjustPrivilege(Privilege, Enable, True, WasEnabled)
    If ntStatus = STATUS_NO_TOKEN Then
        ntStatus = RtlAdjustPrivilege(Privilege, Enable, False, WasEnabled)
    End If
    If ntStatus = 0 Then
        EnablePrivilege1 = True
    Else
        EnablePrivilege1 = False
    End If
End Function


Public Function EnablePrivilege(ByVal seName As String) As Boolean
        On Error Resume Next
        Dim p_lngRtn As Long
        Dim p_lngToken As Long
        Dim p_lngBufferLen As Long
        Dim p_typLUID As LUID
        Dim p_typTokenPriv As TOKEN_PRIVILEGES
        Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
        p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        If Err.LastDllError <> 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_typTokenPriv.PrivilegeCount = 1
        p_typTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        p_typTokenPriv.TheLuid = p_typLUID
        EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function





