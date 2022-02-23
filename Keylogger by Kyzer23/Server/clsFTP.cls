VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Copyright Graham Daly December 2001
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

' ---------------------------------------------------------
' -------------------   API Functions  --------------------
' ---------------------------------------------------------
   
'File time information
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'File information
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Wininet.dll API's
Private Declare Function FtpFindFirstFile& Lib "wininet.dll" _
    Alias "FtpFindFirstFileA" _
    (ByVal hOutboundSession&, _
    ByVal lpszSearchFile$, _
    lpFindFileData As WIN32_FIND_DATA, _
    ByVal dwFlags&, _
    ByVal dwContent&)

Private Declare Function InternetFindNextFile& Lib "wininet.dll" _
    Alias "InternetFindNextFileA" _
    (ByVal hFind&, _
    lpvFindData As WIN32_FIND_DATA)
    
Private Declare Function FtpGetFile Lib "wininet.dll" _
    Alias "FtpGetFileA" _
    (ByVal hOutboundSession&, _
    ByVal lpszRemoteFile$, _
    ByVal lpszNewFile$, _
    ByVal fFailIfExists As Boolean, _
    ByVal dwFlagsAndAttributes&, _
    ByVal dwFlags&, _
    ByVal dwContext&) As Boolean

Private Declare Function FtpPutFile Lib "wininet.dll" _
    Alias "FtpPutFileA" _
    (ByVal hOutboundSession&, _
    ByVal lpszLocalFile$, _
    ByVal lpszRemoteFile$, _
    ByVal dwFlags&, _
    ByVal dwContext&) As Boolean

Private Declare Function FtpRenameFile Lib "wininet.dll" _
    Alias "FtpRenameFileA" _
    (ByVal hOutboundSession&, _
    ByVal sExistingName$, _
    ByVal sNewName$) As Boolean

Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" _
    (ByVal hOutboundSession&, _
    ByVal lpszFileName$) As Boolean

Private Declare Function FtpGetCurrentDirectory& Lib "wininet.dll" _
    Alias "FtpGetCurrentDirectoryA" _
    (ByVal hConnect&, _
    ByVal lpszCurrentDirectory$, _
    lpdwCurrentDirectory&)

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpSetCurrentDirectoryA" _
    (ByVal hOutboundSession&, _
    ByVal lpszDirectory$) As Boolean

Private Declare Function InternetConnect& Lib "wininet.dll" _
    Alias "InternetConnectA" _
    (ByVal hInternetSession&, _
    ByVal sServerName$, _
    ByVal nServerPort%, _
    ByVal sUserName$, _
    ByVal sPassword$, _
    ByVal lService&, _
    ByVal lFlags&, _
    ByVal lContext&)

Private Declare Function InternetOpen& Lib "wininet.dll" _
    Alias "InternetOpenA" _
    (ByVal sAgent$, _
    ByVal lAccessType&, _
    ByVal sProxyName$, _
    ByVal sProxyBypass$, _
    ByVal lFlags&)

Private Declare Function InternetCloseHandle& Lib "wininet.dll" _
    (ByVal hInet&)

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" _
    Alias "InternetGetLastResponseInfoA" _
    (lpdwError&, _
    ByVal lpszBuffer$, _
    lpdwBufferLength&) As Boolean

' ---------------------------------------------------------
' ---------------- Module-level Constants -----------------
' ---------------------------------------------------------

Private Const FAILURE& = 0
Private Const SUCCESS& = 1

'Error messages
Private Const ERR_CHANGE_DIR$ = "Cannot change directory to <%s>. It either doesn't exist, or is protected"
Private Const ERR_CONNECT_ERROR$ = "Cannot connect to FTP server <%s> using User and Password Parameters"
Private Const ERR_ALREADY_CONNECTED$ = "Cannot change property while connected to FTP server"
Private Const ERR_NO_CONNECTION$ = "Cannot connect to FTP server"
Private Const ERR_DOWNLOAD$ = "Cannot get file <%s> from FTP server"
Private Const ERR_RENAME$ = "Cannot rename file <%s>"
Private Const ERR_DELETE$ = "Cannot delete file <%s> from FTP server"
Private Const ERR_CRITICAL$ = "Cannot get connection to WinInet.dll!"

'Type of service to access
Private Const INTERNET_SERVICE_FTP = 1

'Flags
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_MULTIPART = &H200000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000

'File Transfer modes
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

'Other
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INET_SESSION_NAME$ = "ICB FTP Sesh"
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Private Const NO_CONNECTION& = FAILURE&
Private Const FWDSLASH$ = "/"
Private Const BACKSLASH$ = "\"
Private c_lConnection   As Long
' ---------------------------------------------------------
' ---------------- Module-level Variables -----------------
' ---------------------------------------------------------

'Our INET handle
Private m_hInet&

'Our FTP connection Handle
Private m_hSession&

'Standard FTP properties
Private m_HostAddr$
Private m_HostPort&
Private m_User$
Private m_Password$
Private m_Dir$

' ---------------------------------------------------------
' ------------------ User-defined Types -------------------
' ---------------------------------------------------------

Public Enum FtpError
    ERR_CANNOT_CONNECT = vbObjectError + 2001
    ERR_NO_DIR_CHANGE = vbObjectError + 2002
    ERR_CANNOT_RENAME = vbObjectError + 2003
    ERR_CANNOT_DELETE = vbObjectError + 2004
    ERR_NOT_CONNECTED_TO_SITE = vbObjectError + 2005
    ERR_CANNOT_GET_FILE = vbObjectError + 2006
    ERR_INVALID_PROPERTY = vbObjectError + 2007
    ERR_FATAL = vbObjectError + 2008
End Enum

'File Transfer types
Public Enum FileTransferType
    fttUnknown = FTP_TRANSFER_TYPE_UNKNOWN
    fttAscii = FTP_TRANSFER_TYPE_ASCII
    fttBinary = FTP_TRANSFER_TYPE_BINARY
End Enum
'________________________________________________________________________

Private Sub Class_Initialize()
    
    'Initialise variables
    m_hSession& = NO_CONNECTION&
    m_hInet& = NO_CONNECTION&
    m_HostAddr$ = vbNullString
    m_HostPort& = FAILURE&
    m_User$ = vbNullString
    m_Password$ = vbNullString
    m_Dir$ = vbNullString

End Sub
'________________________________________________________________________

Private Sub Class_Terminate()

    'Kill off connection
    If m_hSession& Then InternetCloseHandle m_hSession&

    'Kill off API Handle
    If m_hInet& Then InternetCloseHandle m_hInet&

    m_hSession& = NO_CONNECTION&
    m_hInet& = NO_CONNECTION&
    
End Sub
'________________________________________________________________________

' ---------------------------------------------------------
' ------------------- Class Properties --------------------
' ---------------------------------------------------------

Public Property Let Host(ByVal HostAddr$)
Attribute Host.VB_Description = "Get/Set the Host Name of the FTP server"
'Set the host address - only if un-connected

    If m_hSession& = NO_CONNECTION& Then
        m_HostAddr$ = HostAddr$
    Else
        Err.Raise ERR_INVALID_PROPERTY, "clsFTP:Host [Let]", ERR_ALREADY_CONNECTED$
    End If
    
End Property
'________________________________________________________________________

Public Property Get Host$()
'Get host address

    Host = m_HostAddr$
    
End Property
'________________________________________________________________________

Public Property Let Port(ByVal HostPort&)
'Set the host port - only if un-connected

    If m_hSession& = NO_CONNECTION& Then
        m_HostPort& = HostPort&
    Else
        Err.Raise ERR_INVALID_PROPERTY, "clsFTP:Port [Let]", ERR_ALREADY_CONNECTED$
    End If
    
End Property
'________________________________________________________________________

Public Property Get Port&()
'Get host port

    Port& = m_HostPort&
    
End Property
'________________________________________________________________________

Public Property Let User(ByVal UserName$)
Attribute User.VB_Description = "Get/Set the login user id for the FTP server"
'Set the user - only if un-connected

    If m_hSession& = NO_CONNECTION& Then
        m_User$ = UserName$
    Else
        Err.Raise ERR_INVALID_PROPERTY, "clsFTP:User [Let]", ERR_ALREADY_CONNECTED$
    End If
    
End Property
'________________________________________________________________________

Public Property Get User$()
'Get user

    User$ = m_User$
    
End Property
'________________________________________________________________________

Public Property Let Password(ByVal Pwd$)
Attribute Password.VB_Description = "Get/Set the password for logging into the FTP server"
'Set the password - only if un-connected

    If m_hSession& = NO_CONNECTION& Then
        m_Password$ = Pwd$
    Else
        Err.Raise ERR_INVALID_PROPERTY, "clsFTP:Password [Let]", ERR_ALREADY_CONNECTED$
    End If
      
End Property
'________________________________________________________________________

Public Property Get Password$()
'Get the password

    Password = m_Password$
    
End Property
'________________________________________________________________________

Public Property Get Directory$()
'Get directory
Dim TempDir$

    If (GetDirName&(TempDir$) <> SUCCESS&) Then
        TempDir$ = "<Unknown>"
    End If
    
    Directory$ = TempDir$
    
End Property
'________________________________________________________________________

Public Property Let Directory(ByVal Folder$)
'Set the directory - only if connected

    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:Directory [Let]", ERR_NO_CONNECTION$
    Else
        ChangeDir Folder$
    End If

End Property
'________________________________________________________________________

Public Property Get IsConnected() As Boolean
Attribute IsConnected.VB_Description = "Are we connected to an FTP server ? (True/False)"
'Are we connected? Read-only
Dim Temp$

    IsConnected = (GetDirName&(Temp$) = SUCCESS&)
    
End Property
'________________________________________________________________________

' ---------------------------------------------------------
' --------------- Exposed Class Methods -------------------
' ---------------------------------------------------------

Public Function Connect&( _
    Optional ByVal Host$ = vbNullString, _
    Optional ByVal Port& = 0, _
    Optional ByVal User$ = vbNullString, _
    Optional ByVal Password$ = vbNullString)
Attribute Connect.VB_Description = "Connect to the FTP Server using the passed parameters (if any)"
    
'Attempt to connect to FTP server
On Local Error GoTo Handler
Dim ErrMsg$

    Connect& = FAILURE&
    
    If m_hInet& = FAILURE& Then
        'Create internet session handle
        m_hInet& = InternetOpen(INET_SESSION_NAME$, _
            INTERNET_OPEN_TYPE_DIRECT, _
            vbNullString, _
            vbNullString, 0)
        If m_hInet& = FAILURE& Then
            Err.Raise ERR_FATAL, "clsFTP:Class [Initialize]", ERR_CRITICAL$
        End If
    End If
    
    'If we already have an FTP session open then raise error
    If m_hSession& Then
        Err.Raise ERR_INVALID_PROPERTY, "clsFTP:Connect", "You are already connected to FTP Server " & m_HostAddr$
    End If

    'Overwrite any existing properties if they were supplied in the
    'arguments to this method
    If Host$ <> vbNullString Then m_HostAddr$ = Host$
    If Port& <> 0 Then m_HostPort& = Port&
    If User$ <> vbNullString Then m_User$ = User$
    If Password$ <> vbNullString Then m_Password$ = Password$
   
    m_hSession& = InternetConnect(m_hInet&, m_HostAddr$, m_HostPort&, _
        m_User$, m_Password$, INTERNET_SERVICE_FTP, 0, 0)

    'Check for connection errors
    If m_hSession& = NO_CONNECTION& Then
        ErrMsg$ = Replace(ERR_CONNECT_ERROR$, "%s", m_HostAddr$)
        ErrMsg$ = ErrMsg$ & vbCrLf & GetINETErrorMsg$(Err.LastDllError)
        Err.Raise ERR_CANNOT_CONNECT, "clsFTP:Connect", ErrMsg$
    End If
    
    Connect& = SUCCESS&

ExitProc:
    Exit Function
    
Handler:
    Connect& = Err.Number
    Resume ExitProc
    
End Function
'________________________________________________________________________

Public Function Disconnect&()
Attribute Disconnect.VB_Description = "Disconnect from the remote FTP server"
'Attempt to disconnect
    
On Local Error GoTo Handler

    Disconnect& = FAILURE&
    
    'Kill off API Handles
    If m_hInet& Then InternetCloseHandle m_hInet&
    If m_hSession& Then InternetCloseHandle m_hSession&
    
    m_hSession& = NO_CONNECTION&
    m_hInet& = NO_CONNECTION&
    
    m_HostAddr$ = vbNullString
    m_User$ = vbNullString
    m_Password$ = vbNullString
    m_Dir$ = vbNullString
            
    Disconnect& = SUCCESS&
    
ExitProc:
    Exit Function
    
Handler:
    Disconnect& = Err.Number
    Resume ExitProc
      
End Function
'________________________________________________________________________

Public Function GetFile&(ByVal HostFile$, _
    ByVal ToLocalFile$, _
    Optional tt As FileTransferType = fttUnknown)
'Get the specified file and move to the desired location using
'[optional] specified file transfer type

On Local Error GoTo Handler
Dim ReturnVal&
Dim RemoteFile$
Dim RemoteDir$
Dim LocalFile$
Dim Pos&
Dim ErrMsg$

    GetFile& = FAILURE&
    
    'If not connected, raise an error
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:GetFile", ERR_NO_CONNECTION$
    End If
    
    'Get the file
    ReturnVal& = FtpGetFile(m_hSession&, HostFile$, _
        ToLocalFile$, False, INTERNET_FLAG_RELOAD, tt, 0)
    
    If ReturnVal& = FAILURE& Then
        ErrMsg$ = Replace(ERR_DOWNLOAD$, "%s", HostFile$)
        Err.Raise ERR_CANNOT_GET_FILE, "clsFTP:GetFile", ErrMsg$
    End If
    
    GetFile& = SUCCESS&

ExitProc:
    Exit Function
    
Handler:
    GetFile& = Err.Number
    Resume ExitProc

End Function
'________________________________________________________________________

Public Function PutFile&(ByVal LocalFile$, _
    ByVal ToHostFile$, _
    Optional tt As FileTransferType = fttUnknown)
    
On Local Error GoTo Handler
Dim ReturnVal&
Dim Pos&
Dim ErrMsg$
    
    PutFile& = FAILURE&
    
    'If not connected, raise an error
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:PutFile", ERR_NO_CONNECTION$
    End If
    
    ReturnVal& = FtpPutFile(m_hSession&, LocalFile$, _
        ToHostFile$, tt, 0)
        
    If ReturnVal& = FAILURE& Then
        ErrMsg$ = Replace(ERR_DOWNLOAD$, "%s", ToHostFile$)
        ErrMsg$ = ErrMsg$ & vbCrLf & GetINETErrorMsg(Err.LastDllError)
        Err.Raise ERR_CANNOT_RENAME, "clsFTP:PutFile", ErrMsg$
    End If
    
    PutFile& = SUCCESS&

ExitProc:
    Exit Function
    
Handler:
    PutFile& = Err.Number
    Resume ExitProc
    
End Function
'________________________________________________________________________

Public Function GetDirListing&(FileNames$(), FileSizes&(), Optional ByVal SubDir$ = vbNullString)
On Local Error GoTo Handler

Dim WFD As WIN32_FIND_DATA
Dim Filter$
Dim hFind&, hFindConnect&
Dim FileSize&
Dim TempFileName$, TempFileSize&
Dim FullDir$
Dim i%

    GetDirListing& = FAILURE&
    
    Screen.MousePointer = vbHourglass
    
    'Obtain the current FTP path
    Filter$ = "*.*"
    
    FullDir$ = m_Dir$ & SubDir$
    
    AddRemFwdSlash FullDir$, 1
    
    'If not connected, raise an error
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:PutFile", ERR_NO_CONNECTION$
    End If
        
    'Connection handles used by the FtpFindFirstFile
    'API go out of scope once the files are
    'listed, therefore it can not be reused.
    'This restriction is handled by obtaining
    'a fresh connection handle each time a call
    'to FtpFindFirstFile is required, and releasing
    'it once finished.
    hFindConnect& = GetInternetConnectHandle()

    hFind& = FtpFindFirstFile(m_hSession&, _
        FullDir$ & Filter$, WFD, _
        INTERNET_FLAG_RELOAD Or _
        INTERNET_FLAG_NO_CACHE_WRITE, 0&)

    If hFind& Then
        i% = 0
        Do
            ReDim Preserve FileNames$(i%)
            ReDim Preserve FileSizes&(i%)
            TempFileName$ = ClipNull(WFD.cFileName)
            If Len(TempFileName$) Then
                If WFD.dwFileAttributes And vbDirectory Then
                    TempFileName$ = TempFileName$ & FWDSLASH$
                    TempFileSize& = 0
                Else
                    TempFileSize& = WFD.nFileSizeLow
                End If
                FileNames$(i%) = TempFileName$
                FileSizes&(i%) = TempFileSize&
            End If
            'Continue while valid
            i% = i% + 1
        Loop While InternetFindNextFile(hFind&, WFD)
    
    End If 'If hFind&
    
    InternetCloseHandle hFindConnect&
    InternetCloseHandle hFind&
    
    Screen.MousePointer = vbDefault
    GetDirListing& = SUCCESS&
    
ExitProc:
    Exit Function
    
Handler:
    GetDirListing& = Err.Number
    Resume ExitProc

End Function
'________________________________________________________________________
Public Function CreateDirectory(sDir As String) As Boolean
    CreateDirectory = Not (FtpCreateDirectory(c_lConnection, sDir) = 0)
End Function
Public Function SetCurrentDirectory(sDir As String) As Boolean
    SetCurrentDirectory = Not (FtpSetCurrentDirectory(c_lConnection, sDir) = 0)
End Function
Public Function RenameFile&(ByVal FileNameOld$, ByVal FileNameNew$)
Attribute RenameFile.VB_Description = "Rename a file on the FTP server"

On Local Error GoTo Handler

'The FTP rename command can be thought of as being more than just a simple host
'file renaming facility. When invoking a rename command on a file in a sub-folder from
'the current FTP folder, the file is essentially 'moved' to its new location
'in one simple step. Actually, the file is NOT physically copied to a new location
'and the old one deleted. Because the file on the FTP server is being renamed to
'a target directory in the same file system, FTP is clever enough just
'to change the file's directory pointer so the file never gets recopied. This
'is critical in preserving file integity, and is of significant importance to
'us in our use of buffer directories.


Dim ReturnVal&
Dim ErrMsg$

    RenameFile& = FAILURE&
    
    'If not connected, raise an error
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:RenameFile", ERR_NO_CONNECTION$
    End If
   
    ReturnVal& = FtpRenameFile(m_hSession&, FileNameOld$, FileNameNew$)

    'Raise an error if we couldn't rename the file (most likely that
    'a file with the new name already exists
    If ReturnVal& = FAILURE& Then
        ErrMsg$ = Replace(ERR_RENAME$, "%s", FileNameOld$)
        ErrMsg$ = ErrMsg$ & vbCrLf & GetINETErrorMsg(Err.LastDllError)
        Err.Raise ERR_CANNOT_RENAME, "clsFTP:RenameFile", ErrMsg$
    End If
    
    RenameFile& = SUCCESS&
    
ExitProc:
    Exit Function
    
Handler:
    RenameFile& = Err.Number
    Resume ExitProc
    
End Function
'________________________________________________________________________

Public Function DeleteFile&(ByVal FileToDelete$)
Attribute DeleteFile.VB_Description = "Delete a file from the FTP server"
On Local Error GoTo Handler

Dim ReturnVal&
Dim ErrMsg$
    
    DeleteFile& = FAILURE&
    
    'Check for a connection
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:DeleteFile", ERR_NO_CONNECTION$
    End If
    
    ReturnVal& = FtpDeleteFile(m_hSession&, FileToDelete$)

    'Raise an error if the file couldn't be deleted
    If ReturnVal& = FAILURE& Then
        ErrMsg$ = Replace(ERR_DELETE$, "%s", FileToDelete$)
        Err.Raise ERR_CANNOT_DELETE, "clsFTP:DeleteFile", ErrMsg$
    End If
    
    DeleteFile& = SUCCESS&

ExitProc:
    Exit Function
    
Handler:
    DeleteFile& = Err.Number
    Resume ExitProc

End Function
'________________________________________________________________________

' ---------------------------------------------------------
' ----------------- Private Functions ---------------------
' ---------------------------------------------------------

Private Function GetDirName&(FTPDir$)
On Local Error GoTo Handler

Dim BufferLen&
Dim BufferStr$
Dim ReturnVal&

    GetDirName& = FAILURE&
    FTPDir$ = vbNullString
    m_Dir$ = vbNullString
    
    'If not connected, raise an error
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:GetDirName", ERR_NO_CONNECTION$
    End If
    
    BufferStr$ = Space$(256)
    BufferLen& = Len(BufferStr$)
       
    ReturnVal& = FtpGetCurrentDirectory(m_hSession&, BufferStr$, BufferLen&)
    
    If ReturnVal& = SUCCESS& Then
        'return a properly qualified path
        BufferStr$ = ClipNull(BufferStr$)
        
        m_Dir$ = BufferStr$
        FTPDir$ = m_Dir$
        GetDirName& = SUCCESS&
    End If

ExitProc:
    Exit Function
    
Handler:
    GetDirName& = Err.Number
    Resume ExitProc

End Function
'________________________________________________________________________

Private Function ChangeDir&(ByVal HostDir$)
On Local Error GoTo Handler

Dim ReturnVal&
Dim ErrMsg$

    ChangeDir& = FAILURE&
    
    'Ensure that rightmost character is a backslash
    AddRemSlash HostDir$, 1
    
    'Replace all back-slashes with forward-slashes: Telnet standard
    HostDir$ = Replace(HostDir$, BACKSLASH$, FWDSLASH$)

    'Check for a connection
    If m_hSession& = NO_CONNECTION& Then
        Err.Raise ERR_NOT_CONNECTED_TO_SITE, "clsFTP:ChangeDir", ERR_NO_CONNECTION$
    End If
    
    ReturnVal& = FtpSetCurrentDirectory(m_hSession&, HostDir$)

    'If we can't change directory - raise an error
    If ReturnVal& = FAILURE& Then
        ErrMsg$ = ERR_CHANGE_DIR$
        ErrMsg$ = Replace(ErrMsg$, "%s", HostDir$)
        Err.Raise ERR_CHANGE_DIR$, "clsFTP:ChangeDir", ErrMsg$
    End If
    
    ChangeDir& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    ChangeDir& = Err.Number
    Resume ExitProc

End Function
'________________________________________________________________________

Private Function GetINETErrorMsg$(ByVal ErrNum&)
Dim LenError&, LenBuffer&
Dim Buffer$

    'Get extra info from the WinInet.DLL
    If ErrNum& = ERROR_INTERNET_EXTENDED_ERROR Then
        'Get message size and number
        InternetGetLastResponseInfo LenError&, vbNullString, LenBuffer&
        Buffer$ = String$(LenBuffer& + 1, vbNullChar)
        'Get message
        InternetGetLastResponseInfo LenError&, Buffer$, LenBuffer&
        GetINETErrorMsg = vbCrLf & Buffer$
    End If
    
End Function
'________________________________________________________________________

Private Function GetInternetConnectHandle&()
Dim sServerName As String
Dim H&
    
'Obtains a new handle expressly for use with the
'FtpFindFirstFile API.
'
'Care must be taken to close only the handle
'returned by this function once the listing
'of the directory has been obtained.
    
    If m_hInet& Then
        H& = InternetConnect(m_hInet&, _
            m_HostAddr$, _
            m_HostPort&, _
            m_User$, _
            m_Password$, _
            INTERNET_SERVICE_FTP, _
            INTERNET_FLAG_EXISTING_CONNECT Or INTERNET_FLAG_PASSIVE, _
            &H0)
    End If

    GetInternetConnectHandle& = H&
      
End Function
'________________________________________________________________________

Private Function AddRemFwdSlash&(PathName$, ByVal IsSlash As Byte)
On Local Error GoTo Handler
    
    AddRemFwdSlash& = FAILURE&
    
    If IsSlash Then       'We want a "\" at end
        If Right(PathName$, 1) <> FWDSLASH$ Then PathName$ = PathName$ & FWDSLASH$
        
    Else                        'We don't want a "\" at end
        If Right(PathName$, 1) = FWDSLASH$ Then
            PathName$ = Mid(PathName$, 1, Len(PathName$) - 1)
        End If
    End If
    
    AddRemFwdSlash& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    AddRemFwdSlash& = Err.Number
    'MsgBox Err.Number & ": " & Err.Description, _
            vbExclamation, "AddRemFwdSlash Error"
    Resume ExitProc
    
End Function
'________________________________________________________________________

Public Function AddRemSlash&(PathName$, ByVal IsSlash As Byte)
On Local Error GoTo Handler
    
    AddRemSlash& = FAILURE&
    
    If IsSlash Then       'We want a "\" at end
        If Right(PathName$, 1) <> BACKSLASH$ Then PathName$ = PathName$ & BACKSLASH$
    Else                        'We don't want a "\" at end
        If Right(PathName$, 1) = BACKSLASH$ Then
            PathName$ = Mid(PathName$, 1, Len(PathName$) - 1)
        End If
    End If
    
    AddRemSlash& = SUCCESS&
    
ExitProc:
    Exit Function

Handler:
    AddRemSlash& = Err.Number
    'MsgBox Err.Number & ": " & Err.Description, _
            vbExclamation, "AddRemSlash Error"
    Resume ExitProc
    
End Function
'________________________________________________________________________

Private Function ClipNull$(ByVal str$)
Dim Pos%

    Pos% = InStr(1, str$, vbNullChar)
    If Pos% > 0 Then
        ClipNull$ = Left(str$, Pos% - 1)
    End If
    
End Function
