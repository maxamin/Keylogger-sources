VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : cFTP
' DateTime    : 23/09/2008 18:34
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' WebPage     : http://www.advancevb.com.ar
' Purpose     : FTP Class
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Comments    : Not fully complete, TODO up/download files in chynks, better error handling
'
' History     : 23/09/2008 First Cut....................................................
'---------------------------------------------------------------------------------------
Option Explicit

Private Const AGENT                         As String = "COBEIN_FTP_CLASS"

Private Const INTERNET_FLAG_PASSIVE         As Long = &H8000000

Private Const INTERNET_SERVICE_FTP          As Long = 1

Private Const INTERNET_OPEN_TYPE_PRECONFIG  As Long = 0
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = 1
Private Const INTERNET_OPEN_TYPE_PROXY      As Long = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Long = 4

'Private Const INTERNET_FLAG_ASYNC           As Long = &H10000000

Private Const INTERNET_DEFAULT_FTP_PORT     As Long = 21

Private Const MAX_PATH                      As Long = 260

Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10

Enum eAccessType
    eAT_Preconfig = INTERNET_OPEN_TYPE_PRECONFIG
    eAT_Direct = INTERNET_OPEN_TYPE_DIRECT
    eAT_Proxy = INTERNET_OPEN_TYPE_PROXY
    eAT_PreconfigNoAutoProxy = INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY
End Enum

Enum eConnStatus
    eCN_Disconnected
    eCN_Connected
End Enum

Private Type FILETIME
    dwLowDateTime                           As Long
    dwHighDateTime                          As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes                        As Long
    ftCreationTime                          As FILETIME
    ftLastAccessTime                        As FILETIME
    ftLastWriteTime                         As FILETIME
    nFileSizeHigh                           As Long
    nFileSizeLow                            As Long
    dwReserved0                             As Long
    dwReserved1                             As Long
    cFileName                               As String * MAX_PATH
    cAlternate                              As String * 14
End Type

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private c_lSession      As Long
Private c_lConnection   As Long

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Call Disconnect
End Sub

Public Function Connect( _
       ByVal sUrl As String, _
       Optional ByVal sUsername As String = "anonymous", _
       Optional ByVal sPassword As String = vbNullString, _
       Optional ByVal lPort As Long = INTERNET_DEFAULT_FTP_PORT, _
       Optional ByVal bPassiveSemantic As Boolean = True, _
       Optional ByVal eAccessType As eAccessType = eAT_Preconfig, _
       Optional ByVal sProxyName As String = vbNullString, _
       Optional ByVal sProxyBypass As String = vbNullString) As Boolean
    
    c_lSession = InternetOpen( _
       AGENT, _
       eAccessType, _
       sProxyName, _
       sProxyBypass, 0)
        
    If c_lSession = 0 Then
        Exit Function
    End If
    
    Dim lFlag As Long
    
    If bPassiveSemantic Then
        lFlag = INTERNET_FLAG_PASSIVE
    End If
    
    c_lConnection = InternetConnect( _
       c_lSession, _
       sUrl, _
       lPort, _
       sUsername, _
       sPassword, _
       INTERNET_SERVICE_FTP, _
       lFlag, 0)
       
    If c_lConnection = 0 Then
        Call InternetCloseHandle(c_lSession): c_lSession = 0
        Exit Function
    End If
    
    Connect = True
    
End Function

Public Function Disconnect() As Boolean
    If Not c_lSession = 0 Then
        Call InternetCloseHandle(c_lSession):       c_lSession = 0
        Call InternetCloseHandle(c_lConnection):    c_lConnection = 0
        Disconnect = True
    End If
End Function

Public Function Status() As eConnStatus
    Dim sRet As String * MAX_PATH
    Status = IIf(FtpGetCurrentDirectory(c_lConnection, sRet, MAX_PATH) = 0, _
       eCN_Disconnected, eCN_Connected)
End Function

Public Function GetCurrentDirectory() As String
    Dim lRet    As Long
    Dim sBuffer As String
    
    sBuffer = Space(MAX_PATH): lRet = MAX_PATH
    If Not FtpGetCurrentDirectory(c_lConnection, sBuffer, lRet) = 0 Then
        GetCurrentDirectory = Left$(sBuffer, lRet)
    Else
        GetCurrentDirectory = "/"
    End If
End Function

Public Function DirectoryUp() As Boolean
    Dim sDir        As String
    Dim svFolders() As String
    
    sDir = GetCurrentDirectory
    If Not Len(sDir) = 1 Then
        sDir = IIf(Right$(sDir, 1) = "/", sDir, sDir & "/")
        svFolders = Split(sDir, "/")
        ReDim Preserve svFolders(UBound(svFolders) - 2)
        sDir = Join(svFolders, "/")
        DirectoryUp = SetCurrentDirectory(sDir)
    End If
End Function

Public Function CreateDirectory(sDir As String) As Boolean
    CreateDirectory = Not (FtpCreateDirectory(c_lConnection, sDir) = 0)
End Function

Public Function DeleteDirectory(sDir As String) As Boolean
    DeleteDirectory = Not (FtpRemoveDirectory(c_lConnection, sDir) = 0)
End Function

Public Function SetCurrentDirectory(sDir As String) As Boolean
    SetCurrentDirectory = Not (FtpSetCurrentDirectory(c_lConnection, sDir) = 0)
End Function

Public Function EnumDirectories(Optional ByVal sFilter As String = "*") As Collection
    Dim pData   As WIN32_FIND_DATA
    Dim hFind   As Long
    Dim lRet    As Long
    Dim cDirs   As New Collection
    
    Set EnumDirectories = cDirs
        
    pData.cFileName = Space(MAX_PATH)
    hFind = FtpFindFirstFile(c_lConnection, sFilter, pData, 0, 0)
    If Not hFind = 0 Then
        If (pData.dwFileAttributes Or FILE_ATTRIBUTE_DIRECTORY = _
           pData.dwFileAttributes) Then
            cDirs.Add Left(pData.cFileName, lstrlen(pData.cFileName))
            pData.cFileName = Space(MAX_PATH)
        End If
        Do While InternetFindNextFile(hFind, pData)
            If (pData.dwFileAttributes Or FILE_ATTRIBUTE_DIRECTORY = _
               pData.dwFileAttributes) Then
                cDirs.Add Left(pData.cFileName, lstrlen(pData.cFileName))
                pData.cFileName = Space(MAX_PATH)
            End If
        Loop
        Call InternetCloseHandle(hFind)
    End If
    
    Set EnumDirectories = cDirs
    
End Function

Public Function PutFile(ByVal sLocalFile As String, ByVal sRemoteFile As String) As Boolean
    PutFile = Not (FtpPutFile(c_lConnection, sLocalFile, sRemoteFile, 0, 0) = 0)
End Function

'Public Function GetFile(ByVal sLocalFile As String, ByVal sRemoteFile As String) As Boolean
'    PutFile = Not (FTPGetFile(c_lConnection, sRemoteFile, sLocalFile, False, 0, 0, 0) = 0)
'End Function

Public Function DeleteFile(ByVal sFile As String) As Boolean
    DeleteFile = Not (FtpDeleteFile(c_lConnection, sFile) = 0)
End Function

'Public Function RenameFile(ByVal sFile As String) As Boolean
'    RenameFile = Not (FtpRenameFile(c_lConnection, sFile) = 0)
'End Function

Public Function EnumFiles(Optional ByVal sFilter As String = "*") As Collection
    Dim pData   As WIN32_FIND_DATA
    Dim hFind   As Long
    Dim lRet    As Long
    Dim cFiles  As New Collection
    
    Set EnumFiles = cFiles
        
    pData.cFileName = Space(MAX_PATH)
    hFind = FtpFindFirstFile(c_lConnection, sFilter, pData, 0, 0)
    If Not hFind = 0 Then
        If Not (pData.dwFileAttributes Or FILE_ATTRIBUTE_DIRECTORY = _
           pData.dwFileAttributes) Then
            cFiles.Add Left(pData.cFileName, lstrlen(pData.cFileName))
            pData.cFileName = Space(MAX_PATH)
        End If
        Do While InternetFindNextFile(hFind, pData)
            If Not (pData.dwFileAttributes Or FILE_ATTRIBUTE_DIRECTORY = _
               pData.dwFileAttributes) Then
                cFiles.Add Left(pData.cFileName, lstrlen(pData.cFileName))
                pData.cFileName = Space(MAX_PATH)
            End If
        Loop
        Call InternetCloseHandle(hFind)
    End If
    
    Set EnumFiles = cFiles
    
End Function

Private Function ShowResponse() As Boolean
    Dim lError          As Long
    Dim sDescription    As String
    Dim lRet            As Long
    
    Call InternetGetLastResponseInfo(lError, sDescription, lRet)
    sDescription = Space(lRet)
    Call InternetGetLastResponseInfo(lError, sDescription, lRet)
    Debug.Print lError, sDescription
End Function


