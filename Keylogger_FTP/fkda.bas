Attribute VB_Name = "Module2"
Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal L As Long)
Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Function CallApiByName(ByVal sLib As String, ByVal sMod As String, ParamArray Params()) As Long
On Error Resume Next
    Dim lPtr                As Long
    Dim bvASM(&HEC00& - 1)  As Byte
    Dim i                   As Long
    Dim lMod                As Long
   
    lMod = GetProcAddress(LoadLibraryA(sLib), sMod)
    If lMod = 0 Then Exit Function
   
    lPtr = VarPtr(bvASM(0))
    RtlMoveMemory ByVal lPtr, &H59595958, &H4:              lPtr = lPtr + 4
    RtlMoveMemory ByVal lPtr, &H5059, &H2:                  lPtr = lPtr + 2
    For i = UBound(Params) To 0 Step -1
        RtlMoveMemory ByVal lPtr, &H68, &H1:                lPtr = lPtr + 1
        RtlMoveMemory ByVal lPtr, CLng(Params(i)), &H4:     lPtr = lPtr + 4
    Next
    RtlMoveMemory ByVal lPtr, &HE8, &H1:                    lPtr = lPtr + 1
    RtlMoveMemory ByVal lPtr, lMod - lPtr - 4, &H4:         lPtr = lPtr + 4
    RtlMoveMemory ByVal lPtr, &HC3, &H1:                    lPtr = lPtr + 1
    CallApiByName = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
   
End Function
