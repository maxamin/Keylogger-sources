Attribute VB_Name = "modFunctions"
Option Explicit

'String functions
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Long, ByVal ByteLen As Long)

Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long


Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767


Public Function UnsignedToLong(Value As Double) As Long
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    UnsignedToLong = Value - IIf(Value <= MAXINT_4, 0, OFFSET_4)
End Function


Public Function LongToUnsigned(Value As Long) As Double
    LongToUnsigned = Value + IIf(Value < 0, OFFSET_4, 0)
End Function


Public Function UnsignedToInteger(Value As Long) As Integer
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    UnsignedToInteger = Value - IIf(Value <= MAXINT_2, 0, OFFSET_2)
End Function


Public Function IntegerToUnsigned(Value As Integer) As Long
    IntegerToUnsigned = Value + IIf(Value < 0, OFFSET_2, 0)
End Function


Public Function HiWord(lngValue As Long) As Long
    HiWord = (lngValue And &HFFFF0000) \ &H40000
End Function


Public Function LoWord(lngValue As Long) As Long
    LoWord = lngValue And &HFFFF
End Function


Public Function MakeDWord(HiWord As Integer, LoWord As Integer) As Long
    MakeDWord = LoWord
    CopyMemory ByVal VarPtr(MakeDWord) + 2, ByVal VarPtr(HiWord), 2
End Function


Public Function HiByte(intValue As Integer) As Long
    HiByte = (intValue And &HFF00) \ 256
End Function


Public Function LoByte(intValue As Integer) As Long
    LoByte = intValue And &HFF
End Function


Public Function MakeWord(HiByte As Byte, LoByte As Byte) As Integer
    MakeWord = LoByte
    CopyMemory ByVal VarPtr(MakeWord) + 1, ByVal VarPtr(HiByte), 1
End Function


Public Function HiNibble(bytValue As Byte) As Byte
    HiNibble = (bytValue And &HF0) \ 16
End Function


Public Function LoNibble(bytValue As Byte) As Byte
    LoNibble = (bytValue And &HF)
End Function


Public Function MakeByte(HiNibble As Byte, LoNibble As Byte) As Integer
    MakeByte = LoNibble
    MakeByte = MakeByte Or (HiNibble * 2)
End Function


Public Function StringFromPointer(ByVal lPointer As Long) As String

  Dim strTemp As String
  Dim lRetVal As Long

    'prepare the strTemp buffer
    strTemp = String$(lstrlen(ByVal lPointer), 0)

    'copy the string into the strTemp buffer
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)

    'return a string
    If lRetVal Then StringFromPointer = strTemp
End Function
