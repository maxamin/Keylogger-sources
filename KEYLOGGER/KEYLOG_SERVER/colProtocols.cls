VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colProtocols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection


Public Property Get ItemExists(Item As Integer) As Boolean
On Error GoTo No

  Dim NewItem As clsProtocol
  
    Set NewItem = mCol.Item(CStr(Item))
    ItemExists = True
No:
End Property


Public Property Get Item(IndexKey As Integer) As clsProtocol
  
  On Error GoTo NoSuchKey

    Set Item = mCol(CStr(IndexKey))
    
Exit Property
NoSuchKey:

End Property


Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(IndexKey As Integer)
    mCol.Remove CStr(IndexKey)
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub



Public Function Add(NewMember As clsProtocol, key As Integer)

    If Len(key) = 0 Then
        mCol.Add NewMember
    Else
        mCol.Add NewMember, CStr(key)
    End If

Exit Function
KeyExists:

End Function


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

