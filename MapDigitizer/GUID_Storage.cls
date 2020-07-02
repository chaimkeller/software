VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUID_Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Cannot pass public types between classes due to compiler limitation
' So create a class wrapper for passing around GUIDs

Option Explicit

Private Type intGUIDtype
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private intGUID As intGUIDtype

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'Pointer based copying
Public Sub CopyFromThisPTR(PtrToSrcGUID As Long)
    CopyMemory intGUID, ByVal PtrToSrcGUID, Len(intGUID)
End Sub
Public Sub CopyToThisPTR(PtrToDstGUID As Long)
    CopyMemory ByVal PtrToDstGUID, intGUID, Len(intGUID)
End Sub

'Other Class Instance copying
Public Sub CopyFromThisClass(Src As GUID_Storage)
    'Dim i As Integer
    'Me.Data1 = Src.Data1
    'Me.Data2 = Src.Data2
    'Me.Data3 = Src.Data3
    'For i = 0 To 7
    '    Me.Data4(i) = Src.Data4(i)
    'Next i
    Me.CopyFromThisPTR Src.Ptr
End Sub

'By law, a compare operation
Public Function IsEqual(V As GUID_Storage) As Boolean
    Dim temp As Boolean
    Dim i As Integer
    temp = True
    If Me.Data1 <> V.Data1 Then temp = False
    If Me.Data2 <> V.Data2 Then temp = False
    If Me.Data3 <> V.Data3 Then temp = False
    For i = 0 To 7
        If Me.Data4(i) <> V.Data4(i) Then temp = False
    Next i
    IsEqual = temp
End Function

Public Function Ptr() As Long
    Ptr = VarPtr(intGUID)
End Function

Public Sub Clear()
    Dim i As Integer
    intGUID.Data1 = 0
    intGUID.Data2 = 0
    intGUID.Data3 = 0
    For i = 0 To 7
        intGUID.Data4(i) = 0
    Next i
End Sub

Public Function ToString(Optional Delim As String = ":") As String
    Dim temp As String, i As Integer
    temp = PadStr(Hex(intGUID.Data1), 8) & Delim
    temp = temp & PadStr(Hex(intGUID.Data2), 4) & Delim
    temp = temp & PadStr(Hex(intGUID.Data3), 4) & Delim
    For i = 0 To 7
        temp = temp & PadStr(Hex(intGUID.Data4(i)), 2)
        If i < 7 Then temp = temp & Delim
    Next i
    ToString = temp
End Function

Private Sub Class_Initialize()
    Clear
End Sub

Public Property Get Data1() As Long
    Data1 = intGUID.Data1
End Property
Public Property Let Data1(V As Long)
    intGUID.Data1 = V
End Property

Public Property Get Data2() As Integer
    Data2 = intGUID.Data2
End Property
Public Property Let Data2(V As Integer)
    intGUID.Data2 = V
End Property

Public Property Get Data3() As Integer
    Data3 = intGUID.Data3
End Property
Public Property Let Data3(V As Integer)
    intGUID.Data3 = V
End Property

Public Property Get Data4(Index As Integer) As Byte
    Data4 = intGUID.Data4(Index)
End Property
Public Property Let Data4(Index As Integer, V As Byte)
    intGUID.Data4(Index) = V
End Property

Public Function PadStr(sInput As String, size As Integer, Optional PadChar As String = "0") As String
    Dim x As Integer, temp As String
    
    temp = sInput
    For x = 1 To (size - Len(sInput))
        temp = PadChar & temp
    Next x
    PadStr = temp
End Function

