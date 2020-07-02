VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://support.microsoft.com/kb/259695

'====== TYPES AND DECLARES
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Type SP_DEVINFO_DATA
    cbSize As Long
    ClassGUID As GUID
    DevInst As Long
    Reserved As Long
End Type
Private Declare Function SetupDiGetClassDevs Lib "setupapi" Alias "SetupDiGetClassDevsA" (ByRef ClassGUID As GUID, ByVal Enumerator As Long, ByVal hwndParent As Long, ByVal Flags As Long) As Long
Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef PropertyRegDataType As Long, ByVal PtrPropertyBuffer As Long, ByVal PropertyBufferSize As Long, ByRef RequiredSize As Long) As Long
Private Declare Function SetupDiClassGuidsFromName Lib "setupapi" Alias "SetupDiClassGuidsFromNameA" (ByVal ClassName As String, ByVal PtrClassGUIDList As Long, ByVal GUIDListSize As Long, ByRef ReqGUIDListSize As Long) As Long
Private Declare Function SetupDiClassNameFromGuid Lib "setupapi" Alias "SetupDiClassNameFromGuidW" (ByRef ClassGUID As GUID, ByVal StrPtrClassName As Long, ByVal NameSize As Long, ByRef ReqNameSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi" (ByVal DevInfoSet As Long) As Long
Private Const DIGCF_PRESENT As Long = &H2
Private Const DIGCF_DEVICEINTERFACE As Long = &H10
Private Const DIGCF_ALLCLASSES As Long = &H4
Private Const DIGCF_PROFILE As Long = &H8
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const ERROR_INSUFFICIENT_BUFFER = 122
Private Const ERROR_INVALID_DATA = 13

Private Const SPDRP_DEVICEDESC = &H0                         ' DeviceDesc =R/W
Private Const SPDRP_HARDWAREID = &H1                         ' HardwareID =R/W
Private Const SPDRP_COMPATIBLEIDS = &H2                      ' CompatibleIDs =R/W
Private Const SPDRP_NTDEVICEPATHS = &H3                      ' Unsupported, DO NOT USE
Private Const SPDRP_SERVICE = &H4                            ' Service =R/W
Private Const SPDRP_CONFIGURATION = &H5                      ' Configuration =R
Private Const SPDRP_CONFIGURATIONVECTOR = &H6                ' ConfigurationVector =R
Private Const SPDRP_CLASS = &H7                              ' Class =R--tied to ClassGUID
Private Const SPDRP_CLASSGUID = &H8                          ' ClassGUID =R/W
Private Const SPDRP_DRIVER = &H9                             ' Driver =R/W
Private Const SPDRP_CONFIGFLAGS = &HA                        ' ConfigFlags =R/W
Private Const SPDRP_MFG = &HB                                ' Mfg =R/W
Private Const SPDRP_FRIENDLYNAME = &HC                       ' FriendlyName =R/W
Private Const SPDRP_LOCATION_INFORMATION = &HD               ' LocationInformation =R/W
Private Const SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = &HE        ' PhysicalDeviceObjectName =R
Private Const SPDRP_CAPABILITIES = &HF                       ' Capabilities =R
Private Const SPDRP_UI_NUMBER = &H10                         ' UiNumber =R
Private Const SPDRP_UPPERFILTERS = &H11                      ' UpperFilters =R/W
Private Const SPDRP_LOWERFILTERS = &H12                      ' LowerFilters =R/W
Private Const SPDRP_MAXIMUM_PROPERTY = &H13                  ' Upper bound on ordinals

'====== GLOBAL VARS

Private DeviceList As Collection

Private Sub Class_Initialize()
    Set DeviceList = New Collection
End Sub

Private Sub Class_Terminate()
    Set DeviceList = Nothing
End Sub

Public Sub ListClear()
    Set DeviceList = Nothing
    Set DeviceList = New Collection
End Sub

Public Function ListCount() As Long
    ListCount = DeviceList.Count
End Function

Public Function Item(Index As Long) As DevInfo_Storage
    Set Item = DeviceList(Index)
End Function

'Give Class Name get GUID list, returns an error string or OK
Public Function GetGUIDByName(ClassName As String, ByRef ClassGUID() As GUID_Storage) As String
    Dim GUIDs() As GUID, GUIDsize As Long, ReqBuffSize As Long
    Dim i As Long
    Dim ReturnStr As String
    ReturnStr = "OK"
    'call first to get size
    GUIDsize = 0
    i = SetupDiClassGuidsFromName(ClassName, 0, 0, ReqBuffSize) '<> 0 Then
        If ReqBuffSize = 0 Then
            ReturnStr = "Incorrect class name"
        Else
            'size up and get the name this time
            GUIDsize = ReqBuffSize
            ReDim GUIDs(0 To GUIDsize - 1)
            If SetupDiClassGuidsFromName(ClassName, GetAddressOfGUIDArray1D(GUIDs), GUIDsize, ReqBuffSize) <> 0 Then
                ReDim ClassGUID(0 To GUIDsize - 1)
                For i = LBound(ClassGUID) To UBound(ClassGUID)
                    Set ClassGUID(i) = New GUID_Storage
                    ClassGUID(i).CopyFromThisPTR VarPtr(GUIDs(i))
                Next i
            Else
                ReturnStr = DecodeGetLastError(GetLastError)
            End If
        End If
    'Else
    '    ReturnStr = DecodeGetLastError(GetLastError)
    'End If
    GetGUIDByName = ReturnStr
End Function

'Give GUID get Class Name, returns an error string or OK
Public Function GetNameByGUID(ClassGUID As GUID_Storage, ByRef ClassName As String) As String
    Dim ReturnStr As String, aGUID As GUID
    Dim GetClassName As String, ReqBuffSize As Long
    ReturnStr = "OK"
    ClassName = ""
    
    ClassGUID.CopyToThisPTR VarPtr(aGUID)
    GetClassName = Space$(500) ' hope thats big enough
    If SetupDiClassNameFromGuid(aGUID, StrPtr(GetClassName), Len(GetClassName), ReqBuffSize) Then
        ClassName = StrTrimToNull(Trim$(GetClassName))
    Else
        ReturnStr = DecodeGetLastError(GetLastError)
    End If
    GetNameByGUID = ReturnStr
End Function

'THIS ONE JUST WON'T WORK, use presumption version until fixed if ever
Public Function GetNameByGUID2(ClassGUID As GUID_Storage, ByRef ClassName As String) As String
    Dim ReturnStr As String, aGUID As GUID
    Dim GetClassName As String, ReqBuffSize As Long
    ReturnStr = "OK"
    ClassName = ""
    
    ClassGUID.CopyToThisPTR VarPtr(aGUID)
    Debug.Print ClassGUID.ToString
    Debug.Print GUIDtoString(aGUID)
    Call SetupDiClassNameFromGuid(aGUID, 0, 0, ReqBuffSize) 'Then
        If ReqBuffSize = 0 Then
            ReturnStr = "Class not found"
        Else
            GetClassName = Space$(ReqBuffSize)
            If SetupDiClassNameFromGuid(aGUID, StrPtr(GetClassName), Len(GetClassName), ReqBuffSize) Then
                ClassName = GetClassName
            Else
                ReturnStr = DecodeGetLastError(GetLastError)
            End If
        End If
    'Else
    '    ReturnStr = DecodeGetLastError(GetLastError)
    'End If
    GetNameByGUID2 = ReturnStr
End Function

'This adds devices to the list. If you supply a class GUID, it will list only those devices.
' Otherwise it will list all. It accumulates so if you don't want it to, call ListClear first
Public Function AddToList(Optional ClassID As GUID_Storage = Nothing, Optional bOnlyPresent As Boolean = True) As String
    Dim hDevInfo As Long, i As Long
    Dim DevInfoData As SP_DEVINFO_DATA
    Dim ClassGUID As GUID, ReturnStr As String, CurDevInfo As DevInfo_Storage
    Dim StrTemp As String, DwTemp As Long
    Dim Flags As Long
    hDevInfo = INVALID_HANDLE_VALUE
    ReturnStr = "OK"
    
    DevInfoData.cbSize = Len(DevInfoData)
    
    If bOnlyPresent Then
        Flags = DIGCF_PRESENT
    Else
        Flags = 0
    End If
    
    If ClassID Is Nothing Then
        ClearGUID ClassGUID
        hDevInfo = SetupDiGetClassDevs(ClassGUID, ByVal 0&, ByVal 0&, Flags Or DIGCF_ALLCLASSES)
    Else
        ClassID.CopyToThisPTR VarPtr(ClassGUID)
        hDevInfo = SetupDiGetClassDevs(ClassGUID, ByVal 0&, ByVal 0&, Flags)
    End If
    If hDevInfo <> INVALID_HANDLE_VALUE Then
        i = 0
        While SetupDiEnumDeviceInfo(hDevInfo, i, DevInfoData)
            'start new entry
            Set CurDevInfo = New DevInfo_Storage
            'copy stuff from SP_DEVINFO_DATA struct that may be of use
            CurDevInfo.Info_ClassGUID.CopyFromThisPTR VarPtr(DevInfoData.ClassGUID)
            CurDevInfo.Info_DevInst = DevInfoData.DevInst
            'If i < 5 Then
            'Debug.Print i
            'Debug.Print CurDevInfo.Info_ClassGUID.ToString
            'Debug.Print GUIDtoString(DevInfoData.ClassGUID)
            'End If
            
            'Get specific values from API query, Implement what you need
            ReturnStr = GetDevRegisterValue(hDevInfo, DevInfoData, SPDRP_DEVICEDESC, StrTemp, DwTemp)
            If ReturnStr = "OK" Then
                CurDevInfo.SPDRP_DEVICEDESC = StrTemp
            Else
                ReturnStr = "SPDRP_DEVICEDESC " & ReturnStr
            End If
            
            ReturnStr = GetDevRegisterValue(hDevInfo, DevInfoData, SPDRP_FRIENDLYNAME, StrTemp, DwTemp)
            If ReturnStr = "OK" Then
                CurDevInfo.SPDRP_FRIENDLYNAME = StrTemp
            Else
                ReturnStr = "SPDRP_FRIENDLYNAME " & ReturnStr
            End If
            
            ReturnStr = GetDevRegisterValue(hDevInfo, DevInfoData, SPDRP_CLASS, StrTemp, DwTemp)
            If ReturnStr = "OK" Then
                CurDevInfo.SPDRP_CLASS = StrTemp
            Else
                ReturnStr = "SPDRP_CLASS " & ReturnStr
            End If
            
            'Add it the collection
            DeviceList.Add CurDevInfo
            i = i + 1
        Wend
        'SetupDiDestroyDeviceInfoList hDevInfo
    Else
        ReturnStr = "(BAD HANDLE) " & DecodeGetLastError(GetLastError)
    End If
    
EndHere:
    If hDevInfo <> INVALID_HANDLE_VALUE Then SetupDiDestroyDeviceInfoList hDevInfo
    AddToList = ReturnStr
End Function

Private Function GetDevRegisterValue(ByVal hDevInfo As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal Property As Long, ByRef StrValue As String, ByRef DwordValue As Long) As String
    Dim DataT As Long, BuffSize As Long, ReqBuffSize As Long, Ret As Long
    Dim Buffer() As Byte, ReturnStr As String
    DataT = 0
    BuffSize = 500
    ReDim Buffer(0 To BuffSize * 2 - 1)
    StrValue = "" 'establish defaults incase we got nothing
    DwordValue = 0
    Buffer(0) = 0 'clear buffer string
    ReturnStr = "OK"
    
    Ret = SetupDiGetDeviceRegistryProperty(hDevInfo, DeviceInfoData, Property, DataT, GetAddressOfByteArray1D(Buffer), BuffSize, ReqBuffSize)
    If Ret = 0 Then
        'ReturnStr = DecodeGetLastError(GetLastError)
        'Data may be invalid for this device, not a critical error
        GoTo EndHere
    End If
    If ReqBuffSize > BuffSize Then
        ReturnStr = "Buffer too small " & CStr(ReqBuffSize) & " vs " & CStr(BuffSize)
        GoTo EndHere
    End If
    If Ret = ERROR_INVALID_DATA Then 'this device doesn't have this property
        GoTo EndHere
    End If
    'string value
    StrValue = GetStringFromBytes(Buffer, ReqBuffSize - 1)
    'Long (DWORD) value
    CopyMemory ByVal VarPtr(DwordValue), ByVal GetAddressOfByteArray1D(Buffer), Len(DwordValue)
EndHere:
    GetDevRegisterValue = ReturnStr
End Function

Private Sub ClearGUID(ByRef V As GUID)
    Dim i As Integer
    V.Data1 = 0
    V.Data2 = 0
    V.Data3 = 0
    For i = 0 To 7
        V.Data4(i) = 0
    Next i
End Sub

Private Function GetAddressOfByteArray1D(Var() As Byte) As Long
    Dim SA1D As SAFEARRAY1D
    Dim SA1D_Ptr As Long
    
    'copy SA-1D structure out of bar()
    CopyMemory SA1D_Ptr, ByVal VarPtrArray(Var), 4
    CopyMemory SA1D, ByVal SA1D_Ptr, Len(SA1D)
    'copy data
     GetAddressOfByteArray1D = SA1D.pvData
End Function

'This was a lovely example in API Guide
Private Function DecodeGetLastError(Error As Long) As String
    Dim Buffer As String
    'Create a string buffer
    Buffer = Space(200)
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Error, LANG_NEUTRAL, Buffer, 200, ByVal 0&
    'Show the message
    DecodeGetLastError = StrTrimToNull(Trim$(Buffer))
End Function

Private Function GetStringFromBytes(Var() As Byte, Optional MaxLen As Long = -1) As String
    Dim temp As String, i As Long, Cnt As Long
    On Error GoTo EmptyArr
    temp = ""
    Cnt = 0
    For i = LBound(Var) To UBound(Var)
        If MaxLen > -1 And Cnt = MaxLen Then Exit For
        'If Var(i) = 0 Then Exit For
        temp = temp & Chr$(Var(i))
        Cnt = Cnt + 1
    Next i
    GetStringFromBytes = temp
    Exit Function
EmptyArr:
    GetStringFromBytes = ""
End Function

Private Function GetAddressOfGUIDArray1D(Var() As GUID) As Long
    Dim SA1D As SAFEARRAY1D
    Dim SA1D_Ptr As Long
    
    'copy SA-1D structure out of bar()
    CopyMemory SA1D_Ptr, ByVal VarPtrArray(Var), 4
    CopyMemory SA1D, ByVal SA1D_Ptr, Len(SA1D)
    'copy data
     GetAddressOfGUIDArray1D = SA1D.pvData
End Function

Public Function StrTrimToNull(V As String) As String
    Dim i As Long
    i = InStr(1, V, Chr$(0))
    If i > 0 Then
        StrTrimToNull = Left$(V, i - 1)
    Else
        StrTrimToNull = V
    End If
End Function

'for internal debugging
Private Function GUIDtoString(V As GUID) As String
    Dim temp As String, i As Integer
    temp = PadStr(Hex(V.Data1), 8) & ":"
    temp = temp & PadStr(Hex(V.Data2), 4) & ":"
    temp = temp & PadStr(Hex(V.Data3), 4) & ":"
    For i = 0 To 7
        temp = temp & PadStr(Hex(V.Data4(i)), 2)
        If i < 7 Then temp = temp & ":"
    Next i
    GUIDtoString = temp
End Function

'for some reason using Format() was not always consistent
Public Function PadStr(sInput As String, size As Integer, Optional PadChar As String = "0") As String
    Dim x As Integer, temp As String
    
    temp = sInput
    For x = 1 To (size - Len(sInput))
        temp = PadChar & temp
    Next x
    PadStr = temp
End Function


