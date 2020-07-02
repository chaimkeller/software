VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DevInfo_Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'These guys from SP_DEVINFO_DATA
Public Info_DevInst As Long
Public Info_ClassGUID As GUID_Storage

'These guys from SetupDiGetDeviceRegistryProperty()
Public SPDRP_ADDRESS As Long
Public SPDRP_BUSNUMBER As Long
Public SPDRP_BUSTYPEGUID As GUID_Storage
Public SPDRP_CAPABILITIES As Long
Public SPDRP_CHARACTERISTICS As Long
Public SPDRP_CLASS As String
Public SPDRP_CLASSGUID As String
Public SPDRP_COMPATIBLEIDS As String 'multi string
Public SPDRP_CONFIGFLAGS As Long
'public SPDRP_DEVICE_POWER_DATA as CM_POWER_DATA
Public SPDRP_DEVICEDESC As String
Public SPDRP_DEVTYPE As Long
Public SPDRP_DRIVER As String
Public SPDRP_ENUMERATOR_NAME As String
Public SPDRP_EXCLUSIVE As Long
Public SPDRP_FRIENDLYNAME As String
Public SPDRP_HARDWAREID As String 'multi string
Public SPDRP_INSTALL_STATE As Long
'public SPDRP_LEGACYBUSTYPE as ?
Public SPDRP_LOCATION_INFORMATION As String
'public SPDRP_LOCATION_PATHS() as String 'Win2003+
Public SPDRP_LOWERFILTERS As String 'multi string
Public SPDRP_MFG As String
Public SPDRP_PHYSICAL_DEVICE_OBJECT_NAME As String
Public SPDRP_REMOVAL_POLICY As Long 'WinXP+
Public SPDRP_REMOVAL_POLICY_HW_DEFAULT As Long 'WinXP+
Public SPDRP_REMOVAL_POLICY_OVERRIDE As Long 'WinXP+
'public SPDRP_SECURITY as SECURITY_DESCRIPTOR
Public SPDRP_SECURITY_SDS As String
Public SPDRP_SERVICE As String
Public SPDRP_UI_NUMBER As Long
Public SPDRP_UI_NUMBER_DESC_FORMAT As String
Public SPDRP_UPPERFILTERS As String 'multi string

Private Sub Class_Initialize()
    Set Info_ClassGUID = New GUID_Storage
    Set SPDRP_BUSTYPEGUID = New GUID_Storage
End Sub

Private Sub Class_Terminate()
    Set Info_ClassGUID = Nothing
    Set SPDRP_BUSTYPEGUID = Nothing
End Sub

Public Function GetName() As String
    Dim temp As String
    temp = SPDRP_FRIENDLYNAME
    If temp = "" Then temp = SPDRP_DEVICEDESC
    GetName = temp
End Function
