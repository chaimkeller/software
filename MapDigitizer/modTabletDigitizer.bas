Attribute VB_Name = "modTabletDigitizer"
Public DigiMouseUp As Boolean
Public DigiMouseDown As Boolean
Public DigiDrag As Boolean
Public DigiDrag1 As POINTAPI
Public DigiDrag2 As POINTAPI
Public DigiTime1 As Long
Public DigiTime2 As Long
Public DigiButton As Boolean

Public field_xSize As Long
Public field_ySize As Long

Public Message_Text As String 'Used to create text strings.
Public Command_Called As String 'Used to display the command called in the Error MessageBox Caption.

'opens communication to the GTCO table digitizer
'if communication can't be established, an error message is returned
Public Sub OpenTablet()
On Error GoTo Exit_Sub

    DigiTime1 = INIT_VALUE
    DigiTime2 = INIT_VALUE
    DigiDrag1.x = INIT_VALUE
    DigiDrag1.y = INIT_VALUE
    DigiDrag2.x = INIT_VALUE
    DigiDrag2.y = INIT_VALUE

    
    Command_Called = "TabletOpen" 'Error Caption
    'This line Opens the the tablet and returns true or false.
    
    Command_Called = "Set my_TabletControl = New TabletControl" 'Error Caption
    'This line creates a communication link to the Tablet Interface Control
    Set my_TabletControl = New VBTabletControl
    
    Get_Tablet_Properties
    Exit Sub
    
Exit_Sub:
    MsgBox Err.Description, , Command_Called
    
    Err.Clear
End Sub
'closes communication to Tablet
Public Sub CloseTablet()
On Error GoTo Exit_Sub

    Command_Called = "TabletClose"
    'This line disconnects the device.
    Call my_TabletControl.TabletClose
    
    Exit Sub
Exit_Sub:
    MsgBox Err.Description, , Command_Called
End Sub
Public Sub Get_Tablet_Properties()
On Error GoTo Exit_Sub
Dim rc As Long
Dim XOrg As Long
Dim YOrg As Long
Dim XExt As Long
Dim YExt As Long

'These Commands can be called at any time after TabletOpen
        
'        Command_Called = "TabletGetMake" 'Error Caption
'    field_Make = my_TabletControl.TabletGetMake()
'
'        Command_Called = "TabletGetModel" 'Error Caption
'    field_Model = my_TabletControl.TabletGetModel()
'
'        Command_Called = "TabletGetInterface" 'Error Caption
'    Dim Interface_Type As Long
'    Message_Text = my_TabletControl.TabletGetInterface(Interface_Type)
'    ' Note: InterfaceID is not reliable when returned from TabletGetInterface.
'    ' Instead, use the new TabletGetInterfaceID
'    Interface_Type = my_TabletControl.TabletGetInterfaceID
'    field_Interface = Interface_Type & ", " & Message_Text
'
'        Command_Called = "TabletGetAddress" 'Error Caption
'    field_Address = my_TabletControl.TabletGetAddress()
'
'        Command_Called = "TabletGetBatteryLevel" 'Error Caption
'    field_BatteryLevel = my_TabletControl.TabletGetBatteryLevel()
'
'        Command_Called = "TabletGetIDString" 'Error Caption
'    field_IDString = my_TabletControl.TabletGetIDString()
'
'        Command_Called = "TabletGetVersion" 'Error Caption
'    field_Version = my_TabletControl.TabletGetVersion()
'
'        Command_Called = "TabletGetNetworkName" 'Error Caption
'    field_NetworkName = my_TabletControl.TabletGetNetworkName()
        
        Command_Called = "TabletGetXSize" 'Error Caption
    field_xSize = my_TabletControl.TabletGetXSize()
        
        Command_Called = "TabletGetYSize" 'Error Caption
    field_ySize = my_TabletControl.TabletGetYSize()
        
'        Command_Called = "TabletGetZSize" 'Error Caption
'    field_zSize = my_TabletControl.TabletGetZSize()
'
'        Command_Called = "TabletResolution" 'Error Caption
'    field_Resolution = my_TabletControl.TabletResolution()
'
'        Command_Called = "TabletUnits" 'Error Caption
'    field_UnitOfMeasure = my_TabletControl.TabletUnits()
'
'        Command_Called = "TabletMouseArea" 'Error Caption
'    Call my_TabletControl.TabletGetMouseArea(XOrg, YOrg, XExt, YExt)
'    field_MouseArea = "{" + str(XOrg) + "," + str(YOrg) + "} {" + str(XOrg + XExt) + "," + str(YOrg + YExt) + "}"
'
'    Call Get_Transducer_Properties
'
'    Enable_Form_Fields True
    
    Exit Sub
Exit_Sub:
    MsgBox Err.Description, , Command_Called
    Err.Clear
    Resume Next
End Sub

Public Function HexStr$(a As Long)
'    HexStr$ = "0x" + Hex$(a \ &H1000) + Hex$((a Mod &H1000) \ &H100) + Hex$((a Mod &H100) \ &H10) + Hex$(a Mod &H10)
    HexStr$ = Hex$(a \ &H1000) + Hex$((a Mod &H1000) \ &H100) + Hex$((a Mod &H100) \ &H10) + Hex$(a Mod &H10)
End Function

