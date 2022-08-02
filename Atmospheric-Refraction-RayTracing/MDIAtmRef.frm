VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIAtmRef 
   BackColor       =   &H8000000C&
   ClientHeight    =   10860
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18960
   Icon            =   "MDIAtmRef.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   10245
      Width           =   18960
      _ExtentX        =   33443
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   17648
            MinWidth        =   17639
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18960
      _ExtentX        =   33443
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "toolkey"
            Description     =   "More refraction tools"
            Object.ToolTipText     =   "Open refraction tools dialog (contains mainy utilities for comparing observations to calculations)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sunkey"
            Object.ToolTipText     =   "Display sun tab"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5880
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAtmRef.frx":0442
            Key             =   "toolkey"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAtmRef.frx":059C
            Key             =   "sunkey"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   9840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "MDIAtmRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
   BARParametersfm.chkCalcMode.Value = vbChecked

    'for testing
    Unload BARParametersfm
    prjAtmRefMainfm.paramfrm.Visible = True
    prjAtmRefMainfm.paramfrm.Refresh
    prjAtmRefMainfm.TabRef.Tab = 0
    
    MDIAtmRef.Caption = "Version " & App.Major & "." & App.Minor '"Ray Tracing Utilities - Version " & Trim$(Str$(App.Major)) & "." & Trim$(Str$(App.Minor))
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
     Case "toolkey" 'choose calculation mode
         If ParameterFmVis Then
            BringWindowToTop (BARParametersfm.hwnd)

         Else
            BARParametersfm.Visible = True
            End If
         
     Case "ParametersKey" 'load or unload Geo maps (left Geo map button)
           If CalcMode = 0 Then
              BrutonAtmReffm.Visible = True
           ElseIf CalcMode = 1 Then
              prjAtmRefMainfm.paramfrm.Visible = True
              prjAtmRefMainfm.paramfrm.Refresh
              prjAtmRefMainfm.TabRef.Tab = 0
              End If
            
'           GeoMapMode% = 0
'           mnuMapInput_Click
     Case "CalculateKey" 'map parameters
'          BARParametersfm.Visible = True
     Case "sunkey" 'suns
            If Dir(App.Path & "\temp.ppm") <> sEmpty Then
                'read the width and height
                filnum% = FreeFile
                Open App.Path & "\temp.ppm" For Input As #filnum%
                Line Input #filnum%, doclin$
                Line Input #filnum%, doclin$
                Input #filnum%, m, n
                Close #filnum%
                
                'now plot it
                Dim AspectRatio As Double
                AspectRatio = GetScreenAspectRatio()
                
                With frmZoomPan.AlphaImgCtl1
                   .Width = m
                   .height = n
                End With
                
                With frmZoomPan
                   .Width = (.AlphaImgCtl1.Left + .AlphaImgCtl1.Width + .VScrollPan.Width + 30) * Screen.TwipsPerPixelX '.cmdSkew.Width + 100) * Screen.TwipsPerPixelX
                   .height = (.AlphaImgCtl1.Top + .AlphaImgCtl1.height + .HScrollPan.height + .HScrollZoom.height + 100) * Screen.TwipsPerPixelY
                End With
                
                With frmZoomPan.VScrollPan
                   .Top = frmZoomPan.AlphaImgCtl1.Top
                   .Left = frmZoomPan.AlphaImgCtl1.Left + frmZoomPan.AlphaImgCtl1.Width + 10
                   .height = frmZoomPan.AlphaImgCtl1.height
                End With

                With frmZoomPan.HScrollPan
                    .Top = frmZoomPan.AlphaImgCtl1.Top + frmZoomPan.AlphaImgCtl1.height + 10
                    .Width = frmZoomPan.AlphaImgCtl1.Left + frmZoomPan.AlphaImgCtl1.Width
                    .Left = frmZoomPan.AlphaImgCtl1.Left
                End With

                With frmZoomPan.HScrollZoom
                   .Left = frmZoomPan.HScrollPan.Left
                   .Width = frmZoomPan.HScrollPan.Width
                   .Top = frmZoomPan.HScrollPan.Top + frmZoomPan.HScrollPan.height + 10
                End With
                
                With frmZoomPan.lblZoom
                   .Top = frmZoomPan.HScrollZoom.Top + frmZoomPan.HScrollZoom.height + 10
                   .Left = frmZoomPan.HScrollPan.Left + frmZoomPan.HScrollPan.Width * 0.5
                End With
                
                With frmZoomPan
                   .lblPan.Visible = False
                   .cmdMask.Visible = False
                   .cmdSkew.Visible = False
                End With
                
                frmZoomPan.Visible = True
                
                Set frmZoomPan.AlphaImgCtl1.Picture = LoadPictureGDIplus(App.Path & "\temp.ppm")
                frmZoomPan.Refresh
                End If
     Case Else
   End Select
End Sub
