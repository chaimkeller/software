VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GDCoordinateList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bat file coordinate plot"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Change color of markers"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Clear the list"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select all"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Select all the listed points"
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Plot"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Plot the selected points"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Browse for  bat file"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ListBox CoordList 
      Height          =   4785
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog comDlgList 
      Left            =   3840
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "GDCoordinateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
   Dim i%
   
   If GDCoordinateList.cmdAll.Caption = "Select all" Then
        For i% = 1 To GDCoordinateList.CoordList.ListCount
           GDCoordinateList.CoordList.Selected(i% - 1) = True
        Next i%
        GDCoordinateList.cmdAll.Caption = "Clear all"
   ElseIf GDCoordinateList.cmdAll.Caption = "Clear all" Then
        For i% = 1 To GDCoordinateList.CoordList.ListCount
           GDCoordinateList.CoordList.Selected(i% - 1) = False
        Next i%
        GDCoordinateList.cmdAll.Caption = "Clear all"
      End If
      GDCoordinateList.cmdAll.Caption = "Select all"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdBrowse_Click
' Author    : Dr-John-K-Hall
' Date      : 11/1/2018
' Purpose   : Plot out bat file coordinates
'---------------------------------------------------------------------------------------
'
Private Sub cmdBrowse_Click()
  Dim BatFileName As String
  Dim BatFilnum%
  Dim BatLineParse() As String
  Dim doclin$
  
   On Error GoTo cmdBrowse_Click_Error

  comDlgList.CancelError = True
  comDlgList.Filter = "Bat files (*.bat)|*.bat|" & _
                     "All files (*.*)|*.*"
  'specify default filter
  comDlgList.FilterIndex = 1
  comDlgList.FileName = "*.bat"
  comDlgList.ShowOpen
  BatFileName = comDlgList.FileName
  'check that it exists
  If Dir(BatFileName) = sEmpty Then
     Call MsgBox("Can't find the file", vbCritical, "file missing")
     Exit Sub
     End If
     
   BatFilnum% = FreeFile
   Open BatFileName For Input As #BatFilnum%
   Do Until EOF(BatFilnum%)
      Line Input #BatFilnum%, doclin$
      'parse it out (comma dilimmeted)
      BatLineParse = Split(doclin$, ",")
      If InStr(LCase(BatLineParse(0)), "netz") Or InStr(LCase(BatLineParse(0)), "skiy") Then
         'add the line to the list
         CoordList.AddItem doclin$
         End If
   Loop
   Close #BatFilnum%

   On Error GoTo 0
   Exit Sub

cmdBrowse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBrowse_Click of Form GDCoordinateList"
End Sub

Private Sub cmdClear_Click()
  'clear list
  GDCoordinateList.CoordList.Clear
  
'Redraw map

ier = ReDrawMap(0)

If DigitizeOn Then
   If Not InitDigiGraph Then
      InputDigiLogFile 'load up saved digitizing data for the current map sheet
   Else
      ier = RedrawDigiLog
      End If
   End If
   
End Sub

Private Sub cmdColor_Click()

    comDlgList.CancelError = True
    On Error Resume Next

    comDlgList.ShowColor
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting color." & vbCrLf & Err.Description
        Exit Sub
    End If

    MarkerColor = comDlgList.color

   On Error GoTo 0
   Exit Sub

      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdDraw_Click
' Author    : Dr-John-K-Hall
' Date      : 11/1/2018
' Purpose   : Draws out checked bat file items
'---------------------------------------------------------------------------------------
'
Public Sub cmdDraw_Click()

   Dim i%, j%
   Dim ParseCoordString() As String, PrintLabel As String
   
   Dim GeoX As Double, GeoY As Double
   Dim GeoToPixelX As Double, GeoToPixelY As Double
   Dim CurrentX As Long, CurrentY As Long
   Dim LastDrawMode As Long, LastFillMode As Long, LastFillColor As Long
   Dim LastPrintFont As String, LastPrintSize As Long, LastFontBold As Boolean, LastFontColor As Long
   
   On Error GoTo cmdDraw_Click_Error
   
   If Not DigiRubberSheeting Then
      Call MsgBox("You first must activate rubber sheeting", vbInformation, "rubber sheeting error")
      Exit Sub
      End If
    
    GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
    GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
    
    LastDrawMode = GDform1.Picture2.DrawMode
    LastFillMode = GDform1.Picture2.FillStyle
    LastFillColor = GDform1.Picture2.FillColor
    LastPrintFont = GDform1.Picture2.FontName
    LastPrintSize = GDform1.Picture2.FontSize
    LastFontBold = GDform1.Picture2.FontBold
    LastFontColor = GDform1.Picture2.ForeColor
    
    GDform1.Picture2.DrawMode = 13
    GDform1.Picture2.FillStyle = 0
    GDform1.Picture2.FillColor = MarkerColor
    GDform1.Picture2.FontName = "Times New Roman"
    GDform1.Picture2.FontSize = 12
    GDform1.Picture2.FontBold = False
    GDform1.Picture2.ForeColor = MarkerColor 'QBColor(4)
    
    Screen.MousePointer = vbHourglass
    
    'first Redraw map
    If Not CoordListZoom Then
        ier = ReDrawMap(0)
        End If
        
    'redraw digitilization if flagged
    If DigitizeOn Then
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
       End If

    For i% = 1 To GDCoordinateList.CoordList.ListCount
      If GDCoordinateList.CoordList.Selected(i% - 1) Then
         'draw that item
         'parse out the coordinates
         ParseCoordString = Split(GDCoordinateList.CoordList.List(i% - 1), ",")
         GeoY = val(ParseCoordString(1))
         GeoX = -val(ParseCoordString(2))
         If Not IsNumeric(GeoX) Or Not IsNumeric(GeoY) Then
            Call MsgBox("Not valid coordinates", vbCritical, "coordinate error")
            Exit Sub
            End If
         
         CurrentX = ((GeoX - ULGeoX) * GeoToPixelX) + ULPixX
         CurrentY = ((ULGeoY - GeoY) * GeoToPixelY) + ULPixY
         If DigiZoom.Zoom <> 1 Then
            CurrentX = CurrentX * DigiZoom.Zoom
            CurrentY = CurrentY * DigiZoom.Zoom
            End If
         
         'also print out name
         
         GDform1.Picture2.CurrentX = CurrentX + Max(2, CInt(DigiZoom.LastZoom))
         GDform1.Picture2.CurrentY = CurrentY
         For j% = Len(ParseCoordString(0)) To 1 Step -1
            If Mid$(ParseCoordString(0), j%, 1) = "." Then
               PrintLabel = Mid$(ParseCoordString(0), j% + 1, Len(ParseCoordString(0)) - j%)
               Exit For
               End If
         Next j%
         
         GDform1.Picture2.Print PrintLabel

         GDform1.Picture2.Circle (CurrentX, CurrentY), Max(2 * DigiZoom.Zoom, CInt(DigiZoom.LastZoom)), MarkerColor
         
         End If
         
     Next i%
     
     If DigiZoom.Zoom > 0.99 And DigiZoom.Zoom < 1.01 Then
        'center at last coordinate plotted
        GDMDIform.Text5.Text = Format(str$(GeoX), "#######.####0")
        GDMDIform.Text6.Text = Format(str$(GeoY), "#######.####0")
        Call gotocoord
        End If
     
     Screen.MousePointer = vbDefault
         
     'reset defaults
     GDform1.Picture2.DrawMode = LastDrawMode
     GDform1.Picture2.FillStyle = LastFillMode
     GDform1.Picture2.FillColor = LastFillColor
     GDform1.Picture2.FontName = LastPrintFont
     GDform1.Picture2.FontSize = LastPrintSize
     GDform1.Picture2.FontBold = LastFontBold
     GDform1.Picture2.ForeColor = LastFontColor

   On Error GoTo 0
   Exit Sub

cmdDraw_Click_Error:

     Screen.MousePointer = vbDefault

     GDform1.Picture2.DrawMode = LastDrawMode
     GDform1.Picture2.FillStyle = LastFillMode
     GDform1.Picture2.FillColor = LastFillColor
     GDform1.Picture2.FontName = LastPrintFont
     GDform1.Picture2.FontSize = LastPrintSize
     GDform1.Picture2.FontBold = LastFontBold
     GDform1.Picture2.ForeColor = LastFontColor

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDraw_Click of Form GDCoordinateList"
End Sub

Private Sub Form_Load()
   Call sCenterForm(Me)
   MarkerColor = QBColor(12)
   CoordListVis = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbDefault
   Set GDCoordinateList = Nothing
   CoordListVis = False
End Sub
