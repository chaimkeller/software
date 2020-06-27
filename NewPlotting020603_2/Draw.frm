VERSION 5.00
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Plot"
   ClientHeight    =   8190
   ClientLeft      =   1650
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "Draw.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10215
   Begin VB.Label lblCoord 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape shpMouseDraw 
      Height          =   495
      Left            =   1380
      Top             =   6960
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fStartXMD As Single 'startposition X of rectangle
Public fStartYMD As Single 'startposition Y of rectangle
Public fHeightMD As Single 'difference Y direction of rectangle
Public fWidthMD As Single 'difference X direction of rectangle

Private Sub Form_DblClick()

   DblClickForm

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim udtMyCoord As COORDINATE
Dim sCoord As String 'containing coordinate-info
Dim nLenStr As Integer 'length of string coordinate-info

If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  shpMouseDraw.Left = X
  fStartXMD = X
  shpMouseDraw.Top = Y
  fStartYMD = Y
  shpMouseDraw.Width = 100
  shpMouseDraw.Height = 100
  shpMouseDraw.BorderStyle = 3
  shpMouseDraw.Visible = True
  End If
  
If Button = 1 And Shift <> vbShiftMask Then
  'maybe beginning of drag operation
  drag1x = X
  drag1y = Y
  dragbegin = True
  drag2x = drag1x
  drag2y = drag1y
End If

If Button = 2 Then 'left mouse click
  udtMyCoord = GetValues(X, Y)
  Unload frmShowValues
  Load frmShowValues
  frmShowValues.Top = Y 'frmDraw.Top + Y
  frmShowValues.Left = frmDraw.Left + X
  frmShowValues.Show vbModeless
  frmShowValues.CurrentX = 10
  frmShowValues.CurrentY = 10
  sCoord = "(X,Y)=(" & Str$(udtMyCoord.X) & " , " & Str$(udtMyCoord.Y) & ")"
  nLenStr = Len(sCoord)
  frmShowValues.Width = nLenStr * 80
  If (frmShowValues.Left + frmShowValues.Width) > (frmDraw.Width + frmDraw.Left) Then
    frmShowValues.Left = frmDraw.Width + frmDraw.Left - frmShowValues.Width
  End If
  frmShowValues.Print sCoord
End If
  
  
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim udtMyCoord As COORDINATE
Dim sCoord As String 'containing coordinate-info
Dim nLenStr As Integer 'length of string coordinate-info


If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  If (X - fStartXMD) > 0 Then
    shpMouseDraw.Width = Abs(X - fStartXMD)
  Else
    shpMouseDraw.Left = X
    shpMouseDraw.Width = Abs(fStartXMD - X)
  End If
  If (Y - fStartYMD) > 0 Then
    shpMouseDraw.Height = Abs(Y - fStartYMD)
  Else
    shpMouseDraw.Top = Y
    shpMouseDraw.Height = Abs(fStartYMD - Y)
  End If
            
End If

If Button = 1 And dragbegin = True Then
   frmDraw.DrawMode = 7
   frmDraw.DrawStyle = vbDot
   frmDraw.DrawWidth = 1
   frmDraw.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
   frmDraw.Line (drag1x, drag1y)-(X, Y), QBColor(15), B
   drag2x = X
   drag2y = Y
   End If
   


End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Coord1 As COORDINATE, Coord2 As COORDINATE


If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  shpMouseDraw.Visible = False
  If (X - fStartXMD) < 0 Then fStartXMD = fStartXMD - Abs(X - fStartXMD)
  If (Y - fStartYMD) < 0 Then fStartYMD = fStartYMD - Abs(Y - fStartYMD)
  fWidthMD = Abs(X - fStartXMD)
  fHeightMD = Abs(Y - fStartYMD)
  SetZoomValues fStartXMD, fStartYMD, fWidthMD, fHeightMD
  Plot frmDraw, dPlot, udtMyGraphLayout 'plot the zoomed area
  
End If
  
If Button = 1 And Shift <> vbShiftMask And (Abs(drag1x - drag2x) > 50 And Abs(drag1y - drag2y) > 50) Then
   drag2x = X
   drag2y = Y
   
'  'erase final line and refresh graph with new limits
   frmDraw.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
   Coord1 = GetValues(drag1x, drag1y)
   Coord2 = GetValues(drag2x, drag2y)
   YMin0 = frmSetCond.txtValueY0
   YRange0 = frmSetCond.txtValueY1
   XMin0 = frmSetCond.txtValueX0
   XRange0 = frmSetCond.txtValueX1
   Dim temp As Double
   If Coord2.Y < Coord1.Y Then
      temp = Coord2.Y
      Coord2.Y = Coord1.Y
      Coord1.Y = temp
      End If
   frmSetCond.txtValueY0 = Coord2.Y
   frmSetCond.txtValueY1 = Coord1.Y
   If Coord2.X < Coord1.X Then
      temp = Coord2.X
      Coord2.X = Coord1.X
      Coord1.X = temp
      End If
   frmSetCond.txtValueX0 = Coord1.X
   frmSetCond.txtValueX1 = Coord2.X
   Call Form_DblClick
   dragbegin = False
   frmSetCond.txtValueY0 = YMin0
   frmSetCond.txtValueY1 = YRange0
   frmSetCond.txtValueX0 = XMin0
   frmSetCond.txtValueX1 = XRange0
ElseIf Button = 1 And Shift <> vbShiftMask And (Abs(drag1x - drag2x) <= 50 Or Abs(drag1y - drag2y) <= 50) Then
  'erase box
   frmDraw.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set frmDraw = Nothing
   Unload frmShowValues
End Sub

Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

   If frmDraw.WindowState = vbMaximized Then
      Form_DblClick
      ReSized = True
   Else
      If ReSized Then
         Form_DblClick
         End If
      End If
      
   If UBound(dPlot, 3) > 0 Then Form_DblClick

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:
          
End Sub


