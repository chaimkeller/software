VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mapPLACfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Stored Places and their Israel Grid Coordinates"
   ClientHeight    =   6345
   ClientLeft      =   6495
   ClientTop       =   1935
   ClientWidth     =   5535
   Icon            =   "mapPLACfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton placCLOSEbut 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   4935
   End
   Begin MSFlexGridLib.MSFlexGrid sky1 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   5000
      Cols            =   5
      BackColor       =   12648447
      ForeColor       =   32768
      ForeColorFixed  =   8388608
      GridColor       =   16576
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^Entry |<Place Name                     |^ITMx      |^ITMy      |^height (m)"
   End
End
Attribute VB_Name = "mapPLACfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   
   If world = True Then
      mapPLACfm.Caption = "List of World-wide Stored Places and their Coordinates"
      mapPLACfm.sky1.FormatString = "^Entry |<Place Name                     |^longitude |^latitude  |^height (m)"
      filsav1$ = drivjk_c$ + "skyworld.sav"
   Else
      filsav1$ = drivjk_c$ + "skylight.sav"
      End If
   filsav% = FreeFile
   myfile = Dir(filsav1$)
   If myfile <> sEmpty Then
      'determine number of columns
      Open filsav1$ For Input As #filsav%
      nplac% = 0
      On Error GoTo formerrhandler
      Do Until EOF(filsav%)
         Line Input #filsav%, doclin$
         nplac% = nplac% + 1
f50:  Loop
      Close #filsav%
      sky1.Rows = nplac% + 1
      'If sky1.Rows < 23 Then sky1.Rows = 23
      Open filsav1$ For Input As #filsav%
      For i% = 1 To nplac%
         Input #filsav%, placnam$, itmx, itmy, itmhgt
         sky1.TextArray(skyp(i%, 0)) = i%
         sky1.TextArray(skyp(i%, 1)) = placnam$
         sky1.TextArray(skyp(i%, 2)) = itmx
         sky1.TextArray(skyp(i%, 3)) = itmy
         sky1.TextArray(skyp(i%, 4)) = itmhgt
      Next i%
      Close #filsav%
      Call dosort
      End If
   sky1.row = 1 'highlight only the first row
   Screen.MousePointer = vbDefault
   Exit Sub
formerrhandler:
   GoTo f50
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lResult As Long
    Unload mapPLACfm
    Set mapPLACfm = Nothing
    lplac% = 0
    tblbuttons%(9) = 0
    Maps.Toolbar1.Buttons(9).value = tbrUnpressed
'    ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    BringWindowToTop (mapPictureform.hWnd)
    If world = False Then
      lResult = FindWindow(vbNullString, terranam$)
      If lResult > 0 Then
'         ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (lResult)
         End If
    Else
      lResult = FindWindow(vbNullString, "3D Viewer")
      If lResult > 0 Then
'         ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (lResult)
         End If
      End If
End Sub

Function skyp(row As Integer, col As Integer) As Long
     skyp = row * sky1.Cols + col
End Function
Private Sub placCLOSEbut_Click()
   Call Form_QueryUnload(i%, j%)
   'Unload mapPLACfm
   'Set mapPLACfm = Nothing
   'mappictureform.mappicture.SetFocus
End Sub
Private Sub sky1_DblClick()
  Dim lResult As Long
  If sky1.MouseRow = 0 Then Exit Sub
  nplachos% = sky1.MouseRow
  placdblclk = True
  If world = False Then
     Maps.Label5.Caption = "ITMx"
     Maps.Label6.Caption = "ITMy"
     kmxc = sky1.TextArray(skyp(sky1.MouseRow, 2))
     kmyc = sky1.TextArray(skyp(sky1.MouseRow, 3))
    lResult = FindWindow(vbNullString, terranam$)
    If lResult > 0 Then
        dx1 = 500 'position cursor off screen above dbl click
        dy1 = -500
        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
        End If
     End If
  Maps.Text7.Text = sky1.TextArray(skyp(sky1.MouseRow, 4))
  If world = True Then
     l2 = sky1.TextArray(skyp(sky1.MouseRow, 2))
     l1 = sky1.TextArray(skyp(sky1.MouseRow, 3))
     Call Form_QueryUnload(i%, j%)
     Maps.Text6.Text = l1
     Maps.Text5.Text = l2
     Call goto_click
     Exit Sub
     End If
  Call Form_QueryUnload(i%, j%)
  Maps.Text6.Text = kmyc
  Maps.Text5.Text = kmxc
'  lResult = FindWindow(vbNullString, terranam$)
'  If lResult > 0 Then
'      dx1 = 500 'position cursor off screen above dbl click
'      dy1 = -500
'      Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0) 'move mouse to Location item
'      End If
  Call goto_click
End Sub
Function fgi(r As Integer, c As Integer) As Integer
   fgi = c + sky1.Rows * r
 End Function
Private Sub dosort()
   sky1.row = 1
   sky1.RowSel = sky1.Rows - 1
   sky1.Sort = 1 'generic ascending
End Sub
