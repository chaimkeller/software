VERSION 5.00
Begin VB.Form netzskiyfm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   5595
   ClientLeft      =   2940
   ClientTop       =   1995
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboGoogle 
      Height          =   315
      ItemData        =   "netzskiyfm.frx":0000
      Left            =   6600
      List            =   "netzskiyfm.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Choose map layer"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton NetzSkiyOkbut0 
      BackColor       =   &H0080FF80&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Picture         =   "netzskiyfm.frx":002A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton NetzskiyCancelbut 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   4680
      Picture         =   "netzskiyfm.frx":046C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ListBox Netzskiylist 
      Height          =   4110
      Index           =   2
      ItemData        =   "netzskiyfm.frx":08AE
      Left            =   120
      List            =   "netzskiyfm.frx":08B5
      Style           =   1  'Checkbox
      TabIndex        =   1
      ToolTipText     =   "Double Click to display on a google map"
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "netzskiyfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAll_Click()
   For i% = 6 To nn4%
      Netzskiylist(2).Selected(i% - 6) = True
   Next i%
End Sub

Private Sub cmdClear_Click()
   For i% = 6 To nn4%
      Netzskiylist(2).Selected(i% - 6) = False
   Next i%
End Sub

Private Sub NetzskiyCancelbut_Click(Index As Integer)
   'version: 04/08/2003
  
'  exit from rountine and go back to main menu
   Screen.MousePointer = vbDefault
   netzskiyok = False
   Unload netzskiyfm
   Set netzskiyfm = Nothing
   SunriseSunset.Visible = False
   'If calnode.Visible = True Then
   If eroscityflag = True Then
      Caldirectories.Visible = False
      Exit Sub
      End If
   Caldirectories.Label1.Enabled = True
   Caldirectories.Drive1.Enabled = True
   Caldirectories.Dir1.Enabled = True
   'Caldirectories.List1.Enabled = True
   Caldirectories.Text1.Enabled = True
   Caldirectories.OKbutton.Enabled = True
   Caldirectories.ExitButton.Enabled = True
   'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   'myfile = Dir(drivjk$+":\jk\netzskiy.tm3")
   'If myfile <> sEmpty Then
   '   Kill drivjk$+":\jk\netzskiy.tm3"
   '   End If
   'myfile = Dir(drivjk$+":\jk\netzskiy.tm4")
   'If myfile <> sEmpty Then
   '   Kill drivjk$+":\jk\netzskiy.tm4"
   '   End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Netzskiylist_DblClick
' Author    : chaim
' Date      : 10/31/2023
' Purpose   : identify position of vantage point on a google map
'---------------------------------------------------------------------------------------
'
Private Sub Netzskiylist_DblClick(Index As Integer)

   On Error GoTo Netzskiylist_DblClick_Error

Select Case MsgBox("Do you want to locate this place on a Google Map?" _
                   & vbCrLf & "" _
                   & vbCrLf & "(Hint: you can also change the type of map layer in the combo box below)" _
                   , vbYesNo Or vbQuestion Or vbDefaultButton2, "Location on map")

    Case vbYes

    Case vbNo
       Exit Sub

End Select

' The basic map URL without the address information.
Const URL_BASE As String = "http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=@ADDR@&ie=UTF8&t=@TYPE@"

Dim addr As String
Dim url As String
Dim DataLine() As String
Dim latitude As Double
Dim longitude As Double
Dim PntSelected As Boolean
Dim lgAT As Double
Dim ltAT As Double

    '39.358008,-76.688316
    waitime = Timer
    Do Until Timer > waitime + 0.1
       DoEvents
    Loop
    With netzskiyfm.Netzskiylist(2)
        PntSelected = Not .Selected(.ListIndex) 'store checked status
        DataLine = Split(.List(.ListIndex), ",")
    
        longitude = Val(DataLine(1))
        latitude = Val(DataLine(2))
        
        'check for ITM coordinates
        If longitude > 80 And latitude > 100 Then
           'convert from ITM to geo
            Call casgeo(longitude, latitude, lgAT, ltAT)
    '           lgAT = -lgAT 'this is convention for WorldClim
            longitude = lgAT
            latitude = ltAT
        Else
            Call MsgBox("These coordinates don't seem be inside Israel:" _
                        & vbCrLf & "These are the coordinates read:" _
                        & vbCrLf & "ITMx =" & Str$(longitude) _
                        & vbCrLf & "ITMy =" & Str$(latitude) _
                        , vbInformation, "Coordinates error")
            Exit Sub
            End If
           
        ' A very simple URL encoding.
        '**original code modified to contain coordinates instead of address - 011122****************
        addr = Str$(latitude) & "," & Str$(-longitude)
        '**********************************************************
        addr = Replace$(addr, " ", "+")
        addr = Replace$(addr, ",", "%2c")
    
        ' Insert the encoded address into the base URL.
        url = Replace$(URL_BASE, "@ADDR@", addr)
    
        ' Insert the proper type.
        Select Case cboGoogle.Text
            Case "Map"
                url = Replace$(url, "@TYPE@", "m")
            Case "Satellite"
                url = Replace$(url, "@TYPE@", "h")
            Case "Terrain"
                url = Replace$(url, "@TYPE@", "p")
        End Select
    
        ' "Execute" the URL to make the default browser display it.
        ShellExecute ByVal 0&, "open", url, _
            vbNullString, vbNullString, SW_SHOWMAXIMIZED
            
        'restore the check mark if checked
        If PntSelected Then
          .Selected(.ListIndex) = True
       Else
          .Selected(.ListIndex) = False
          End If
    
    End With

   On Error GoTo 0
   Exit Sub

Netzskiylist_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Netzskiylist_DblClick of Form netzskiyfm"
End Sub

Private Sub NetzskiyOKbut0_Click()
   Screen.MousePointer = vbHourglass
   If automatic = True Then
      waittime = Timer + 1#
      Do While waittime > Timer
         DoEvents
      Loop
      End If
   numchecked% = 0
   For i% = 6 To nn4%
      If Netzskiylist(2).Selected(i% - 6) = True Then
         nchecked%(i% - 5) = 1
         numchecked% = numchecked% + 1
      Else
         nchecked%(i% - 5) = 0
         End If
   Next i%
   netzskiyok = True
   Unload netzskiyfm
   Set netzskiyfm = Nothing
End Sub


