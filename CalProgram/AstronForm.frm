VERSION 5.00
Begin VB.Form AstronForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "Input the ITM coordinates"
   ClientHeight    =   5145
   ClientLeft      =   8220
   ClientTop       =   2790
   ClientWidth     =   2985
   Icon            =   "AstronForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   15
      ToolTipText     =   "Negative for Western Hemisphere"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Geo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "Use longitude, latitude"
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ITM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   13
      ToolTipText     =   "Use ITM coordinates"
      Top             =   840
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Cancelbut 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Savebut 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Closebut 
      BackColor       =   &H00FFFF00&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "elevation (meters)"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "ITM Y coordinate"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "ITM X coordinate"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Time Zone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Previously Saved Entries"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Hebrew Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "hgt(m)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "kmy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "kmx:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "AstronForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancelbut_Click()
   Katz = False
   astronplace = False
   astronfm = False
   If Katz = False Then Caldirectories.Astroncheck.Value = vbUnchecked
   Unload Caldirectories
   geo = False
   Call Form_QueryUnload(i%, j%)
   CalMDIform.Visible = True
End Sub

Private Sub Closebut_Click()
   check% = False
   astkmx = Val(Text1.Text)
   astkmy = Val(Text2.Text)
'   If geo = True And (astkmx < -66 Or astkmx > 66 Or Not IsNumeric(astkmx)) Then
'      response = MsgBox("Permissable range of latitudes is from -66 S. to 66 N. ", vbExclamation + vbOKOnly, "Cal Program")
'      Exit Sub
'      End If
   If Not IsNumeric(astkmy) Then
      response = MsgBox("You inputed a string instead of a number!", vbExclamation + vbOKOnly, "Cal Program")
      Exit Sub
      End If
      
   asthgt = Val(Text3.Text)
   If Not IsNumeric(asthgt) Then
      response = MsgBox("Input a height in numeric form!", vbExclamation + vbOKOnly, "Cal Program")
      Exit Sub
      End If
   If geo = False Then
      If astkmx > 80000 Then
         astkmx = astkmx * 0.001
         Text1.Text = astkmx
         check% = True
         End If
      If astkmy > 1000000 Then
         astkmy = (astkmy - 1000000) * 0.001
         Text2.Text = astkmy
         check% = True
      ElseIf astkmy > 10000 And astkmy < 1000000 Then
         astkmy = astkmy * 0.001
         Text2.Text = astkmy
         check% = True
         End If
      avekmxnetz = astkmx 'parameters used for z'manim tables
      avekmynetz = astkmy
      avekmxskiy = astkmx
      avekmyskiy = astkmy
   Else
      geotz! = Val(Text5.Text)
      avekmxnetz = astkmy 'parameters used for z'manim tables
      avekmynetz = astkmx
      avekmxskiy = astkmy
      avekmyskiy = astkmx
      If Abs(geotz! > 12.5) Or Not IsNumeric(geotz!) Then
         response = MsgBox("TZ seems to be incorrect, please check it!", vbExclamation + vbOKOnly, "Cal Program")
         End If
      End If
   avehgtnetz = asthgt
   astname$ = Text4.Text
   If check% = True Then
      response = MsgBox("The coordinates have been converted to a standard format, are they correct?", vbQuestion + vbYesNo, "Cal Program")
      If response = vbNo Then
         Exit Sub
         End If
      End If
   'Caldirectories.Text1.Text = drivcities$ + "ast\" + LTrim$(astname$)
   astronplace = True
   astronfm = True
   'create dummy batch files for that place
   response = MsgBox("Do you wan't to add the typical 1.8 m for observer height to the place height?", vbYesNo + vbQuestion, "Cal Program")
   If response = vbYes Then
      obshgt = 1.8
   Else
      obshgt = 0#
      End If
   filnum% = FreeFile
   Open drivcities$ + "ast\netz\astr.bat" For Output As #filnum%
   Write #filnum%, drivfordtm$ + "netz\astronom.pr1", Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt
   Print #filnum%, "version"; ","; "1"; ","; "0"; ","; "0"
   Close #filnum%
   Open drivcities$ + "ast\skiy\astr.bat" For Output As #filnum%
   Write #filnum%, drivfordtm$ + "skiy\astronom.pr1", Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt
   Print #filnum%, "version"; ","; "1"; ","; "0"; ","; "0"
   Close #filnum%
   'now make dummy profile files to place in c:\cities\ast netz and skiy subdirectories
   Open drivcities$ + "ast\netz\astronom.pr1" For Output As #filnum%
   'Write #filnum%, "FILENAME, KMX, KMY, HGT: ", "c:\prof\astronom.fnz", Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt
   Write #filnum%, "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR"
   Write #filnum%, Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt, 0, 0, 0, 0, 0

   'Print #filnum%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
   For i% = 1 To 601
      xentry = -30 + (i% - 1) * 0.1
      Write #filnum%, CInt(xentry * 10) * 0.1, 0, 0, 0, 0, 0
'      If xentry <= -10 Then
'         Print #filnum%, Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry > -10 And xentry < 0 Then
'         Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry >= 0 And xentry < 10 Then
'         Print #filnum%, "  "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry >= 10 Then
'         Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'         End If
   Next i%
   Close #filnum%
   Open drivcities$ + "ast\skiy\astronom.pr1" For Output As #filnum%
   'Write #filnum%, "FILENAME, KMX, KMY, HGT: ", "c:\prof\astronom.fsk", Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt
   'Print #filnum%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
   Write #filnum%, "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR"
   Write #filnum%, Val(Text1.Text), Val(Text2.Text), Val(Text3.Text) + obshgt, 0, 0, 0, 0, 0
   
   For i% = 1 To 601
       Write #filnum%, CInt(xentry * 10) * 0.1, 0, 0, 0, 0, 0
'      If xentry <= -10 Then
'         Print #filnum%, Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry > -10 And xentry < 0 Then
'         Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry >= 0 And xentry < 10 Then
'         Print #filnum%, "  "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'      ElseIf xentry >= 10 Then
'         Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
'         End If
   Next i%
   Close #filnum%
      
   startedscan = False
   If Katz = True Then GoTo 100
   Caldirectories.Text1.Text = drivcities$ + "ast\" + LTrim$(astname$)
   Caldirectories.OKbutton.Enabled = False
   Caldirectories.ExitButton.Enabled = False
   Caldirectories.Label1.Enabled = False
   Caldirectories.Drive1.Enabled = False
   Caldirectories.Dir1.Enabled = False
   'List1.Enabled = False
   Caldirectories.Text1.Enabled = False
   Caldirectories.Astroncheck.Value = vbUnchecked
100:
   SunriseSunset.Timer1.Enabled = False
   SunriseSunset.Show
   SunriseSunset.Enabled = True
   SunriseSunset.ProgressBar1.Enabled = True
   SunriseSunset.ProgressBar1.Visible = True
   'SunriseSunset.Label3.Visible = False
   currentdir = Trim$(Text1.Text)
   SunriseSunset.Visible = True
   SunriseSunset.Check1.Enabled = True
   SunriseSunset.Check2.Enabled = True
   SunriseSunset.Check3.Enabled = True
   SunriseSunset.Check4.Enabled = True
   SunriseSunset.Check5.Enabled = True
   SunriseSunset.Combo1.Enabled = True
   SunriseSunset.Label2.Enabled = True
   SunriseSunset.Cancelbut.Enabled = True
   If SunriseSunset.Label1.Caption <> sEmpty Then SunriseSunset.Label1.Caption = captmp$
   SunriseSunset.Label1.Enabled = True
   SunriseSunset.OKbut0.Enabled = True
   SunriseSunset.ProgressBar1.Visible = False
   SunriseSunset.Option1.Enabled = True
   SunriseSunset.Option2.Enabled = True
   If suntop% <> 0 Then SunriseSunset.Top = suntop%
   SunriseSunset.Check1.Enabled = False
   SunriseSunset.Check2.Enabled = False
   SunriseSunset.Check3.Enabled = False
   SunriseSunset.Check4.Enabled = True
   SunriseSunset.Check5.Enabled = True
   SunriseSunset.Check6.Enabled = True
   SunriseSunset.Check7.Enabled = True
   SunriseSunset.Check4.Value = vbChecked
   Option1b = True
   '*******Katz changes*********
   If Katz = True Then
      SunriseSunset.Option2.Value = True
      SunriseSunset.Check6.Value = vbChecked
      SunriseSunset.Check7.Value = vbChecked
      End If
   'ret = SetWindowPos(SunriseSunset.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   'Caldirectories.Astroncheck.Value = vbUnchecked
   Call Form_QueryUnload(i%, j%)
End Sub


Private Sub Combo1_Click()
   Combo1.Text = astrplaces$(Combo1.ListIndex + 1)
   Text1.Text = astcoord(1, Combo1.ListIndex + 1)
   Text2.Text = astcoord(2, Combo1.ListIndex + 1)
   Text3.Text = astcoord(3, Combo1.ListIndex + 1)
   Text4.Text = astrplaces$(Combo1.ListIndex + 1)
   Text5.Text = astcoord(4, Combo1.ListIndex + 1)
   If astcoord(5, Combo1.ListIndex + 1) = 0 Then
      Option2.Value = True
   ElseIf astcoord(5, Combo1.ListIndex + 1) = 1 Then
      Option1.Value = True
      End If
   If Katz = True Then
      katztotal% = AstronForm.Combo1.ListIndex
      End If
End Sub

Private Sub Form_Load()
   geo = False
   
   If Katz = True Then
      drivjk$ = "c:/jk/"
      drivfordtm$ = "c:/fordtm/"
      openfil$ = "Katzplaces.sav"
   Else
      openfil$ = "astronplaces.sav"
      End If
   
   myfile = Dir(drivjk$ + openfil$)
   If myfile = sEmpty Then
      Exit Sub
   Else
      Combo1.Enabled = True
      Label5.Enabled = True
      filnum% = FreeFile
      If Dir(drivjk$ + openfil$) = sEmpty Then
         Call MsgBox("The file: " & drivjk$ & openfil$ & " doesn't exist." _
                     & vbCrLf & "" _
                     & vbCrLf & "Restore the file to the directory: " + drivjk$ _
                     , vbExclamation, "astronomical places")
         Exit Sub
         End If
      Open drivjk$ + openfil$ For Input As #filnum%
      i% = 0
      numAstPlaces% = 0
      Do Until EOF(filnum%)
         i% = i% + 1
         Input #filnum%, astrplaces$(i%), astcoord(1, i%), astcoord(2, i%), astcoord(3, i%), astcoord(4, i%), astcoord(5, i%)
         Combo1.AddItem astrplaces$(i%)
         numAstPlaces% = numAstPlaces% + 1
      Loop
      If Katz = True And katztotal% <= AstronForm.Combo1.ListCount - 1 Then
         AstronForm.Combo1.ListIndex = katztotal%
         AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
         If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
            AstronForm.Option2.Value = True
         ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
            AstronForm.Option1.Value = True
            End If
      ElseIf Katz = True And katztotal% > AstronForm.Combo1.ListCount - 1 Then
         katztotal% = 0
         AstronForm.Combo1.ListIndex = 0
         AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
         If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
            AstronForm.Option2.Value = True
         ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
            AstronForm.Option1.Value = True
            End If
      ElseIf Katz = False Then
         AstronForm.Combo1.ListIndex = AstronForm.Combo1.ListCount - 1
         AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
         AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
         If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
            AstronForm.Option2.Value = True
         ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
            AstronForm.Option1.Value = True
            End If
         End If
      Close #filnum%
      End If
   Closebut.Visible = True
   If astcoord(5, i%) = 1 Then
      geotz! = 2
      Option1.Value = True
   ElseIf astcoord(5, i%) = 0 Then
      geotz! = astcoord(4, i%)
      Option2.Value = True
      End If
  Text5.Text = geotz!
End Sub

Private Sub Option1_Click()
   Label1.Caption = "kmx:"
   Label2.Caption = "kmy:"
   geotz! = 2
   Text1.ToolTipText = "ITM X coordinate"
   Text2.ToolTipText = "ITM Y coordinate"
   geo = False
   AstronForm.Caption = "Input the ITM coordinates"
   Label6.Visible = False
   Text5.Visible = False
End Sub

Private Sub Option2_Click()
   Label2.Caption = "long:"
   Label1.Caption = "lat:"
   Text2.ToolTipText = "longitude (positive for West longitude)"
   Text1.ToolTipText = "latitude (positive for North latitude)"
   geo = True
   AstronForm.Caption = "Input the Geo coordinates"
   Label6.Visible = True
   Text5.Visible = True
End Sub

Private Sub savebut_Click()
    If Katz = True Then
       openfil$ = "Katzplaces.sav"
    Else
       openfil$ = "astronplaces.sav"
       End If

   myfile = Dir(drivjk$ + openfil$)
   filnum% = FreeFile
   If myfile = sEmpty Then
      Open drivjk$ + openfil$ For Output As #filnum%
   Else
      Open drivjk$ + openfil$ For Append As #filnum%
      End If
   If geo = False Then
      Write #filnum%, Text4.Text, Val(Text1.Text), Val(Text2.Text), Val(Text3.Text), 2, 1
   ElseIf geo = True Then
      Write #filnum%, Text4.Text, Val(Text1.Text), Val(Text2.Text), Val(Text3.Text), Val(Text5.Text), 0
      End If
   Close #filnum%
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set AstronForm = Nothing
End Sub
