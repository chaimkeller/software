VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form mapbatlistfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "contents of bat file"
   ClientHeight    =   6210
   ClientLeft      =   6690
   ClientTop       =   1830
   ClientWidth     =   4680
   Icon            =   "mapbatlistfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdPlotAll 
      Caption         =   "&Plot All"
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2855
      Left            =   60
      ScaleHeight     =   2790
      ScaleWidth      =   4515
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   2355
      Left            =   0
      Picture         =   "mapbatlistfm.frx":0442
      ScaleHeight     =   2295
      ScaleWidth      =   4620
      TabIndex        =   17
      Top             =   3420
      Width           =   4680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change the Center Coordinates"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "to change center  coordinates of world place....."
      Enabled         =   0   'False
      Height          =   2355
      Left            =   60
      TabIndex        =   4
      Top             =   3420
      Width           =   4575
      Begin VB.CommandButton Command3 
         Height          =   555
         Left            =   1680
         Picture         =   "mapbatlistfm.frx":45BA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Use Maps & More Center Coordinates"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   14
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Record changes"
         Height          =   315
         Left            =   1380
         TabIndex        =   13
         Top             =   1980
         Width           =   1995
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3540
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2100
         TabIndex        =   10
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3540
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Save File Name:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "New coord."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Old coord."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   5775
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8114
            MinWidth        =   8114
            Text            =   "Click on the desired place"
            TextSave        =   "Click on the desired place"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Top             =   180
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "bat file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "mapbatlistfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PlotSearchPoints2 As Boolean
Private Sub cmdPlotAll_Click()
   'plot all the search results on the map

   On Error GoTo cmdPlotAll_Click_Error
   
   ret = SetWindowPos(mapbatlistfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
   
   If PlotSearchPoints2 Then
      'button pressed twice, so erase plot points
      blitpictures
      PlotSearchPoints2 = False
      Exit Sub
      End If

    For j& = 1 To List1.ListCount
        doclin$ = List1.List(j& - 1)
        'find the coordinates
        pos1% = InStr(1, doclin$, ",")
        If pos1% = 0 Then
           'response = MsgBox("Not a coordinate!", vbOKOnly + vbCritical, "Maps & More")
           GoTo cpa500
           End If
        pos2% = InStr(pos1% + 1, doclin$, ",")
        If pos2% = 0 Then
           'response = MsgBox("Not a coordiate!", vbOKOnly + vbCritical, "Maps & More")
           GoTo cpa500
           End If
        xcoor = Val(Mid$(doclin$, pos1% + 1, pos2% - pos1% - 1))
        pos3% = InStr(pos2% + 1, doclin$, ",")
        If pos3% = 0 Then
           'response = MsgBox("Not a coordinate!", vbOKOnly + vbCritical, "Maps & More")
           GoTo cpa500
           End If
        ycoor = Val(Mid$(doclin$, pos2% + 1, pos3% - pos2% - 1))
        If world = True Then
           Call ScreenToGeo(X, Y, -ycoor, xcoor, 2, ier%)
        Else
           If xcoor And ycoor < 1000 Then 'convert format
              xcoor = xcoor * 1000
              ycoor = ycoor * 1000 + 1000000
              End If
           Call ScreenToGeo(X, Y, xcoor, ycoor, 2, ier%)
           End If
      
       'plot the points
       mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
       mapPictureform.mapPicture.Circle (X, Y), 20, 255 '20 * mag, 255
       mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
cpa500:
    Next j&
    
    PlotSearchPoints2 = True
    
   On Error GoTo 0
   Exit Sub

cmdPlotAll_Click_Error:
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPlotAll_Click of Form mapbatlistfm"

End Sub

Private Sub Command1_Click()
   On Error GoTo errhandler
   If Text5.Text = sEmpty Or Text6.Text = sEmpty Then
      response = MsgBox("You must provide new values for the latitude and longitude!", vbCritical + vbOKOnly, "Maps & More")
      Exit Sub
      End If
   If resetorigin = True Then 'find city in skyworld.sav and
                              'determine origin shift, and then
                              'record new values in skyworld.sav
     If Text8.Text = sEmpty Then
        response = MsgBox("You must enter the name of the city!", vbExclamation + vbOKOnly, "Maps & More")
        Exit Sub
        End If
     If Text2.Text = sEmpty Or Text3.Text = sEmpty Then
        response = MsgBox("You must enter a reference coordinates!", vbExclamation + vbOKOnly, "Maps & More")
        Exit Sub
        End If
    
    filsav% = FreeFile
    Open drivjk_c$ + "skyworld.sav" For Input As #filsav%
    found% = 0
    Do Until EOF(filsav%)
       Input #filsav%, savcity$, savlog, savlat, savhgt
       If InStr(1, savcity$, Text8.Text) <> 0 Then
          response = MsgBox("In skyworld.sav found city: " + savcity$ + " Is this the right city?", vbQuestion + vbYesNoCancel, "Maps & More")
          If response = vbYes Then
             found% = 1
             Exit Do
             End If
          End If
    Loop
    Close #filsav%
    If found% = 0 Then
       response = MsgBox("City not found in skyworld.sav! Check your spelling.", vbCritical + vbOKOnly, "Maps & More")
       Exit Sub
       End If
        
    'now process changes
    Text2.Text = Maps.Text5.Text
    Text3.Text = Maps.Text6.Text
    Text4.Text = Maps.Text7.Text
    diflogtmp = Val(Text5.Text) - Val(Text2.Text)
    diflattmp = Val(Text6.Text) - Val(Text3.Text)
    If Abs(diflogtmp) > 0.5 Or Abs(diflattmp) > 0.5 Then
       response = MsgBox("The differences are very big. Are you sure?", vbExclamation + vbYesNoCancel, "Maps & More")
       If response <> vbYes Then
          Exit Sub
          End If
       End If
    savlog = savlog + diflogtmp
    savlat = savlat + diflattmp
  
    response = MsgBox("Append the city's new coordinates to skyworld.sav?", vbQuestion + vbYesNo, "Maps & More")
    If response = vbYes Then
      filsav% = FreeFile
      Open drivjk_c$ + "skyworld.sav" For Append As #filsav%
      Call worldheights(-savlog, savlat, newhgt)
      If newhgt = -9999 Then newhgt = 0
      Text7.Text = newhgt
      Write #filsav%, savcity$, CSng(savlog), CSng(savlat), CSng(newhgt)
      Close #filsav%
      End If
    resetorigin = False
    Exit Sub
    End If
    
   'warn user of reprecations, and ask if wan't to backup first
   response = MsgBox("Warning, you are about to make significant changes in all the coordinates of your place.  Do you wan't to copy the old files to backup files first?", vbExclamation + vbYesNoCancel, "Maps & More")
   If response = vbYes Then
     'make backups
     FileCopy LTrim$(RTrim$(Text8.Text)) + ".sav", LTrim$(RTrim$(Text8.Text)) + ".bak"
   ElseIf response = vbCancel Then
     Exit Sub
     End If
     
  'now copy sav file to tmp file, and rewrite sav file
  Screen.MousePointer = vbHourglass
  FileCopy LTrim$(RTrim$(Text8.Text)) + ".sav", LTrim$(RTrim$(Text8.Text)) + ".tmp"
  
  'Kill Text8.Text + ".sav"
  diflattmp = Val(Text2.Text) - Val(Text5.Text)
  diflogtmp = Val(Text3.Text) - Val(Text6.Text)
  filtmp% = FreeFile
  Open LTrim$(RTrim$(Text8.Text)) + ".tmp" For Input As #filtmp%
mb200:
  filsav% = FreeFile
  Open LTrim$(RTrim$(Text8.Text)) + ".sav" For Output As #filsav%
  found% = 0
  Do Until EOF(filtmp%)
     Input #filtmp%, doccity$, lattmp, logtmp, loghgt
     newlat = lattmp - diflattmp
     newlog = logtmp - diflogtmp
     'now determine the height at the new place if not inputed
     Call worldheights(-newlog, newlat, newhgt)
     If newhgt = -9999 Then newhgt = 0
     Write #filsav%, doccity$, CSng(newlat), CSng(newlog), newhgt
     If InStr(1, Text8.Text, doccity$) <> 0 Then
        Text7.Text = CStr(newhgt)
        found% = 1
        savcity$ = doccity$
        savlog = CSng(newlog)
        savlat = CSng(newlat)
        savhgt = newhgt
        End If
  Loop
  Close #filtmp%
  Close #filsav%
  Kill Text8.Text + ".tmp"
  If found% = 0 Then
     Screen.MousePointer = vbDefault
     response = MsgBox("Couldn't find a city that is a subset of the string " + Text8.Text + " in skyworld.sav. Check your spelling!  Do you wan't to try again?", vbCritical + vbYesNoCancel, "Maps & More")
     If response = vbYes Then
        GoTo mb200
     Else
        GoTo mb500
        End If
     End If
  'now add this new coordinates to skyworld.sav if desired
  response = MsgBox("Append new center coordinates to skyworld.sav?", vbQuestion + vbYesNo, "Maps & More")
  If response = vbYes Then
    filsav% = FreeFile
    Open drivjk_c$ + "skyworld.sav" For Append As #filsav%
    Write #filsav%, savcity$, -savlog, savlat, savhgt
    Close #filsav%
    End If
mb500:
  ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
  Screen.MousePointer = vbDefault
  Exit Sub
  
errhandler:
   Screen.MousePointer = vbDefault
   response = MsgBox("mapbatlistfm encountered error number: " + Str(Err.Number) + ". The error message is: " + Err.Description + " Sorry!", vbCritical + vbOKOnly, "Maps & More")
   ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
   'close all files and exit
   Close
   'kill the sav file and copy the tmp file back to it
   If Err.Number = 54 Then
      'attempt to restore the sav file (it has been erased)
      FileCopy Text8.Text + ".tmp", Text8.Text + ".sav"
      Call form_queryunload(0, 0)
      End If
End Sub

Private Sub Command2_Click()
   'determine root name of place
   On Error GoTo errhandler
   If resetorigin = True Then
      mapbatlistfm.Picture1.Visible = False
      mapbatlistfm.Frame1.Enabled = True
      Exit Sub
      End If
   pos% = InStr(1, Text1.Text, "\netz")
   If pos% = 0 Then Exit Sub
   Text8.Text = Mid$(Text1.Text, 1, pos% - 1)
   'open sav file and read center coordinates
   If Dir(Text8.Text + ".sav") = sEmpty Then
      response = MsgBox("Sav file: " + Text8.Text + ".sav not found!", vbCritical + vbOKOnly, "Maps & More")
      Exit Sub
      End If
   mapbatlistfm.Picture1.Visible = False
   mapbatlistfm.Frame1.Enabled = True
   mapbatlistfm.StatusBar1.Panels(1) = "Enter new center coordinates coordinates for this place"
   filsav% = FreeFile
   Open Text8.Text + ".sav" For Input As #filsav%
   Do Until EOF(filsav%)
      Input #filsav%, doclin$, latsav, lonsav, hgtsav
      pos2% = InStr(1, UCase(Text8.Text), UCase(doclin$))
      If pos2% <> 0 Then
         Text2.Text = latsav
         Text3.Text = lonsav
         Text4.Text = hgtsav
         Exit Do
         End If
   Loop
   Close #filsav%
   Exit Sub

errhandler:
   response = MsgBox("mapbatlistfm encountered error number: " + CStr(Err.Number) + " Error message is: " + Err.Description, vbCritical + vbOKOnly, "Maps & More")
   
End Sub

Private Sub Command3_Click()
   If resetorigin = True Then
      Text2.Text = Maps.Text5.Text
      Text3.Text = Maps.Text6.Text
      Text4.Text = Maps.Text7.Text
   Else
      Text2.Text = Maps.Text6.Text
      Text3.Text = -Maps.Text5.Text
      Text4.Text = Maps.Text7.Text
      End If
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   resetorigin = False
   Unload Me
   Set mapbatlistfm = Nothing
   If MapOn Then ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub List1_Click()
   doclin$ = List1.List(List1.ListIndex)
   'find the coordinates
   pos1% = InStr(1, doclin$, ",")
   If pos1% = 0 Then
      response = MsgBox("Not a coordinate!", vbOKOnly + vbCritical, "Maps & More")
      ret = BringWindowToTop(mapbatlistfm.hWnd)
      'ret = SetWindowPos(mapbatlistfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      'Call form_queryunload(0, 0)
      Exit Sub
      End If
   pos2% = InStr(pos1% + 1, doclin$, ",")
   If pos2% = 0 Then
      response = MsgBox("Not a coordiate!", vbOKOnly + vbCritical, "Maps & More")
      ret = BringWindowToTop(mapbatlistfm.hWnd)
      'ret = SetWindowPos(mapbatlistfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      'Call form_queryunload(0, 0)
      Exit Sub
      End If
   xcoor = Mid$(doclin$, pos1% + 1, pos2% - pos1% - 1)
   pos3% = InStr(pos2% + 1, doclin$, ",")
   If pos3% = 0 Then
      response = MsgBox("Not a coordinate!", vbOKOnly + vbCritical, "Maps & More")
      ret = BringWindowToTop(mapbatlistfm.hWnd)
      'ret = SetWindowPos(mapbatlistfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      'Call form_queryunload(0, 0)
      Exit Sub
      End If
   ycoor = Mid$(doclin$, pos2% + 1, pos3% - pos2% - 1)
   If world = True Then
      Maps.Text6 = xcoor
      Maps.Text5 = -ycoor
      If tblbuttons(3) = 1 Then
        lResult = FindWindow(vbNullString, "Overview")
        If lResult <> 0 Then
            TdxhWnd = 0
            bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
            iposit% = InStr(Tdxname, "-  ")
            lat3d = Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60
            lon3d = Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60
            OverhWnd = FindWindow(vbNullString, "Overview")
            ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            'Call BringWindowToTop(OverhWnd)
            dx1 = -1000 '-30 '30
            dy1 = -1000 '-240 '60
            Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
            waitime = Timer + 0.01
            Do Until Timer > waitime
               DoEvents
            Loop
            lon3dnew = Val(ycoor)
            lat3dnew = Val(xcoor)
            dx1 = -(lon3dnew - lon3d) * 516.6 + 96
            dy1 = -(lat3dnew - lat3d) * 516.6 + 156
            Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
            waitime = Timer + 0.01
            Do Until Timer > waitime
               DoEvents
            Loop
            Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
            Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
            ret = SetWindowPos(mapsearchfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            waitime = Timer + 2
            Do Until Timer > waitime
               DoEvents
            Loop
            Call BringWindowToTop(mapbatlistfm.hWnd)
            Exit Sub
            End If
         End If
   Else
      Maps.Text5 = xcoor
      Maps.Text6 = ycoor
      End If
   Call goto_click
End Sub
