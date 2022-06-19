VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form calnearsearchfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for contributing sites"
   ClientHeight    =   6570
   ClientLeft      =   4125
   ClientTop       =   1635
   ClientWidth     =   5325
   Icon            =   "calnearsearchfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   5325
   Begin VB.ComboBox cboGoogle 
      Height          =   315
      ItemData        =   "calnearsearchfm.frx":0442
      Left            =   3600
      List            =   "calnearsearchfm.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Choose map layer"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ClearAll 
      Caption         =   "Clear &All"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      ToolTipText     =   "Uncheck all results"
      Top             =   5940
      Width           =   1095
   End
   Begin VB.CommandButton cmd_SelectAll 
      Caption         =   "&Select All"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Check all results"
      Top             =   5680
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   6300
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search results"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   5115
      Begin VB.ListBox List1 
         BackColor       =   &H00C0C0C0&
         Height          =   2985
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         ToolTipText     =   "double click to locate search result on map"
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Accept results and proceed to calculation"
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   14
         Text            =   "sunrise"
         ToolTipText     =   "Choose horizon to search"
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "Press to search"
         Top             =   1740
         Width           =   1755
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   2460
         TabIndex        =   8
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text3"
         BuddyDispid     =   196619
         OrigLeft        =   2040
         OrigTop         =   1440
         OrigRight       =   2280
         OrigBottom      =   1815
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1860
         TabIndex        =   7
         Text            =   "7"
         ToolTipText     =   "Search radius (km)"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         TabIndex        =   4
         Text            =   "longitude +/- degrees for E/W longitude."
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         TabIndex        =   3
         Text            =   "latitude in degrees of center of search"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "search radius in kilometers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   420
         TabIndex        =   9
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "latitude of center coordinates:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   420
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "longitude of center coord.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   420
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "calnearsearchfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nn%, sun$, lat As Double, lon As Double, dist As Double
Private xdegkm As Double, ydegkm As Double ', pi As Double

Private Sub cmd_ClearAll_Click()
   For i% = 1 To List1.ListCount
      List1.Selected(i% - 1) = False
   Next i%
End Sub

Private Sub cmd_SelectAll_Click()
   For i% = 1 To List1.ListCount
      List1.Selected(i% - 1) = True
   Next i%
End Sub

Private Sub Command1_Click()
   Dim batdoc$(1, 1000), batdocnum%(1)
   Dim nextskiy As Boolean
   
   On Error GoTo searcherrhand
   foundvantage = False
   'read the checked files, and copy then to
   'd:\cities\eros\visual_tmp\netz
   'also generate .bat file and place it in that directory
   startedscan = False
   currentdir = drivcities$ & "eros\visual_tmp"
   'kill everything in visual_tmp directory
   If internet = False Then
      Close
      If Dir(currentdir$ + "\netz\*.*") <> sEmpty Then Kill currentdir$ + "\netz\*.*"
      If Dir(currentdir$ + "\skiy\*.*") <> sEmpty Then Kill currentdir$ + "\skiy\*.*"
   Else
      mypath = currentdir$ + "\netz\*.*"
      myname = Dir(mypath, vbArchive + vbNormal + vbHidden)
      Do While myname <> sEmpty
         'response = MsgBox("About to delete " & currentdir$ & "\netz\" & myname, vbOKOnly, "Cal Debug")
         Kill currentdir$ + "\netz\" + myname
         'response = MsgBox("Deleted: " & currentdir$ & "\netz\" & myname, vbOKOnly, "Cal Debug")
         myname = Dir
      Loop
      mypath = currentdir$ + "\skiy\*.*"
      myname = Dir(mypath, vbArchive + vbNormal + vbHidden)
      Do While myname <> sEmpty
         'response = MsgBox("About to delete " & currentdir$ & "\skiy\" & myname, vbOKOnly, "Cal Debug")
         Kill currentdir$ + "\skiy\" + myname
         'response = MsgBox("Deleted: " & currentdir$ & "\skiy\" & myname, vbOKOnly, "Cal Debug")
         myname = Dir
      Loop
      'response = MsgBox("FINISHED WITH THE DELETING!", vbOKCancel, "Cal Debug")
      'If response = vbCancel Then
      '   Close
      '   Exit Sub
      '   End If
      End If
      
   If sun$ = "both" Then
      nextskiy = True
      sun$ = "netz"
      End If
      
   If internet = False Then
      batnum% = FreeFile
      Open currentdir & "\" & sun$ & "\visu.bat" For Output As #batnum%
   Else
      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      'batnum% = FreeFile
      'Open currentdir & "\netz\visu.bat" For Output As #batnum% '<--can delete
      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      batdocnum%(0) = -1
      End If
   tmpnum% = FreeFile
   fildir$ = drivcities$ & "eros\" + erosareabat + "\" & sun$ & "\"
   myfile = Dir(fildir$ + "*.bat")
   If myfile <> sEmpty Then
      filnum% = FreeFile
      Open fildir$ + myfile For Input As #filnum%
      Line Input #filnum%, doclin$
      If internet = False Then
         Print #batnum%, doclin$
      Else
         '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         'Print #batnum%, doclin$ '<-----Can delete this later
         '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         batdocnum%(0) = batdocnum%(0) + 1
         batdoc$(0, batdocnum%(0)) = doclin$
         End If
      avekmxnetz = 0
      avekmynetz = 0
      avehgtnetz = 0
      mm% = 0
      Do Until EOF(filnum%)
         Input #filnum%, citnam$, lat, lon, hgt
         If InStr(LCase$(citnam$), "version") <> 0 Then
            datavernum = lat 'read data version number
            If lon = 0 Then
               SRTMflag = 0
               'don't check for near obstructions
               SunriseSunset.Check3.Value = vbUnchecked
            ElseIf lon = 1 Then
               SRTMflag = 1
               'check for near obstructions
               SunriseSunset.Check3.Value = vbChecked
            ElseIf lon = 2 Then
               SRTMflag = 2
               'check for near obstructions
               SunriseSunset.Check3.Value = vbChecked
            ElseIf lon = 9 Then
               SRTMflag = 9 'Jerusalem neighborhoods
               'check for near obstructions
               SunriseSunset.Check3.Value = vbChecked
               End If
            Exit Do 'this is the end of the batch file
            End If
         'check if this is to included
         For i% = 0 To nn% - 1
            If calnearsearchfm.List1.Selected(i%) = True Then
               If nextskiy Then
                  If InStr(calnearsearchfm.List1.List(i%), "netz") <> 0 Then
                    CalSearchName$ = Mid$(calnearsearchfm.List1.List(i%), 6, Len(calnearsearchfm.List1.List(i%)) - 5)
                    'parse citnam$
                    If InStr(citnam$, "fordtm") <> 0 Then
                       citnam2$ = Mid$(citnam$, 16, Len(citnam$) - 15)
                    Else
                       citnam2$ = citnam$
                       End If
                       
                    If InStr(drivfordtm$ & sun$ & "\" & UCase(CalSearchName$), UCase(citnam2$)) <> 0 Then
                    
                       If internet = False Then
                          Write #batnum%, citnam$, lat, lon, hgt
                       Else
                          batdocnum%(0) = batdocnum%(0) + 1
                          batdoc$(0, batdocnum%(0)) = citnam$ + "," + Str(lat) + "," + Str(lon) + "," + Str(hgt)
                          End If
                          
                       avekmxnetz = avekmxnetz + lon '/ nn%
                       avekmynetz = avekmynetz + lat '/ nn%
                       avehgtnetz = avehgtnetz + hgt '/ nn%
                       transfile$ = fildir$ + Mid$(citnam$, 16, Len(citnam$) - 15)
                       FileCopy transfile$, currentdir$ & "\" & sun$ & "\" & Mid$(citnam$, 16, Len(citnam$) - 15)
                       mm% = mm% + 1
                       Exit For
                       End If
                    End If
                 Else
                    'parse citnam$
                    If InStr(citnam$, "fordtm") <> 0 Then
                       citnam2$ = Mid$(citnam$, 16, Len(citnam$) - 15)
                    Else
                       citnam2$ = citnam$
                       End If
                       
                    If InStr(drivfordtm$ & sun$ & "\" & UCase(calnearsearchfm.List1.List(i%)), UCase(citnam2$)) <> 0 Then
                    
                       If internet = False Then
                          Write #batnum%, citnam$, lat, lon, hgt
                       Else
                          batdocnum%(0) = batdocnum%(0) + 1
                          batdoc$(0, batdocnum%(0)) = citnam$ + "," + Str(lat) + "," + Str(lon) + "," + Str(hgt)
                          End If
                          
                       avekmxnetz = avekmxnetz + lon '/ nn%
                       avekmynetz = avekmynetz + lat '/ nn%
                       avehgtnetz = avehgtnetz + hgt '/ nn%
                       transfile$ = fildir$ + Mid$(citnam$, 16, Len(citnam$) - 15)
                       FileCopy transfile$, currentdir$ & "\" & sun$ & "\" & Mid$(citnam$, 16, Len(citnam$) - 15)
                       mm% = mm% + 1
                       Exit For
                       End If
                       
                    End If
                 
                 End If
         Next i%
      Loop
      avekmxnetz = avekmxnetz / mm% 'normalize and invert
      avekmynetz = avekmynetz / mm%
      avehgtnetz = avehgtnetz / mm%
      aveusa = True
      Close #filnum%
      If internet = False Then
         Close #batnum%
      Else
         '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         'Close #batnum%
         '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         'response = MsgBox("Before writing netz bat file", vbOKOnly, "Cal Debug")
         batnum% = FreeFile
         Open currentdir & "\" & sun$ & "\visu.bat" For Output As #batnum%  '<--can delete
         For ibat% = 0 To batdocnum%(0)
            Print #batnum%, batdoc$(0, ibat%)
         Next ibat%
         Close #batnum%
         'response = MsgBox("After writing netz bat file", vbOKOnly, "Cal Debug")
         End If
   Else
      response = MsgBox("Can't find the " & erosareabat + "\" & nset$ & "*.bat file!", vbCritical + vbOKOnly, "Maps & More")
      Close #batnum%
      End If
      
   If internet = False Then
      batnum% = FreeFile
      Open currentdir & "\skiy\visu.bat" For Output As #batnum%
   Else
      batdocnum%(1) = -1
      End If
   tmpnum% = FreeFile
   '***********************TO CHANGE********************
   'at present-haven't done sunsets, so just read the netz
   'directory in case will want to calculate astronomical sunsets
   '***********************TO CHANGE********************
   '************************BEGAN TO CHANGE****************
   'if combo1.text = "sunset" then use sunset files
   '***********************BEGAN TO CHANGE*****************
   If nextskiy Then
      sun$ = "skiy"
      End If
   
   fildir$ = drivcities$ & "eros\" + erosareabat + "\" + sun$ + "\"
   'fildir$ = drivcities$ & "eros\" + erosareabat + "\netz\"
   myfile = Dir(fildir$ + "*.bat")
   If myfile <> sEmpty Then
      filnum% = FreeFile
      Open fildir$ + myfile For Input As #filnum%
      Line Input #filnum%, doclin$
      If internet = False Then
         Print #batnum%, doclin$
      Else
         batdocnum%(1) = batdocnum%(1) + 1
         batdoc$(1, batdocnum%(1)) = doclin$
         End If
      avekmxskiy = 0
      avekmyskiy = 0
      avehgtskiy = 0
      mm% = 0
      Do Until EOF(filnum%)
         Input #filnum%, citnam$, lat, lon, hgt
         If InStr(citnam$, "Version") Then Exit Do
         'check if this is to included
         For i% = 0 To nn% - 1
            If calnearsearchfm.List1.Selected(i%) = True Then
               If nextskiy Then
                  If InStr(calnearsearchfm.List1.List(i%), "skiy") <> 0 Then
                    CalSearchName$ = Mid$(calnearsearchfm.List1.List(i%), 6, Len(calnearsearchfm.List1.List(i%)) - 5)
                    'parse citnam$
                    If InStr(citnam$, "fordtm") <> 0 Then
                       citnam2$ = Mid$(citnam$, 16, Len(citnam$) - 15)
                    Else
                       citnam2$ = citnam$
                       End If
                       
                    If InStr(drivfordtm$ & sun$ + "\" & UCase(CalSearchName$), UCase(citnam2$)) <> 0 Then
                    'If InStr(drivfordtm$ & "netz\" & UCase(calnearsearchfm.List1.List(i%)), UCase(citnam2$)) <> 0 Then
        '***********************TO CHANGE********************
                       pos% = InStr(citnam$, "netz") 'non zero only if sun$=netz
                       If pos% <> 0 Then
                          citnam$ = Mid(citnam$, 1, pos% - 1) + "skiy" + Mid(citnam$, pos% + 4, Len(citnam$) - 4)
                          End If
        '***********************TO CHANGE********************
                       If internet = False Then
                          Write #batnum%, citnam$, lat, lon, hgt
                       Else
                          batdocnum%(1) = batdocnum%(1) + 1
                          batdoc$(1, batdocnum%(1)) = citnam$ + "," + Str(lat) + "," + Str(lon) + "," + Str(hgt)
                          End If
                       avekmxskiy = avekmxskiy + lon '/ nn%
                       avekmyskiy = avekmyskiy + lat '/ nn%
                       avehgtskiy = avehgtskiy + hgt '/ nn%
                       transfile$ = fildir$ + Mid$(citnam$, 16, Len(citnam$) - 15)
                       FileCopy transfile$, currentdir$ & "\skiy\" & Mid$(citnam$, 16, Len(citnam$) - 15)
                       mm% = mm% + 1
                       Exit For
                       End If
                    End If
                       
               Else
                    If InStr(citnam$, "fordtm") <> 0 Then
                       citnam2$ = Mid$(citnam$, 16, Len(citnam$) - 15)
                    Else
                       citnam2$ = citnam$
                       End If
                    If InStr(drivfordtm$ & sun$ + "\" & UCase(calnearsearchfm.List1.List(i%)), UCase(citnam2$)) <> 0 Then
                    'If InStr(drivfordtm$ & "netz\" & UCase(calnearsearchfm.List1.List(i%)), UCase(citnam2$)) <> 0 Then
        '***********************TO CHANGE********************
                       pos% = InStr(citnam$, "netz") 'non zero only if sun$=netz
                       If pos% <> 0 Then
                          citnam$ = Mid(citnam$, 1, pos% - 1) + "skiy" + Mid(citnam$, pos% + 4, Len(citnam$) - 4)
                          End If
        '***********************TO CHANGE********************
                       If internet = False Then
                          Write #batnum%, citnam$, lat, lon, hgt
                       Else
                          batdocnum%(1) = batdocnum%(1) + 1
                          batdoc$(1, batdocnum%(1)) = citnam$ + "," + Str(lat) + "," + Str(lon) + "," + Str(hgt)
                          End If
                       avekmxskiy = avekmxskiy + lon '/ nn%
                       avekmyskiy = avekmyskiy + lat '/ nn%
                       avehgtskiy = avehgtskiy + hgt '/ nn%
                       transfile$ = fildir$ + Mid$(citnam$, 16, Len(citnam$) - 15)
                       FileCopy transfile$, currentdir$ & "\skiy\" & Mid$(citnam$, 16, Len(citnam$) - 15)
                       mm% = mm% + 1
                       Exit For
                       End If
                    End If
                 End If
         Next i%
      Loop
      avekmxskiy = avekmxskiy / mm%
      avekmyskiy = avekmyskiy / mm%
      avehgtskiy = avehgtskiy / mm%
      aveusa = True
      Close #filnum%
      If internet = False Then
         Close #batnum%
      Else
         'response = MsgBox("Before writing skiy bat file", vbOKOnly, "Cal Debug")
         batnum% = FreeFile
         Open currentdir & "\skiy\visu.bat" For Output As #batnum% '<--can delete
         For ibat% = 0 To batdocnum%(1)
            Print #batnum%, batdoc$(1, ibat%)
         Next ibat%
         Close #batnum%
         'response = MsgBox("After writing skiy bat file", vbOKOnly, "Cal Debug")
         End If
      
   Else
      response = MsgBox("Can't find the " & erosareabat + "\skiy\*.bat file!", vbCritical + vbOKOnly, "Cal Program")
      Close #batnum%
      End If
      
   If internet = False Then
      searchradius = Val(Text3)
      End If
      
   'invert kmx and kmy for Israel eros files
   If InStr(erosareabat, "_Israel_Israel") Then
      avetmp = avekmxnetz
      avekmxnetz = avekmynetz
      avekmynetz = avetmp
      avetmp = avekmxskiy
      avekmxskiy = avekmyskiy
      avekmyskiy = avetmp
      aveusa = False
      End If
      
         
   If nextskiy Then
      sun$ = "both"
      nextskiy = False
      End If
      
   If eroscountry$ = "Israel" Then 'find equivalent hebrew name for neighborhood
      IsraelNeighborhood = False
      If Dir(drivcities$ & "eros\" & "Israelhebneigh_w1255.dir") <> sEmpty Then
        filneigh% = FreeFile
        Open drivcities$ & "eros\" & "Israelhebneigh_w1255.dir" For Input As #filneigh%
        Do Until EOF(filneigh%)
            Line Input #filneigh%, doclinEng$
            Line Input #filneigh%, doclinHeb$
            doclinEng$ = Replace(doclinEng$, "_", " ")
            If InStr(doclinEng$, "/" & eroscity$) <> 0 And Mid$(doclinEng$, Len(doclinEng$) - Len(eroscity$), 1) = "/" Then
               eroshebcity$ = Trim$(doclinHeb$)
'               hebcityname$ = eroshebcity$
'               If optionheb Then eroscity$ = hebcityname$
                IsraelNeighborhood = True
               Exit Do
               End If
        Loop
        Close #filneigh%
      Else
        eroshebcity$ = hebcityname$
        End If
      End If
      
      
   Call Form_QueryUnload(i%, j%)
   Screen.MousePointer = vbHourglass
   SunriseSunset.Timer1.Enabled = False
   SunriseSunset.Show
   SunriseSunset.Enabled = True
   SunriseSunset.ProgressBar1.Enabled = True
   SunriseSunset.ProgressBar1.Visible = True
   'SunriseSunset.Label3.Visible = False
   If InStr(currentdir, "eros") <> 0 Then
      eros = True
      geo = True
      End If
   SunriseSunset.Visible = True
   If internet = True Then
        SunriseSunset.Check4.Enabled = False
        SunriseSunset.Check5.Enabled = False
        SunriseSunset.Combo1.Enabled = False
        SunriseSunset.Label2.Enabled = False
        SunriseSunset.Cancelbut.Enabled = False
        If SunriseSunset.Label1.Caption <> sEmpty Then SunriseSunset.Label1.Caption = captmp$
        SunriseSunset.Label1.Enabled = False
        SunriseSunset.OKbut0.Enabled = False
        SunriseSunset.ProgressBar1.Visible = False
        SunriseSunset.Option2.Enabled = False
        SunriseSunset.Option1.Enabled = False
        SunriseSunset.Label2.Enabled = False
        SunriseSunset.Label4.Enabled = False
        SunriseSunset.Label5.Enabled = False
        SunriseSunset.Label6.Enabled = False
        SunriseSunset.Label7.Enabled = False
        SunriseSunset.Text1.Enabled = False
        SunriseSunset.UpDown1.Enabled = False
        SunriseSunset.Text2.Enabled = False
        SunriseSunset.UpDown2.Enabled = False
        'SunriseSunset.Check2.Enabled = False
        If sun$ = "skiy" And viseros Then
           SunriseSunset.Check1.Value = vbUnchecked
           SunriseSunset.Check1.Enabled = False
           SunriseSunset.Check2.Value = vbChecked
           SunriseSunset.Check2.Enabled = True
        ElseIf sun$ = "netz" And viseros Then
           SunriseSunset.Check2.Value = vbUnchecked
           SunriseSunset.Check2.Enabled = False
           SunriseSunset.Check1.Value = vbChecked
           SunriseSunset.Check1.Enabled = True
        ElseIf sun$ = "both" And viseros Then
           SunriseSunset.Check2.Value = vbChecked
           SunriseSunset.Check2.Enabled = True
           SunriseSunset.Check1.Value = vbChecked
           SunriseSunset.Check1.Enabled = True
           End If
   Else
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
        SunriseSunset.Option2.Enabled = True
        SunriseSunset.Option1.Enabled = True
        SunriseSunset.Label2.Enabled = True
        SunriseSunset.Label4.Enabled = True
        SunriseSunset.Label5.Enabled = True
        SunriseSunset.Label6.Enabled = True
        SunriseSunset.Label7.Enabled = True
        SunriseSunset.Text1.Enabled = True
        SunriseSunset.UpDown1.Enabled = True
        SunriseSunset.Text2.Enabled = True
        SunriseSunset.UpDown2.Enabled = True
        If sun$ = "skiy" Then
           SunriseSunset.Check1.Value = vbUnchecked
           SunriseSunset.Check1.Enabled = False
           SunriseSunset.Check2.Value = vbChecked
           SunriseSunset.Check2.Enabled = True
        ElseIf sun$ = "netz" Then
           SunriseSunset.Check2.Value = vbUnchecked
           SunriseSunset.Check2.Enabled = False
           SunriseSunset.Check1.Value = vbChecked
           SunriseSunset.Check1.Enabled = True
         ElseIf sun$ = "both" Then
           SunriseSunset.Check2.Value = vbChecked
           SunriseSunset.Check2.Enabled = True
           SunriseSunset.Check1.Value = vbChecked
           SunriseSunset.Check1.Enabled = True
          End If
        End If
   hebcal = True
   Option1b = True
   Option2b = False
   If internet = True Then
      SunriseSunset.OKbut0.Value = True
      End If
   'SunriseSunset.Combo1.Text = 5758
   'ret = SetWindowPos(SunriseSunset.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If suntop% <> 0 Then SunriseSunset.Top = suntop%
   Screen.MousePointer = vbDefault
   Exit Sub
   
searcherrhand:
   Screen.MousePointer = vbDefault
'   Resume
   If internet = True Then
   '************DEBUG VERSION*********************
      'response = MsgBox("Encountered Error Number: " & Err.Number & "; ABORTING!", vbOKOnly + vbCritical, "Cal Debug")
   '************DEBUG VERSION*********************
        errlog% = FreeFile
        Open drivjk$ + "Calprog.log" For Output As #errlog%
        Print #errlog%, "Cal Prog exited from calnearsearchfrm with runtime error message " + Str(Err.Number)
        Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
        Close #errlog%
        Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        myfile = Dir(drivfordtm$ + "busy.cal")
        If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
     Else
       response = MsgBox("Encountered error number:" & Str$(Err.Number) & vbLf & Err.Description & vbLf & vbLf & "Aborting this operation", vbOKOnly + vbCritical, "Cal Program")
       End If
   
End Sub

Private Sub Command2_Click()
   Call Form_QueryUnload(i%, j%)
End Sub

Private Sub Command3_Click()

   On Error GoTo errhand
   
   eroslongitude = Text1
   eroslatitude = Text2

   'If Combo1.Text = "sunrise" Then 'Or internet = True Then '<--future
   If Combo1.Text = "sunrise" Then 'Or internet = True Then
      sun$ = "netz"
   '-----------------in the future visible sunset for internet---
   'ElseIf Combo1.Text = "sunset" Then
   ElseIf Combo1.Text = "sunset" Then 'And internet = False Then
      sun$ = "skiy"
   ElseIf Combo1.Text = "both" Then
      sun$ = "both"
      GoTo ret1500
      End If
   If internet = False Then
      If InStr(calnode.StatusBar1.Panels(2).Text, "Israel_Israel") <> 0 Then
         eroscountry$ = "Israel"
         End If
      End If
ret50:
   fildir$ = drivcities$ & "eros\" + erosareabat + "\" + sun$ + "\"
   myfile = Dir(fildir$ + "*.bat")
  If myfile <> sEmpty Then
      filnum% = FreeFile
      Open fildir$ + myfile For Input As #filnum%
      'pi = 3.14159265
      ydegkm = 1# / (cd * 6371.315)
      xdegkm = 1# / (cd * 6371.315 * Cos(CDbl(eroslatitude) * cd))
      Line Input #filnum%, doclin$
      nn% = 0
      List1.Clear
      nn% = 0
      Do Until EOF(filnum%)
         Input #filnum%, citnam$, lat, lon, hgt
         If InStr(LCase(citnam$), "version") Then
            'read version number
            If Len(citnam$) > 7 Then
                pos% = InStr(citnam$, ",")
                pos2% = InStr(citnam$, pos2% + 1, ",")
                datavernum = Val(Mid$(citnam$, pos% + 1, pos2% - pos%))
                GoTo 500
            Else
               datavernum = lat
               GoTo 500
               End If
            End If
         If InStr(eroscountry$, "Israel") = 0 Then
            dist = Sqr(((lat - CDbl(eroslatitude)) / ydegkm) ^ 2 + ((lon - CDbl(eroslongitude)) / xdegkm) ^ 2)
         Else
            dist = Sqr((lon - CDbl(eroslatitude)) ^ 2 + (lat - CDbl(eroslongitude)) ^ 2)
            End If
         If dist <= Val(calnearsearchfm.Text3) + 0.001 Then
            calnearsearchfm.List1.AddItem Mid$(citnam$, 16, Len(citnam$) - 15) & "," & Format(lon, "####.00000") & "," & Format(lat, "###.00000") & "," & Format(hgt, "###0.0") & ", dist = " & Format(dist, "#0.0")
            calnearsearchfm.List1.Selected(nn%) = True
            nn% = nn% + 1
            End If
500
      Loop
      Close #filnum%
      If nn% > 0 Then GoSub VisGoogle
      If internet = True And nn% = 0 Then 'just output error reports
         errorreport = True
         
         'exit program with error message
         filout$ = dirint$ + "\" + Mid(servnam$, 1, 8) + ".html"
         filoutnum% = FreeFile
         Open filout$ For Output As filoutnum%
         'give error report
         Print #filoutnum%, "<!doctype html public ""-//W3C//DTD HTML 4.0 //EN"">"
         Print #filoutnum%, "<html>"
         Print #filoutnum%, "<head>"
         Print #filoutnum%, "<title>Chai Tables Error report</title>"
         Print #filoutnum%, "<meta name="; Author; " content=""Chaim Keller"">"
         Print #filoutnum%, "</head>"
         Print #filoutnum%, "<BODY>"
         Print #filoutnum%, "<H2>The Chai Tables</H2>"
         Print #filoutnum%, "<P>"
         Print #filoutnum%, "<table bgcolor=""#000000"" border=0 cellpadding=1 cellspacing=0><tr><td>"
         Print #filoutnum%, "<table bgcolor=""#ffffcc"" border=0 cellpadding=8  cellspacing=0><tr valign=bottom><td>"
         Print #filoutnum%, "<font face=""Arial,Helvetica"": size=2>"
         Print #filoutnum%, "No calculated vantage point could be found. Please check your inputs.  If you wish, you can attempt to increase the search radius, or calculate astronimical times"
         Print #filoutnum%, "<p>"
         Print #filoutnum%, " </p>"
         Print #filoutnum%, "</body>"
         Print #filoutnum%, "</html>"
         Close #filoutnum%
         
         'repeat this for zemanim table if requested
         If zmanyes% = 1 Then
            filroot$ = dirint$ + "\" + Mid(servnam$, 1, 8)
            If zmantype% = 0 Then
               ext$ = "csv"
            ElseIf zmantype% = 1 Then
               ext$ = "zip"
            ElseIf zmantype% = 2 Then
               ext$ = "xml"
               End If
               
            Call WriteTables(filroot$, servnam$, ext$)
            End If
         
        'give a bit of time to write the reports
        waitime = Timer + 1
        Do Until Timer > waitime
           DoEvents
        Loop
         
        'now record error message and end the program
         
         myfile = Dir(drivfordtm$ + "busy.cal")
         If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
                 
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "No vantage point within search radius!"
         Print #lognum%, "Program termination called from calnearsearchfm:Command3"
         Close #lognum%
                 
         For i% = 0 To Forms.Count - 1
           Unload Forms(i%)
         Next i%
                  
         'kill timer
         If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
        
         'end program abruptly
         End
         
         
         End If
   Else
      If internet = True Then 'check if want astronomical sunset
         If viseros = False Then
            sun$ = "netz"
            GoTo ret50
         Else 'visible directory not found, abort
            myfile = Dir(drivfordtm$ + "busy.cal")
            If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
                 
            lognum% = FreeFile
            Open drivjk$ + "calprog.log" For Append As #lognum%
            If Err.Number <> 0 Then
               Print #lognum%, "Visible directory: " & fildir$; " not found!"
               End If
            Print #lognum%, "Program termination called from calnearsearchfm:Command3"
            Close #lognum%
                 
            For i% = 0 To Forms.Count - 1
              Unload Forms(i%)
            Next i%
                  
            'kill timer
            If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
        
            'end program abruptly
            End
         
            End If
         End If
      response = MsgBox(Combo1.Text & " directory not found for this place!", vbCritical + vbOKOnly, "Cal Program")
      Exit Sub
      End If
      StatusBar1.Visible = True
      StatusBar1.Panels(1) = "Check the desired places, and press OK"
      foundvantage = True
      Exit Sub
      
'////////////////////////////new as of 05/13/16////////////////////////////
ret1500: 'search both netz and skiy directories
   fildir$ = drivcities$ & "eros\" + erosareabat + "\netz\"
   myfile = Dir(fildir$ + "*.bat")
  If myfile <> sEmpty Then
      filnum% = FreeFile
      Open fildir$ + myfile For Input As #filnum%
      'pi = 3.14159265
      ydegkm = 1# / (cd * 6371.315)
      xdegkm = 1# / (cd * 6371.315 * Cos(CDbl(eroslatitude) * cd))
      Line Input #filnum%, doclin$
      nn% = 0
      List1.Clear
      Do Until EOF(filnum%)
         Input #filnum%, citnam$, lat, lon, hgt
         If InStr(LCase(citnam$), "version") Then
            'read version number
            If Len(citnam$) > 7 Then
                pos% = InStr(citnam$, ",")
                pos2% = InStr(citnam$, pos2% + 1, ",")
                datavernum = Val(Mid$(citnam$, pos% + 1, pos2% - pos%))
                GoTo 600
            Else
               datavernum = lat
               GoTo 600
               End If
            End If
         If InStr(eroscountry$, "Israel") = 0 Then
            dist = Sqr(((lat - CDbl(eroslatitude)) / ydegkm) ^ 2 + ((lon - CDbl(eroslongitude)) / xdegkm) ^ 2)
         Else
            dist = Sqr((lon - CDbl(eroslatitude)) ^ 2 + (lat - CDbl(eroslongitude)) ^ 2)
            End If
         If dist <= Val(calnearsearchfm.Text3) + 0.001 Then
            calnearsearchfm.List1.AddItem "netz: " & Mid$(citnam$, 16, Len(citnam$) - 15) & "," & Format(lon, "####.00000") & "," & Format(lat, "###.00000") & "," & Format(hgt, "###0.0") & ", dist = " & Format(dist, "#0.0")
            calnearsearchfm.List1.Selected(nn%) = True
            nn% = nn% + 1
            End If
600
      Loop
      Close #filnum%
      GoSub VisGoogle
   Else
      response = MsgBox(Combo1.Text & " directory not found for this place!", vbCritical + vbOKOnly, "Cal Program")
      Exit Sub
      End If
      
   fildir$ = drivcities$ & "eros\" + erosareabat + "\skiy\"
   myfile = Dir(fildir$ + "*.bat")
  If myfile <> sEmpty Then
      filnum% = FreeFile
      Open fildir$ + myfile For Input As #filnum%
      'pi = 3.14159265
      ydegkm = 1# / (cd * 6371.315)
      xdegkm = 1# / (cd * 6371.315 * Cos(CDbl(eroslatitude) * cd))
      Line Input #filnum%, doclin$
      Do Until EOF(filnum%)
         Input #filnum%, citnam$, lat, lon, hgt
         If InStr(LCase(citnam$), "version") Then
            'read version number
            If Len(citnam$) > 7 Then
                pos% = InStr(citnam$, ",")
                pos2% = InStr(citnam$, pos2% + 1, ",")
                datavernum = Val(Mid$(citnam$, pos% + 1, pos2% - pos%))
                GoTo 700
            Else
               datavernum = lat
               GoTo 700
               End If
            End If
         If InStr(eroscountry$, "Israel") = 0 Then
            dist = Sqr(((lat - CDbl(eroslatitude)) / ydegkm) ^ 2 + ((lon - CDbl(eroslongitude)) / xdegkm) ^ 2)
         Else
            dist = Sqr((lon - CDbl(eroslatitude)) ^ 2 + (lat - CDbl(eroslongitude)) ^ 2)
            End If
         If dist <= Val(calnearsearchfm.Text3) + 0.001 Then
            calnearsearchfm.List1.AddItem "skiy: " & Mid$(citnam$, 16, Len(citnam$) - 15) & "," & Format(lon, "####.00000") & "," & Format(lat, "###.00000") & "," & Format(hgt, "###0.0") & ", dist = " & Format(dist, "#0.0")
            calnearsearchfm.List1.Selected(nn%) = True
            nn% = nn% + 1
            End If
700
      Loop
      Close #filnum%
      GoSub VisGoogle
   Else
      response = MsgBox(Combo1.Text & " directory not found for this place!", vbCritical + vbOKOnly, "Cal Program")
      Exit Sub
      End If
      
   Exit Sub
      
VisGoogle:
   If nn% > 0 Then
      calnearsearchfm.StatusBar1.Panels(1).Text = "Dblclick for map"
      calnearsearchfm.StatusBar1.Panels(1).ToolTipText = "Double click to locate on map"
      cboGoogle.Visible = True
   Else
      calnearsearchfm.StatusBar1.Panels(1).Text = sEmpty
      calnearsearchfm.StatusBar1.Panels(1).ToolTipText = sEmpty
      cboGoogle.Visible = False
      End If
   Return
      
errhand:
      If nn% > 0 Then
        Close
        StatusBar1.Visible = True
        StatusBar1.Panels(1) = "Check the desired places, and press OK"
        foundvantage = True
        End If
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
  cboGoogle.ListIndex = 0
  calnearsearchVis = True
  eros = True
  
  Combo1.AddItem "sunrise"
  Combo1.AddItem "sunset"
  Combo1.AddItem "both"
  Combo1.ListIndex = 2
  If internet Then
     If nsetflag% = 1 Then 'sunrise
        Combo1.Text = "sunrise"
     ElseIf nsetflag% = 2 Then 'sunset
        Combo1.Text = "sunset"
     ElseIf nsetflag% = 3 Then
        Combo1.Text = "both"
        End If
     End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set calnearsearchfm = Nothing
   calnearsearchVis = False
End Sub
'---------------------------------------------------------------------------------------
' Procedure : List1_DblClick
' Author    : chaim
' Date      : 1/11/2022
' Purpose   : double click to show position on Google Map,
' based on  : http://www.vb-helper.com/howto_google_map.html
'---------------------------------------------------------------------------------------
'
Private Sub List1_DblClick()

   On Error GoTo List1_DblClick_Error
   
   
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

    '39.358008,-76.688316
    waitime = Timer
    Do Until Timer > waitime + 0.1
       DoEvents
    Loop
    PntSelected = Not calnearsearchfm.List1.Selected(List1.ListIndex) 'store checked status
    DataLine = Split(calnearsearchfm.List1.List(List1.ListIndex), ",")
    longitude = Val(DataLine(1))
    latitude = Val(DataLine(2))
    
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
      calnearsearchfm.List1.Selected(List1.ListIndex) = True
   Else
      calnearsearchfm.List1.Selected(List1.ListIndex) = False
      End If

   On Error GoTo 0
   Exit Sub

List1_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure List1_DblClick of Form calnearsearchfm"
End Sub
