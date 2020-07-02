VERSION 5.00
Begin VB.Form mapVersionfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Maps & More"
   ClientHeight    =   1590
   ClientLeft      =   2310
   ClientTop       =   1455
   ClientWidth     =   3705
   Icon            =   "mapVersionfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Label Label3 
         Caption         =   "This program is protected by international copyright law"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "© CNK 1999-2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Version: 2.0.12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1370
      Left            =   120
      Picture         =   "mapVersionfm.frx":0442
      Stretch         =   -1  'True
      Top             =   110
      Width           =   1200
   End
End
Attribute VB_Name = "mapVersionfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
   Dim lVersion As Long
   Label1.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   Label2.Caption = "© CNK 1999-" + LTrim$(Str(Year(Date)))
   ret = SetWindowPos(mapVersionfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
End Sub
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub Image1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub Label1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub Label2_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub Label3_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub Frame1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub

