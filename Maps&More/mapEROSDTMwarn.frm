VERSION 5.00
Begin VB.Form mapEROSDTMwarn 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "         USGS EROS DEM CD not found!"
   ClientHeight    =   975
   ClientLeft      =   6285
   ClientTop       =   7530
   ClientWidth     =   4590
   Icon            =   "mapEROSDTMwarn.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   4575
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "CD #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "In order to display heights in this         region you must insert "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "mapEROSDTMwarn.frx":0442
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "mapEROSDTMwarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
   ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
    Unload mapEROSDTMwarn
    Set mapEROSDTMwarn = Nothing
    ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

