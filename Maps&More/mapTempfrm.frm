VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mapTempfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WorldClim Temperature Model Ver. 2"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6840
   Icon            =   "mapTempfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame TKfrm 
      Caption         =   "Termperatures"
      Height          =   5220
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid msFlxGrdTK 
         Height          =   4620
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8149
         _Version        =   393216
         Rows            =   13
         Cols            =   4
         BackColor       =   -2147483624
         BackColorFixed  =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Loadfrm 
      Caption         =   "Load Temperature Data"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.CommandButton cmdLoadTK 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Load Temps for current coordinates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "mapTempfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : cmdLoadTK_Click
' Author    : Dr-John-K-Hall
' Date      : 11/6/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLoadTK_Click()
    Dim lt1 As Double, lg1 As Double, i As Integer
    Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, ier As Integer
    
   On Error GoTo cmdLoadTK_Click_Error

    lg1 = Maps.Text5.Text
    lt1 = Maps.Text6.Text
    
    If Not world Then
       'EY ITM, convert to geo coordinates
       Call casgeo(lg1, lt1, lg, lt)
       lg1 = -lg
       lt1 = lt
    Else
'       tmplt = lt1
'       lt1 = lg1
'       lg1 = -tmplt
       End If
       
    Call Temperatures(lt1, lg1, MinT, AvgT, MaxT, ier)
    
    With msFlxGrdTK
       .ColAlignment(1) = 4
       .ColAlignment(2) = 4
       .ColAlignment(3) = 4
       For i = 1 To 12
         .TextMatrix(i, 1) = MinT(i)
         .TextMatrix(i, 2) = AvgT(i)
         .TextMatrix(i, 3) = MaxT(i)
       Next i
    End With

   On Error GoTo 0
   Exit Sub

cmdLoadTK_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdLoadTK_Click of Form mapTempfrm"
    
End Sub

Private Sub Form_Load()
   With msFlxGrdTK
      .ColAlignment(0) = 1
      .TextMatrix(1, 0) = "January"
      .TextMatrix(2, 0) = "February"
      .TextMatrix(3, 0) = "March"
      .TextMatrix(4, 0) = "April"
      .TextMatrix(5, 0) = "May"
      .TextMatrix(6, 0) = "June"
      .TextMatrix(7, 0) = "July"
      .TextMatrix(8, 0) = "August"
      .TextMatrix(9, 0) = "September"
      .TextMatrix(10, 0) = "October"
      .TextMatrix(11, 0) = "November"
      .TextMatrix(12, 0) = "December"
      .TextMatrix(0, 1) = "Min. Temp."
      .TextMatrix(0, 2) = "Avg. Temp."
      .TextMatrix(0, 3) = "Max. Temp."
   End With
   
   cmdLoadTK_Click
   TempFormVis = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

   TempFormVis = False
   tblbuttons(29) = 0
   Maps.Toolbar1.Buttons(29).value = tbrUnpressed
   Set mapTempfrm = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:
   Resume Next
End Sub
