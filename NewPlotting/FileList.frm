VERSION 5.00
Begin VB.Form FileList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File List"
   ClientHeight    =   3330
   ClientLeft      =   2370
   ClientTop       =   3270
   ClientWidth     =   6540
   Icon            =   "FileList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6540
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6435
   End
End
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set FileList = Nothing
End Sub
