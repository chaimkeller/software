VERSION 5.00
Begin VB.Form FontRotate 
   AutoRedraw      =   -1  'True
   Caption         =   "FontRotate"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1560
      ScaleHeight     =   1995
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "FontRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetClientRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)

'****************************************************************************************************
'
' Name          : FontRotate
' Author        : Dennis Burns
' Email         : nextlemming@aol.com
' Date          : Sept 26, 2001
' Description   : Program to demonstrate text rotation.
'
'
' Notice        : This code is open to the public domain,
'                   just give credit where credit is due.
'
'****************************************************************************************************

' In this example a picture box is used as the target for the text.
' Any object for which you can get an hdc could be used.

Private Sub Form_Load()

    
    Dim rc As RECT          'Rectangle structure to hold screen area
    Dim Result As Long      'Holds result of api calls
    
    'get rectangle of client area of window to calculate placement.
    Result = GetClientRect(Picture1.hwnd, rc)
    
    ' Setting aout redraw to true causes windows to make a persistant
    '   image for the picture, without it the text would be lost the next
    '   time the picture was redrawn.
    Picture1.AutoRedraw = True
    
    'Call function to print text
    Call TextRotate(Picture1.hdc, "test", rc.Left * 0.9, rc.Bottom * 0.9, 90)
    
    
End Sub



