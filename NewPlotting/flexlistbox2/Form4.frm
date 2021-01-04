VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   Caption         =   "Properties From Test FlexListBox"
   ClientHeight    =   5100
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   4968
   LinkTopic       =   "Form3"
   ScaleHeight     =   5100
   ScaleWidth      =   4968
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Info"
      Height          =   300
      Left            =   2976
      TabIndex        =   5
      Top             =   2976
      Width           =   1740
   End
   Begin VB.CheckBox Check 
      Caption         =   "Locked"
      Height          =   300
      Index           =   10
      Left            =   2832
      TabIndex        =   27
      Top             =   1872
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "NoDeSelect"
      Height          =   300
      Index           =   9
      Left            =   2832
      TabIndex        =   26
      Top             =   1632
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "AutoSelectItem"
      Height          =   300
      Index           =   8
      Left            =   2832
      TabIndex        =   25
      Top             =   1392
      Width           =   2076
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1200
      TabIndex        =   24
      Text            =   "0"
      ToolTipText     =   "Height of the ListItems Range 100 - 800"
      Top             =   4608
      Width           =   636
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1200
      TabIndex        =   22
      Text            =   "0"
      Top             =   4224
      Width           =   300
   End
   Begin VB.Frame Frame5 
      Caption         =   "Fonts"
      Height          =   1212
      Left            =   96
      TabIndex        =   17
      Top             =   48
      Width           =   2652
      Begin VB.CommandButton cmdFont 
         Caption         =   "FontDisabeled"
         Height          =   252
         Index           =   0
         Left            =   96
         TabIndex        =   20
         Top             =   240
         Width           =   2412
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "FontUnselected"
         Height          =   252
         Index           =   1
         Left            =   96
         TabIndex        =   19
         Top             =   576
         Width           =   2412
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "FontSelected"
         Height          =   252
         Index           =   2
         Left            =   96
         TabIndex        =   18
         Top             =   912
         Width           =   2412
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   288
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Color"
      Height          =   1980
      Left            =   96
      TabIndex        =   10
      Top             =   1680
      Width           =   2652
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorForeSelected"
         Height          =   252
         Index           =   5
         Left            =   96
         TabIndex        =   16
         Top             =   1632
         Width           =   2412
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorBackSelected"
         Height          =   252
         Index           =   4
         Left            =   96
         TabIndex        =   15
         Top             =   1392
         Width           =   2412
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorForeUnselected"
         Height          =   252
         Index           =   3
         Left            =   96
         TabIndex        =   14
         Top             =   1056
         Width           =   2412
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorBackUnselected"
         Height          =   252
         Index           =   2
         Left            =   96
         TabIndex        =   13
         Top             =   816
         Width           =   2412
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorForeDisabeled"
         Height          =   252
         Index           =   1
         Left            =   96
         TabIndex        =   12
         Top             =   480
         Width           =   2412
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "ColorBackDisabeled"
         Height          =   252
         Index           =   0
         Left            =   96
         TabIndex        =   11
         Top             =   240
         Width           =   2412
      End
      Begin VB.Line Line2 
         X1              =   96
         X2              =   2448
         Y1              =   1344
         Y2              =   1344
      End
      Begin VB.Line Line1 
         X1              =   96
         X2              =   2448
         Y1              =   768
         Y2              =   768
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ListStyle"
      Height          =   636
      Left            =   2880
      TabIndex        =   8
      Top             =   4416
      Width           =   2028
      Begin VB.ComboBox FlexListStyle 
         Height          =   288
         Left            =   96
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   240
         Width           =   1836
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alignment"
      Height          =   636
      Left            =   2880
      TabIndex        =   7
      Top             =   3792
      Width           =   2028
      Begin VB.ComboBox FlexAlignment 
         Height          =   288
         Left            =   48
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   240
         Width           =   1884
      End
   End
   Begin VB.CheckBox Check 
      Caption         =   "MultiSelect"
      Height          =   300
      Index           =   7
      Left            =   2832
      TabIndex        =   4
      Top             =   1164
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "MousePreSelector "
      Height          =   300
      Index           =   6
      Left            =   2832
      TabIndex        =   3
      Top             =   888
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "Enabled"
      Height          =   300
      Index           =   5
      Left            =   2832
      TabIndex        =   2
      Top             =   624
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "CopyColorItemsToIcons"
      Height          =   300
      Index           =   3
      Left            =   2832
      TabIndex        =   1
      Top             =   348
      Width           =   2076
   End
   Begin VB.CheckBox Check 
      Caption         =   "ListExitOnSelection"
      Height          =   300
      Index           =   1
      Left            =   2832
      TabIndex        =   0
      Top             =   96
      Width           =   2076
   End
   Begin VB.Label Label4 
      Caption         =   "ItemHeight"
      Height          =   252
      Left            =   192
      TabIndex        =   23
      Top             =   4656
      Width           =   828
   End
   Begin VB.Label Label1 
      Caption         =   "Columns"
      Height          =   204
      Left            =   192
      TabIndex        =   21
      Top             =   4272
      Width           =   732
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                                Eric de Decker secr VB Belgie
'                                E.B. our API Killer  ProFinance Woerden
'                                Willem secr VBgroup Ned vbg@vbgroup.nl

'This sources may be used freely without the intention of commercial distribution. For all
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

Private Flex As FlexListBox
Private noClick As Boolean

Enum Fonts
   FontDisabeled = 0
   FontUnselected = 1
   FontSelected = 2
End Enum

Enum Colors
   BackColorDisabeled = 0
   ForeColorDisabeled = 1
   BackColorUnselected = 2
   ForeColorUnselected = 3
   BackColorSelected = 4
   ForeColorSelected = 5
End Enum

Public Sub ShowFlexValues(o As FlexListBox)
Set Flex = o
Form3.Show
noClick = False

Check(1).Value = Flex.ListExitOnSelection And vbChecked
Check(3).Value = Flex.CopyColorItemsToIcons And vbChecked
Check(5).Value = Flex.Enabled And vbChecked
Check(6).Value = Flex.MousePreSelector And vbChecked
Check(7).Value = Flex.MultiSelect And vbChecked
Check(8).Value = Flex.AutoSelectItem And vbChecked
Check(9).Value = Flex.NoDeSelect And vbChecked
Check(10).Value = Flex.Locked And vbChecked

Call ReadValues

noClick = True
End Sub

Public Sub ReadValues()
Text1 = CStr(Flex.Columns)
Text2 = CStr(Flex.ItemHeight)
End Sub

Private Sub Check_Click(Index As Integer)
If noClick Then
   With Flex
      .ListExitOnSelection = Check(1).Value = vbChecked
      .CopyColorItemsToIcons = Check(3).Value = vbChecked
      .Enabled = Check(5).Value = vbChecked
      .MousePreSelector = Check(6).Value = vbChecked
      .MultiSelect = Check(7).Value = vbChecked
      .AutoSelectItem = Check(8).Value = vbChecked
      .NoDeSelect = Check(9).Value = vbChecked
      .Locked = Check(10).Value = vbChecked
      .Refresh
   End With
End If

End Sub

Private Sub cmdColor_Click(Index As Integer)
With CommonDialog1
   .CancelError = True
   On Error GoTo ErrHandler
   .Flags = cdlCCRGBInit
   .DialogTitle = cmdColor(Index).Caption
   Select Case Index
      Case Colors.BackColorDisabeled
         .Color = Flex.BackColorDisabeled
      Case Colors.ForeColorDisabeled
         .Color = Flex.ForeColorDisabeled
      Case Colors.BackColorUnselected
         .Color = Flex.BackColorUnselected
      Case Colors.ForeColorUnselected
         .Color = Flex.ForeColorUnselected
      Case Colors.BackColorSelected
         .Color = Flex.BackColorSelected
      Case Colors.ForeColorSelected
         .Color = Flex.ForeColorSelected
   End Select
   .ShowColor
   Select Case Index
      Case Colors.BackColorDisabeled
         Flex.BackColorDisabeled = .Color
      Case Colors.ForeColorDisabeled
         Flex.ForeColorDisabeled = .Color
      Case Colors.BackColorUnselected
         Flex.BackColorUnselected = .Color
      Case Colors.ForeColorUnselected
         Flex.ForeColorUnselected = .Color
      Case Colors.BackColorSelected
         Flex.BackColorSelected = .Color
      Case Colors.ForeColorSelected
         Flex.ForeColorSelected = .Color
   End Select
End With
Flex.RefreshControls
Exit Sub

ErrHandler:
  ' User pressed the Cancel button
End Sub

Private Sub cmdFont_Click(Index As Integer)
Dim oFont As Font
Set oFont = Nothing
With CommonDialog1
   .CancelError = True
   On Error GoTo ErrHandler
   .Flags = cdlCFEffects Or cdlCFBoth
   .DialogTitle = cmdColor(Index).Caption

   Select Case Index
      Case FontDisabeled
         Set oFont = Flex.FontDisabeled
      Case FontUnselected
         Set oFont = Flex.FontUnselected
      Case FontSelected
         Set oFont = Flex.FontSelected
   End Select

   .FontName = oFont.Name
   .FontSize = oFont.Size
   .FontBold = oFont.Bold
   .FontItalic = oFont.Italic
   .FontUnderline = oFont.Underline
   .ShowFont
   oFont.Name = .FontName
   oFont.Size = .FontSize
   oFont.Bold = .FontBold
   oFont.Italic = .FontItalic
   oFont.Underline = .FontUnderline

End With

Flex.RefreshControls 'needed to refresh all

Exit Sub

ErrHandler:
  ' User pressed the Cancel button
End Sub


Private Sub Command2_Click()
Dim item As List
For Each item In Flex.List
   If Not item.Enabeled Then item.Visible = Not item.Visible
Next
Flex.Refresh
End Sub

Private Sub Command3_Click()
If Flex.List.Count > 0 Then
   Flex.RemoveItem Flex.List.Count
End If
End Sub

Private Sub Command4_Click()
Flex.AddItem CStr(Flex.List.Count + 1)
End Sub

Private Sub Command1_Click()
Call Flex.ShowSplash
End Sub

Private Sub FlexAlignment_Click()
Flex.Alignment = FlexAlignment.ListIndex
Flex.Refresh
End Sub

Private Sub FlexListStyle_LostFocus()
Flex.ListStyle = FlexListStyle.ListIndex
End Sub

Private Sub Form_Load()
With FlexAlignment
   .Clear
   .AddItem vbLeftJustify & " - vbLeftJustify"
   .AddItem vbRightJustify & " - vbRightJustify"
   .AddItem vbCenter & " - vbCenter"
   .ListIndex = Flex.Alignment
End With

With FlexListStyle
   .AddItem "Normal"
   .AddItem "CheckBox"
   .AddItem "PictureBox"
   .AddItem "Image"
   .ListIndex = Flex.ListStyle
End With


End Sub

Private Sub Text1_LostFocus()
Flex.Columns = Abs(Val(Text1))
ReadValues
End Sub

Private Sub Text2_LostFocus()
Flex.ItemHeight = Abs(Val(Text2))
ReadValues
End Sub


