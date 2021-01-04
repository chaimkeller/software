VERSION 5.00
Begin VB.UserControl FlexListBox 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   DataBindingBehavior=   1  'vbSimpleBound
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   -1  'True
   EndProperty
   KeyPreview      =   -1  'True
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3435
   ToolboxBitmap   =   "UserControl1.ctx":003D
   Begin VB.PictureBox ToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      FillColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   204
      HelpContextID   =   310
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   945
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   972
      Begin VB.Label LblToolTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "ToolTip"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   0
         WhatsThisHelpID =   10
         Width           =   465
      End
   End
   Begin VB.PictureBox ListBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2220
      HelpContextID   =   340
      Left            =   624
      ScaleHeight     =   2190
      ScaleWidth      =   2625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      WhatsThisHelpID =   10
      Width           =   2652
      Begin VB.CheckBox picItems 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         HelpContextID   =   380
         Index           =   0
         Left            =   0
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   300
      End
      Begin VB.CheckBox chkItems 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         HelpContextID   =   370
         Index           =   0
         Left            =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   195
      End
      Begin VB.HScrollBar hScrollBar 
         Height          =   180
         HelpContextID   =   360
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1968
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   1020
      End
      Begin VB.VScrollBar vScrollBar 
         Height          =   2172
         HelpContextID   =   350
         Left            =   2400
         Max             =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   220
      End
      Begin VB.PictureBox Slider 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   100
         Index           =   0
         Left            =   0
         MouseIcon       =   "UserControl1.ctx":034F
         MousePointer    =   99  'Custom
         ScaleHeight     =   75
         ScaleWidth      =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   50
      End
      Begin VB.Image imgItems 
         Height          =   200
         Index           =   0
         Left            =   70
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   200
      End
      Begin VB.Label lblItems 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   2316
         WordWrap        =   -1  'True
      End
      Begin VB.Line SliderLine1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Visible         =   0   'False
         X1              =   0
         X2              =   1344
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line SliderLine2 
         Visible         =   0   'False
         X1              =   0
         X2              =   1488
         Y1              =   100
         Y2              =   100
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopUpSub 
         Caption         =   "PopUpSub"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FlexListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                                Eric de Decker secr VB Belgie
'                                E.B. our API Killer  ProFinance Woerden
'                                Willem secr VBgroup Ned vbg@vbgroup.nl

'This sources may be used freely without the intention of commercial distribution. For all
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.


Public List                   As New ListCollection

Private bHoldSlider           As Boolean
Private lTopIndex             As Long
Private iMaxVisible           As Integer
Private iLastPos              As Integer
Private iMaxLenItem           As Integer
Private iIndex                As Integer
Private iShift                As Integer
Private bLastKeyReturn        As Boolean
Private bMouseClick           As Boolean
Private bRefreshControls      As Boolean
Private bRefresh              As Boolean
Private iMaxlblItems          As Integer
Private lCurrListItem         As Long
Private lOldValue             As Long
Private bNoUpdate             As Boolean
Private iMaxInColumn          As Integer
Private lCurrPos              As Long

'Default Property Values:
Const m_def_SelectOnMatch = True
Const m_def_Updated = False
Const m_def_TopIndex = 0
Const m_def_Locked = False
Const m_def_NoDeSelect = False
Const m_def_Sort = 0
Const m_def_AutoSelectItem = False
Const m_def_IntegralHeight = False
Const m_def_RefreshControls = False
Const m_def_ListOpenOnFocus = False
Const m_def_ListHeight = 0
Const vScrollBar_PRESSED = 0
Const m_def_ItemHeight = 0
Const m_def_CopyColorItemsToIcons = False
Const m_def_MousePreSelector = True
Const m_def_ListExitOnSelection = False
Const m_def_LineBorderStyle = 1
Const m_def_ListWidth = 0
Const m_def_SelectedAppearence = 0
Const m_def_RowHeight = 0
Const m_def_ListStyle = 0
Const m_def_Columns = 0
Const m_def_MultiSelect = False
Const m_def_Text = ""
Const m_def_ExactMatch = False
Const m_def_Enabled = True
Const m_def_CausesValidation = 0
Const m_def_Alignment = 0
Const m_def_BackColorPicture = &HC0C0C0
Const m_def_ForeColorUnselected = vbWindowText
Const m_def_BackColorUnselected = vbWindowBackground
Const m_def_ForeColorDisabeled = &H808080
Const m_def_BackColorDisabeled = &HC0C0C0
Const m_def_ForeColorSelected = &HFFFFFF
Const m_def_BackColorSelected = &H0&
Const m_def_SHIFTKEY = 1
Const SLIDERHIGHT = 105
Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move


'Property Variables:
Dim m_SelectOnMatch As Boolean
Dim m_Updated                          As Boolean
Dim m_Locked                           As Boolean
Dim m_NoDeSelect                       As Boolean
Dim m_Sort                             As enumSort
Dim m_AutoSelectItem                   As Boolean
Dim m_ListWidth                        As Integer
Dim m_ItemHeight                       As Integer
Dim m_CopyColorItemsToIcons            As Boolean
Dim m_MousePreSelector                 As Boolean
Dim m_ListExitOnSelection              As Boolean
Dim m_SelectedAppearence               As enumAppearance
Dim m_RowHeight                        As Integer
Dim m_ListStyle                        As enumListStyle
Dim m_Columns                          As Integer
Dim m_MultiSelect                      As Boolean
Dim m_ExactMatch                       As Boolean
Dim m_Enabled                          As Boolean
Dim m_CausesValidation                 As Boolean
Dim m_Alignment                        As Integer
Dim m_FontSelected                     As Font
Dim m_FontDisabeled                    As Font
Dim m_FontUnselected                   As Font
Dim m_BackColorPicture                 As OLE_COLOR
Dim m_ForeColorUnselected              As OLE_COLOR
Dim m_BackColorUnselected              As OLE_COLOR
Dim m_ForeColorDisabeled               As OLE_COLOR
Dim m_BackColorDisabeled               As OLE_COLOR
Dim m_ForeColorSelected                As OLE_COLOR
Dim m_BackColorSelected                As OLE_COLOR

Enum enumAlignMent
   LeftJustify = vbLeftJustify
   RightJustify = vbRightJustify
   Center = vbCenter
End Enum

Enum enumListStyle
   Normal = 0
   CheckBox = 1
   PictureBox = 2
   ImageBox = 3
End Enum

Enum enumAppearance
   Flat = 0
   [3 D] = 1
End Enum


Enum enumSort
   NONE = 0
   Ascending = 1
   Descending = 2
End Enum

Private Type POINTAPI
   X As Long
   Y As Long
End Type

'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, ListItem As String)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event DblClick()
Event Click()
Event Change()
Event PopUpItems(MenuItemsArray() As Variant, ListItem As Long)
Event PopUpItemsClick(MenuIndex As Integer, ListItem As Long)

'***********************************
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ScreenToClient& Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Private Sub ShowInfo(Info As String)
Dim MousePoint As POINTAPI
GetCursorPos MousePoint
With ToolTip
    LblToolTip.Caption = Info
   .Top = (MousePoint.Y + 18) * Screen.TwipsPerPixelY
   .Left = (MousePoint.X - 2) * Screen.TwipsPerPixelX
   .Width = LblToolTip.Width + 4 * Screen.TwipsPerPixelX
   .Visible = True
End With
End Sub

Private Sub hScrollBar_Scroll()
Call ShowInfo(" " & CStr(HScrollBar.Value) & " ")
End Sub

Private Sub imgItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y, lblItems(Index).Tag)
Call MouseMove(Index)
iIndex = Index
End Sub

Private Sub MouseButtonPush(Button As Integer, Index As Integer)
Dim ItemsArray()  As Variant
Dim i             As Integer

If Button = vbRightButton And Not m_Locked Then

   RaiseEvent PopUpItems(ItemsArray, Val(lblItems(Index).Tag))
   If aUbound(ItemsArray) = -1 Then Exit Sub
   
   For i = 1 To mnuPopUpSub.Count - 1 Step 1
      Unload mnuPopUpSub(i)
   Next
   
   For i = 0 To aUbound(ItemsArray) Step 1
      If i > 0 Then Load mnuPopUpSub(i)
      mnuPopUpSub(i).Caption = ItemsArray(i)
   Next
   PopupMenu mnuPopUp

Else
   
   bMouseClick = True
   ItemClick Index

End If

End Sub
Private Sub imgItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Call MouseButtonPush(Button, Index)
End Sub

Private Sub mnuPopUpSub_Click(Index As Integer)
RaiseEvent PopUpItemsClick(Index, Val(lblItems(iIndex).Tag))
End Sub

Private Sub vScrollBar_Scroll()
Call ShowInfo(" " & CStr(VScrollBar.Value) & " ")
End Sub

Private Sub Slider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bHoldSlider = True
End Sub

Private Sub Slider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If bHoldSlider Then SliderMouseDown Index
End Sub

Private Sub Slider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bHoldSlider = False
End Sub

Private Sub UserControl_EnterFocus()
If bRefresh Or bRefreshControls Then Refresh
m_Updated = False
End Sub

Private Sub UserControl_Resize()
ListBox.Move 0, 0, ScaleWidth, ScaleHeight
bRefreshControls = True
End Sub

Private Sub ShowDispName()
If Not Ambient.UserMode Then
   lblItems(0).Visible = True
   lblItems(0) = Ambient.DisplayName
End If
End Sub

Private Sub chkItems_Click(Index As Integer)
RaiseEvent Click
Static iLocked As Integer
If iLocked > 1 Then
   iLocked = 0
   Exit Sub
ElseIf m_Locked Then
   iLocked = picItems(Index).Value + 2
   picItems(Index).Value = Abs(iLocked - 3)
   Exit Sub
End If
End Sub

Private Sub chkItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub chkItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y, lblItems(Index).Tag)
Call MouseMove(Index)
iIndex = Index
End Sub

Private Sub chkItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)

Call MouseButtonPush(Button, Index)
End Sub

Private Sub hScrollBar_Change()

ToolTip.Visible = False
bMouseClick = False

If HScrollBar.Value = lCurrPos + 1 Then
   
   Call UserControl_KeyDown(vbKeyDown, 0)

ElseIf HScrollBar.Value = lCurrPos - 1 Then
   
   Call UserControl_KeyDown(vbKeyUp, 0)

Else
   
   Call Selector(HScrollBar.Value)

End If
End Sub

Private Sub ItemClick(Index As Integer)

Dim l                As Long
Dim bMouseClickTmp   As Boolean

If Not m_Locked And Index > -1 Then
   l = CLng(lblItems(Index).Tag)
   With List
      If .item(l).Enabeled Then
         If Not m_MultiSelect Then
            If lCurrListItem <> l And lCurrListItem > 0 Then
               .item(lCurrListItem).Selected = False
               If iLastPos > -1 Then
                  bMouseClickTmp = bMouseClick
                  bMouseClick = False
                  UpdateView lTopIndex, CLng(lblItems(iLastPos).Tag), iLastPos
                  bMouseClick = bMouseClickTmp
                  If Not m_MousePreSelector Then lblItems(iLastPos).BorderStyle = 0
               End If
            End If
         End If
         If Not m_NoDeSelect Or (m_NoDeSelect And Not .item(l).Selected) Then
            .item(l).Selected = Not .item(l).Selected
         End If
         lCurrPos = l
         UpdateBars lCurrPos
         lCurrListItem = IIf(.item(l).Selected, l, 0)
         UpdateView lTopIndex, l, Index
         If Not m_MousePreSelector Then lblItems(Index).BorderStyle = 0
      End If
   End With
End If

End Sub

Private Sub lblItems_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub lblItems_DblClick(Index As Integer)
RaiseEvent DblClick
End Sub

Private Sub lblItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Function aUbound(aArray As Variant) As Long
aUbound = -1
On Error GoTo EndOnError
aUbound = UBound(aArray)
EndOnError:
End Function


Private Sub UpdateBars(ByVal lPos As Long)
bNoUpdate = True
If VScrollBar.Visible Then
   VScrollBar.Value = Min(Max(lPos, 0), Min(VScrollBar.Max, List.Count))
ElseIf HScrollBar.Visible Then
   HScrollBar.Value = Min(Max(lPos, 0), Min(HScrollBar.Max, List.Count))
End If
bNoUpdate = False
End Sub

Private Sub lblItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y, lblItems(Index).Tag)
Call MouseMove(Index)
iIndex = Index
End Sub

Private Sub lblItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Call MouseButtonPush(Button, Index)
End Sub

Private Sub MouseMove(ByVal Index As Integer)
Dim l As Long
Dim i As Integer

If (m_MousePreSelector Or m_AutoSelectItem) And iIndex <> Index Then

   l = CLng(lblItems(Index).Tag)
   
   With List.item(l)
      If .Enabeled Then
         For i = 0 To lblItems.Count - 1 Step 1
            lblItems(i).BorderStyle = 0
         Next
         lblItems(Index).BorderStyle = 1
         lCurrPos = l
         Call UpdateBars(l)
      End If
   End With
   
   If m_AutoSelectItem Then ItemClick Index

ElseIf Not m_MousePreSelector Then
   If iIndex > -1 And iIndex <= iMaxVisible Then
      lblItems(iIndex).BorderStyle = 0
   End If
End If
End Sub

Private Sub imgItems_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub picItems_Click(Index As Integer)
RaiseEvent Click
Static iLocked As Integer
If iLocked > 1 Then
   iLocked = 0
   Exit Sub
ElseIf m_Locked Then
   iLocked = picItems(Index).Value + 2
   picItems(Index).Value = Abs(iLocked - 3)
   Exit Sub
End If
End Sub

Private Sub picItems_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
If m_Locked Then mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub picItems_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y, lblItems(Index).Tag)
Call MouseMove(Index)
iIndex = Index
End Sub

Private Sub picItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Call MouseButtonPush(Button, Index)
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)

Dim bScroll    As Boolean
Dim i          As Integer

iShift = Shift

Select Case KeyCode
   Case vbKeyUp
      KeyCode = 0
      If bRefresh Or bRefreshControls Then Refresh
      
      If iIndex = 0 Then iIndex = -1
      
      If List.Count = 0 Then Exit Sub
         
      If iIndex > 0 And iShift <> m_def_SHIFTKEY And ListBox.Visible Then
         For i = iIndex - 1 To 0 Step -1
            lblItems(Min(i + 1, iMaxVisible)).BorderStyle = 0
            With List.item(CLng(lblItems(i).Tag))
               If .Enabeled Then
                  lblItems(i).BorderStyle = 1
                  lCurrPos = CLng(lblItems(i).Tag)
                  bScroll = True
                  UpdateBars lCurrPos
                  Exit For
               End If
            End With
         Next
         iIndex = i
      End If
      If Not bScroll And lCurrPos > 0 Then
         Call Selector(Min(lCurrPos - 1, List.Count))
      End If
    Case vbKeyDown
      KeyCode = 0
         
      If bRefresh Or bRefreshControls Then Refresh
      
      If List.Count = 0 Then Exit Sub
      
      If iIndex < iMaxVisible And iShift <> m_def_SHIFTKEY And ListBox.Visible Then
         For i = iIndex + 1 To iMaxVisible Step 1
            lblItems(Max(i - 1, 0)).BorderStyle = 0
            With List.item(CLng(lblItems(i).Tag))
               If .Enabeled Then
                  lblItems(i).BorderStyle = 1
                  lCurrPos = CLng(lblItems(i).Tag)
                  bScroll = True
                  iIndex = i
                  UpdateBars lCurrPos
                  Exit For
               End If
            End With
         Next
      End If
      
      If Not bScroll And lCurrPos < List.Count Then
         Call Selector(Max(lCurrPos + 1, 0))
      End If
   
      
   Case vbKeyPageUp
      KeyCode = 0
      If bRefresh Or bRefreshControls Then Refresh
      
      If List.Count = 0 Then Exit Sub
         
      If lCurrPos > 0 Then
         Call Selector(Min(Max(lCurrPos - iMaxVisible, 0), List.Count))
      End If
   Case vbKeyPageDown
      KeyCode = 0
      If bRefresh Or bRefreshControls Then Refresh
      
      If List.Count = 0 Then Exit Sub
         
      If lCurrPos < List.Count Then
         Call Selector(Min(Max(Max(lCurrPos, 1) + iMaxVisible, 0), List.Count))
      End If
End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)

If KeyAscii = vbKeyReturn Then
   
   KeyAscii = 0
   
   bLastKeyReturn = True
   
   If lCurrPos > 0 And iIndex <= iMaxVisible Then ItemClick iIndex
   
   bLastKeyReturn = False

End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Sub AddItem(ByVal sListItem As String, Optional sKey As String = "", Optional bEnabeled As Boolean = True, Optional vUnderlayingValue As Variant = "", Optional sToolTipText As String = "", Optional oPicture As Picture, Optional oDownPicture As Picture, Optional bVisible As Boolean = True)
Dim l As Long

With List
   If m_Sort Then
      For l = 1 To .Count Step 1
         If m_Sort = Ascending Then
            If sListItem < .item(l).Text Then Exit For
         Else
            If sListItem > .item(l).Text Then Exit For
         End If
      Next
      If .Count = 0 Then
         .Add sListItem, sKey
      ElseIf l > .Count Then
         .Add sListItem, sKey
         l = .Count
      Else
         .Add sListItem, sKey, l
      End If
   Else
      .Add sListItem, sKey
      l = .Count
   End If
   
   iMaxLenItem = Max(iMaxLenItem, Len(sListItem))
   
   .item(l).Enabeled = bEnabeled
   .item(l).UnderlayingValue = vUnderlayingValue
   .item(l).ToolTipText = sToolTipText
   .item(l).ItemPicture = oPicture
   .item(l).DownPicture = oDownPicture
   .item(l).Visible = bVisible

End With

bRefresh = True

End Sub


Private Sub Selector(ByVal lNewVal As Long)
Dim l As Long
Dim i As Integer

If Not bNoUpdate Then
   If lCurrPos < lNewVal Then
      For l = Max(lNewVal, 1) To List.Count Step 1
         With List.item(l)
            If .Enabeled And .Visible Then Exit For
         End With
      Next
   ElseIf lCurrPos > lNewVal Then
      For l = Min(lNewVal, List.Count) To 1 Step -1
         With List.item(l)
            If .Enabeled And .Visible Then Exit For
         End With
      Next
   Else
      l = lNewVal
   End If
   
   If l <= List.Count Then lNewVal = l
    
   If lNewVal > CLng(lblItems(Min(iMaxVisible, List.Count)).Tag) Then
      For l = lNewVal - 1 To 1 Step -1
         If List.item(l).Visible Then i = i + 1
         If i = iMaxlblItems Then Exit For
      Next
      lTopIndex = l
   ElseIf lNewVal <= lTopIndex Then
      For l = lNewVal - 1 To 1 Step -1
         If List.item(l).Visible Then Exit For
      Next
      lTopIndex = l
   End If
      
   With List
      If Not ListBox.Visible And Not m_Locked Then
         If lCurrListItem > 0 Then
            If Not m_MultiSelect And .item(lCurrListItem).Selected Then
               If Not m_NoDeSelect Or (m_NoDeSelect And lNewVal > 0) Then
                  .item(lCurrListItem).Selected = False
               Else
                  lNewVal = lCurrListItem
               End If
            End If
         End If
         If lNewVal > 0 Then .item(lNewVal).Selected = True
         lCurrListItem = lNewVal
      ElseIf Not m_Locked Then
         If iShift = m_def_SHIFTKEY Then
            If lCurrPos > 0 Then
               If .item(lCurrPos).Enabeled Then
                  .item(lCurrPos).Selected = Not .item(lCurrPos).Selected
                  If .item(lCurrPos).Selected Then
                     If lCurrListItem > 0 And Not m_MultiSelect Then
                        .item(lCurrListItem).Selected = False
                     End If
                     lCurrListItem = lCurrPos
                  Else
                     If m_NoDeSelect And lCurrPos = lCurrListItem Then
                        .item(lCurrPos).Selected = True
                     Else
                        lCurrListItem = 0
                     End If
                  End If
               End If
            End If
         End If
      End If
      lCurrPos = lNewVal
      UpdateBars lNewVal
      iShift = 0
   End With
   UpdateView lTopIndex
Else
   bNoUpdate = False
End If
End Sub

Private Sub vScrollBar_Change()
bMouseClick = False
ToolTip.Visible = False
If VScrollBar.Value = lCurrPos + 1 Then
   Call UserControl_KeyDown(vbKeyDown, 0)
ElseIf VScrollBar.Value = lCurrPos - 1 Then
   Call UserControl_KeyDown(vbKeyUp, 0)
Else
   Call Selector(VScrollBar.Value)
End If
End Sub

Public Function HwndListBox()
HwndListBox = ListBox.hwnd
End Function

Public Function HwndFlexBox()
HwndFlexBox = UserControl.hwnd
End Function

Private Sub UpdateView(ByVal lNewVal As Long, Optional lItem As Long = -1, Optional iPos As Integer)
Dim i       As Integer
Dim iTmp    As Integer
Dim l       As Long
Dim sTmp    As String

If List.Count = 0 Or iMaxlblItems = 0 Then Exit Sub

lTopIndex = Max(Min(lNewVal, Max(List.Count - iMaxlblItems, 0)), 0)

i = iPos

For l = IIf(lItem > -1, lItem, lTopIndex + 1) To IIf(lItem > -1, lItem, List.Count) Step 1
   With List.item(l)
      If .Visible Then
            lblItems(i).Tag = CStr(l)
         If Not .Enabeled Then
            With lblItems(i)
               .Appearance = 0
               .BorderStyle = 0
               .ForeColor = m_ForeColorDisabeled
               .BackColor = m_BackColorDisabeled
               Set .Font = m_FontDisabeled
            End With
            If m_ListStyle = CheckBox Then
               chkItems(i).Value = IIf(.Selected, 1, 0)
               chkItems(i).Enabled = False
               chkItems(i).BackColor = m_BackColorDisabeled
            ElseIf m_ListStyle = PictureBox Then
               If Not .ItemPicture Is Nothing Then
                  picItems(i).Value = IIf(.Selected, 1, 0)
                  picItems(i).Enabled = False
                  If m_CopyColorItemsToIcons Then
                     picItems(i).BackColor = m_BackColorDisabeled
                  End If
               End If
            ElseIf m_ListStyle = ImageBox Then
               If .Selected And Not .DownPicture Is Nothing Then
                  imgItems(i).Picture = .DownPicture
               Else
                  imgItems(i).Picture = .ItemPicture
               End If
               imgItems(i).Enabled = False
            End If
         ElseIf .Selected Then
            With lblItems(i)
               If lCurrPos = l Then
                  .BorderStyle = 1
                  iIndex = i
               Else
                  .BorderStyle = 0
               End If
               .Appearance = m_SelectedAppearence
               .ForeColor = m_ForeColorSelected
               .BackColor = m_BackColorSelected
               Set .Font = m_FontSelected
               If lCurrListItem = l Then iLastPos = i
            End With
            If m_ListStyle = CheckBox Then
               chkItems(i).Enabled = True
               chkItems(i).Value = 1
               chkItems(i).BackColor = m_BackColorSelected
            ElseIf m_ListStyle = PictureBox Then
               picItems(i).Enabled = True
               picItems(i).Value = 1
               If m_CopyColorItemsToIcons Then
                  picItems(i).BackColor = m_BackColorSelected
               End If
            ElseIf m_ListStyle = ImageBox Then
               imgItems(i).Enabled = True
               If Not .DownPicture Is Nothing Then
                  imgItems(i).Picture = .DownPicture
               Else
                  imgItems(i).Picture = .ItemPicture
               End If
            End If
         Else  'unselected
            With lblItems(i)
               If lCurrPos = l Then
                  .BorderStyle = 1
                  iIndex = i
               Else
                  .BorderStyle = 0
               End If
               .Appearance = 0
               .ForeColor = m_ForeColorUnselected
               .BackColor = m_BackColorUnselected
               Set .Font = m_FontUnselected
            End With
            If m_ListStyle = CheckBox Then
               chkItems(i).Enabled = True
               chkItems(i).Value = 0
               chkItems(i).BackColor = m_BackColorUnselected
            ElseIf m_ListStyle = PictureBox Then
               picItems(i).Enabled = True
               picItems(i).Value = 0
               If m_CopyColorItemsToIcons Then
                  picItems(i).BackColor = m_BackColorUnselected
               End If
            ElseIf m_ListStyle = ImageBox Then
               imgItems(i).Picture = .ItemPicture
            End If
         End If
         If m_ListStyle = PictureBox Then
            If .ItemPicture Is Nothing Then
               picItems(i).Visible = False
            Else
               picItems(i).Visible = True
               picItems(i).Picture = .ItemPicture
               picItems(i).DownPicture = .DownPicture
               picItems(i).DisabledPicture = Nothing
            End If
         ElseIf m_ListStyle = ImageBox Then
            imgItems(i).Visible = .ItemPicture Is Nothing = False
         ElseIf m_ListStyle = CheckBox Then
            chkItems(i).Visible = True
         End If
         lblItems(i).Visible = True
         If m_ListStyle = ImageBox Then
            sTmp = .Text
            If sTmp <> "" Then
               Set ListBox.Font = lblItems(i).Font
               iTmp = ListBox.TextWidth(sTmp) + 300
               Do While ListBox.TextWidth(sTmp) <= iTmp
                  sTmp = " " & sTmp
               Loop
            End If
            lblItems(i).Caption = sTmp
         Else
            lblItems(i).Caption = .Text
         End If
         lblItems(i).ToolTipText = .ToolTipText
         i = i + 1
         If i = iMaxlblItems Then Exit For
      End If
   End With
Next
  
If i < iMaxlblItems And lItem = -1 Then
   iMaxVisible = i - 1
   For i = i To iMaxlblItems - 1 Step 1
      lblItems(i).Visible = False
      If m_ListStyle = CheckBox Then
         chkItems(i).Visible = False
      ElseIf m_ListStyle = PictureBox Then
         picItems(i).Visible = False
      ElseIf m_ListStyle = ImageBox Then
         imgItems(i).Visible = False
      End If
   Next
ElseIf lItem = -1 Then
   iMaxVisible = iMaxlblItems - 1
End If


If lOldValue <> lCurrListItem Then
   lOldValue = lCurrListItem
   m_Updated = True
   RaiseEvent Change
   If m_ListExitOnSelection And lCurrListItem > 0 And (bLastKeyReturn Or bMouseClick) Then
      SendKeys "{tab}"
   End If
End If
bMouseClick = False
End Sub


Private Sub LoadControls()
Dim i          As Integer
Dim i2         As Integer
Dim i3         As Integer
Dim iImagePos  As Integer
                        
iIndex = -1
iLastPos = -1

lTopIndex = 0
iMaxVisible = 0
iMaxlblItems = 0
iMaxInColumn = 0
                        
Call UnLoadControls

chkItems(0).Visible = (m_ListStyle = CheckBox) And List.Count > 0
picItems(0).Visible = (m_ListStyle = PictureBox) And List.Count > 0
imgItems(0).Visible = (m_ListStyle = ImageBox) And List.Count > 0
Slider(0).Visible = False

ListBox.BackColor = m_BackColorUnselected
picItems(0).BackColor = m_BackColorPicture

With lblItems(0)
   Set .Font = m_FontUnselected
   .Caption = ""
   .Visible = False
   .BorderStyle = 0
   If m_ItemHeight = 0 Then
      .Height = 0
      For i = 1 To 3 Step 1
         Set ListBox.Font = Choose(i, m_FontUnselected, m_FontSelected, m_FontDisabeled)
         .Height = Max(.Height, Int(ListBox.TextHeight("X")))
      Next
   Else
      .Height = Int(Min(Max(m_ItemHeight, 100), ListBox.Height * 0.3))
   End If

   If List.Count = 0 Then Exit Sub
   
   If m_Columns > 1 Then
      .Top = SLIDERHIGHT
      picItems(0).Top = SLIDERHIGHT
      imgItems(0).Top = SLIDERHIGHT
      chkItems(0).Top = SLIDERHIGHT
      SliderLine1.Visible = True
      SliderLine2.Visible = True
      SliderLine1.X2 = ListBox.Width
      SliderLine2.X2 = ListBox.Width
   Else
      .Top = 0
      picItems(0).Top = 0
      imgItems(0).Top = 0
      chkItems(0).Top = 0
      SliderLine1.Visible = False
      SliderLine2.Visible = False
   End If
   
   If m_ListStyle = CheckBox Then
      chkItems(0).Visible = True
      chkItems(0).BackColor = m_BackColorUnselected
      chkItems(0).Height = .Height
      chkItems(0).Width = 200
   ElseIf m_ListStyle = PictureBox Then
      picItems(0).Visible = True
      picItems(0).Height = .Height
      picItems(0).Width = .Height
   ElseIf m_ListStyle = ImageBox Then
      imgItems(0).Visible = True
      imgItems(0).Height = Min(.Height, 200)
      iImagePos = Int((.Height - imgItems(0).Height) * 0.5)
      imgItems(0).ZOrder
   End If
   
   .Left = IIf(m_ListStyle = Normal Or m_ListStyle = ImageBox, 20, IIf(m_ListStyle = PictureBox, .Height + 30, 230))
   
   iMaxInColumn = Int((ListBox.Height - IIf(m_Columns > 0, 180 + IIf(m_Columns > 1, SLIDERHIGHT, 0) + IIf(ListBox.Appearance, 20, 0), 0)) / .Height)
   
   If m_Columns > 0 Then
      iMaxlblItems = iMaxInColumn * m_Columns
      .Width = ((ListBox.Width - 40) - m_Columns * .Left) / m_Columns
   Else
      iMaxlblItems = Int(ListBox.Height / .Height)
   End If
   
   'If m_ListStyle = ImageBox Then
   '   imgItems(0).Top = .Top + Int((.Height - 250) * 0.5)
   'End If
   
   i2 = 1
   
   For i = 1 To iMaxlblItems - 1 Step 1
      Load lblItems(i)
      lblItems(i).Move Int(.Left + (.Width + .Left + 25) * i3), .Top + (.Height * i2)
      If i3 + 1 = m_Columns And m_Columns > 1 Then
         lblItems(i).Width = .Width - 80
      End If
      If m_ListStyle = CheckBox Then
         Load chkItems(i)
         chkItems(i).Move (.Left + .Width + 26) * i3 + 20, .Top + (.Height * i2)
      ElseIf m_ListStyle = PictureBox Then
         Load picItems(i)
         picItems(i).Move (.Left + .Width + 30) * i3, .Top + (.Height * i2)
      ElseIf m_ListStyle = ImageBox Then
         Load imgItems(i)
         imgItems(i).Move lblItems(i).Left + 50, lblItems(i).Top + iImagePos
         imgItems(i).ZOrder
      End If
      i2 = i2 + 1
      If i2 = iMaxInColumn And m_Columns > 0 Then
         If i < iMaxlblItems - 1 Then
            If i3 > 0 Then Load Slider(i3)
            Slider(i3).Move lblItems(i).Left + lblItems(i).Width
            Slider(i3).Visible = True
         End If
         i2 = 0
         i3 = i3 + 1
      End If
   Next
End With
bRefreshControls = False
End Sub


Private Sub UserControl_Terminate()
Set List = Nothing
Call UnLoadControls
'This is the "dongle" for an commercial version
Call ShowSplash
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("ForeColorUnselected", m_ForeColorUnselected, m_def_ForeColorUnselected)
Call PropBag.WriteProperty("BackColorUnselected", m_BackColorUnselected, m_def_BackColorUnselected)
Call PropBag.WriteProperty("ForeColorDisabeled", m_ForeColorDisabeled, m_def_ForeColorDisabeled)
Call PropBag.WriteProperty("BackColorDisabeled", m_BackColorDisabeled, m_def_BackColorDisabeled)
Call PropBag.WriteProperty("ForeColorSelected", m_ForeColorSelected, m_def_ForeColorSelected)
Call PropBag.WriteProperty("BackColorSelected", m_BackColorSelected, m_def_BackColorSelected)
Call PropBag.WriteProperty("FontDisabeled", m_FontDisabeled, UserControl.Font)
Call PropBag.WriteProperty("FontUnselected", m_FontUnselected, Slider(0).Font)
Call PropBag.WriteProperty("FontSelected", m_FontSelected, ListBox.Font)
Call PropBag.WriteProperty("Appearance", ListBox.Appearance, 0)
Call PropBag.WriteProperty("CausesValidation", m_CausesValidation, m_def_CausesValidation)
Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
Call PropBag.WriteProperty("ExactMatch", m_ExactMatch, m_def_ExactMatch)
Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
Call PropBag.WriteProperty("Columns", m_Columns, m_def_Columns)
Call PropBag.WriteProperty("ListStyle", m_ListStyle, m_def_ListStyle)
Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, m_def_ItemHeight)
Call PropBag.WriteProperty("BackColorPicture", m_BackColorPicture, m_def_BackColorPicture)
Call PropBag.WriteProperty("SelectedAppearence", m_SelectedAppearence, m_def_SelectedAppearence)
Call PropBag.WriteProperty("ListExitOnSelection", m_ListExitOnSelection, m_def_ListExitOnSelection)
Call PropBag.WriteProperty("MousePreSelector", m_MousePreSelector, m_def_MousePreSelector)
Call PropBag.WriteProperty("CopyColorItemsToIcons", m_CopyColorItemsToIcons, m_def_CopyColorItemsToIcons)
Call PropBag.WriteProperty("ListWidth", m_ListWidth, m_def_ListWidth)
Call PropBag.WriteProperty("AutoSelectItem", m_AutoSelectItem, m_def_AutoSelectItem)
Call PropBag.WriteProperty("Sort", m_Sort, m_def_Sort)
Call PropBag.WriteProperty("NoDeSelect", m_NoDeSelect, m_def_NoDeSelect)
Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
Call PropBag.WriteProperty("MousePointer", ListBox.MousePointer, 0)
Call PropBag.WriteProperty("TopIndex", lTopIndex, m_def_TopIndex)
Call PropBag.WriteProperty("SelectOnMatch", m_SelectOnMatch, m_def_SelectOnMatch)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_ForeColorUnselected = PropBag.ReadProperty("ForeColorUnselected", m_def_ForeColorUnselected)
m_BackColorUnselected = PropBag.ReadProperty("BackColorUnselected", m_def_BackColorUnselected)
m_ForeColorDisabeled = PropBag.ReadProperty("ForeColorDisabeled", m_def_ForeColorDisabeled)
m_BackColorDisabeled = PropBag.ReadProperty("BackColorDisabeled", m_def_BackColorDisabeled)
m_ForeColorSelected = PropBag.ReadProperty("ForeColorSelected", m_def_ForeColorSelected)
m_BackColorSelected = PropBag.ReadProperty("BackColorSelected", m_def_BackColorSelected)
Set lblItems(0).Font = PropBag.ReadProperty("FontSelected", lblItems(0).Font)
Set m_FontDisabeled = PropBag.ReadProperty("FontDisabeled", UserControl.Font)
Set m_FontUnselected = PropBag.ReadProperty("FontUnselected", Slider(0).Font)
Set m_FontSelected = PropBag.ReadProperty("FontSelected", ListBox.Font)
ListBox.Appearance = PropBag.ReadProperty("Appearance", 0)
m_CausesValidation = PropBag.ReadProperty("CausesValidation", m_def_CausesValidation)
m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
m_ExactMatch = PropBag.ReadProperty("ExactMatch", m_def_ExactMatch)
m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
m_Columns = PropBag.ReadProperty("Columns", m_def_Columns)
m_ListStyle = PropBag.ReadProperty("ListStyle", m_def_ListStyle)
m_ItemHeight = PropBag.ReadProperty("ItemHeight", m_def_ItemHeight)
m_BackColorPicture = PropBag.ReadProperty("BackColorPicture", m_def_BackColorPicture)
m_SelectedAppearence = PropBag.ReadProperty("SelectedAppearence", m_def_SelectedAppearence)
m_ListExitOnSelection = PropBag.ReadProperty("ListExitOnSelection", m_def_ListExitOnSelection)
m_MousePreSelector = PropBag.ReadProperty("MousePreSelector", m_def_MousePreSelector)
m_CopyColorItemsToIcons = PropBag.ReadProperty("CopyColorItemsToIcons", m_def_CopyColorItemsToIcons)

m_ListWidth = PropBag.ReadProperty("ListWidth", m_def_ListWidth)

m_AutoSelectItem = PropBag.ReadProperty("AutoSelectItem", m_def_AutoSelectItem)
m_Sort = PropBag.ReadProperty("Sort", m_def_Sort)
m_NoDeSelect = PropBag.ReadProperty("NoDeSelect", m_def_NoDeSelect)
m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
ListBox.MousePointer = PropBag.ReadProperty("MousePointer", 0)
lTopIndex = PropBag.ReadProperty("TopIndex", m_def_TopIndex)
m_SelectOnMatch = PropBag.ReadProperty("SelectOnMatch", m_def_SelectOnMatch)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
m_BackColorUnselected = m_def_BackColorUnselected
m_ForeColorDisabeled = m_def_ForeColorDisabeled
m_BackColorDisabeled = m_def_BackColorDisabeled
m_ForeColorSelected = m_def_ForeColorSelected
m_BackColorSelected = m_def_BackColorSelected
m_ForeColorUnselected = m_def_ForeColorUnselected

Set m_FontSelected = ListBox.Font
Set m_FontDisabeled = UserControl.Font
Set m_FontUnselected = Slider(0).Font

m_Updated = m_def_Updated
m_CausesValidation = m_def_CausesValidation
m_Alignment = m_def_Alignment
m_Enabled = m_def_Enabled
m_ExactMatch = m_def_ExactMatch
m_MultiSelect = m_def_MultiSelect
m_Columns = m_def_Columns
m_ListStyle = m_def_ListStyle
m_ItemHeight = m_def_ItemHeight
m_BackColorPicture = m_def_BackColorPicture
m_SelectedAppearence = m_def_SelectedAppearence
m_ListExitOnSelection = m_def_ListExitOnSelection
m_MousePreSelector = m_def_MousePreSelector
m_CopyColorItemsToIcons = m_def_CopyColorItemsToIcons

m_ListWidth = m_def_ListWidth
m_AutoSelectItem = m_def_AutoSelectItem
m_Sort = m_def_Sort
m_NoDeSelect = m_def_NoDeSelect
m_Locked = m_def_Locked
lTopIndex = m_def_TopIndex
m_SelectOnMatch = m_def_SelectOnMatch
End Sub

Private Function Min(ByVal X As Long, ByVal Y As Long) As Long
Min = IIf(X < Y, X, Y)
End Function

Private Function Max(ByVal X As Long, ByVal Y As Long) As Long
Max = IIf(X > Y, X, Y)
End Function

Public Sub Refresh()
Dim i As Integer

If bRefresh Or bRefreshControls Then

   bRefresh = False
   
   ListBox.Enabled = m_Enabled
   
   If bRefreshControls Then Call LoadControls
   
   VScrollBar.Visible = False
   HScrollBar.Visible = False
   
   
   If List.Count = 0 Then
      lblItems(0).Visible = False
      chkItems(0).Visible = False
      picItems(0).Visible = False
      imgItems(0).Visible = False
      Exit Sub
   End If
   
   
   If m_Columns = 0 Then
      With VScrollBar
         .Height = ListBox.Height - IIf(ListBox.Appearance, 50, 20)
         .Left = ListBox.Width - (.Width + IIf(ListBox.Appearance, 30, 0))
         .Max = List.Count
         .LargeChange = Int(List.Count * 0.1 + 1)
         .Visible = List.Count - iMaxlblItems > 0 Or Not Ambient.UserMode
      End With
   Else
      With HScrollBar
         .Width = ListBox.Width - IIf(ListBox.Appearance = 1, 45, 30)
         .Top = ListBox.Height - (180 + IIf(ListBox.Appearance = 0, 20, 40))
         .Max = List.Count
         .SmallChange = 1
         .LargeChange = iMaxInColumn
         .Visible = List.Count - iMaxlblItems > 0 Or Not Ambient.UserMode
      End With
   End If
   
   For i = 0 To iMaxlblItems - 1 Step 1
      With lblItems(i)
         If m_Columns = 0 Then
            .Width = ListBox.Width - (IIf(m_ListStyle = PictureBox, .Height + 20, IIf(m_ListStyle = CheckBox, 220, 0)) + IIf(List.Count > iMaxlblItems, 260, 50) + IIf(ListBox.Appearance, 40, 0))
         End If
         .Alignment = m_Alignment
         .Enabled = m_Enabled
      End With
      If List.Count > i Then
         If m_ListStyle = CheckBox Then
            chkItems(i).Visible = True
         ElseIf m_ListStyle = PictureBox Then
            picItems(i).Visible = True
         End If
      End If
   Next

End If

UpdateView lTopIndex

End Sub

Public Sub RemoveItem(vItem As Variant)
Dim l As Long

l = aScan(List.item(vItem).Text)

If l > 0 And l < lCurrListItem Then
   lCurrListItem = lCurrListItem - 1
ElseIf l = lCurrListItem Then
   lCurrListItem = 0
End If

List.Remove vItem

If lCurrListItem > 0 Then
   If Not List.item(lCurrListItem).Enabeled Then lCurrListItem = 0
End If

bRefresh = True

End Sub

Public Sub Clear()
Set List = Nothing
Set List = New ListCollection
lCurrListItem = 0
lCurrPos = 0
bRefreshControls = True
Refresh
End Sub

Public Function aScan(ByVal sWhat2Scan As String, Optional lStartPos As Long = 1) As Long
Dim l          As Long
Dim iLenStr    As String

With List
   If m_ExactMatch Then
      For l = lStartPos To .Count Step 1
         If .item(l).Text = sWhat2Scan Then
            aScan = l
            Exit Function
         End If
      Next
   Else
      If sWhat2Scan = "" Then
         For l = lStartPos To .Count Step 1
            If .item(l).Text = "" Then
               aScan = l
               Exit Function
            End If
         Next
      Else
         iLenStr = Len(sWhat2Scan)
         sWhat2Scan = UCase(sWhat2Scan)
         For l = lStartPos To .Count Step 1
            If Left(UCase(.item(l).Text), iLenStr) = sWhat2Scan Then
               aScan = l
               Exit Function
            End If
         Next
      End If
   End If
End With
End Function

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = "Characteristics"
MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
m_MultiSelect = New_MultiSelect
PropertyChanged "MultiSelect"
End Property

Public Property Get Columns() As Integer
Attribute Columns.VB_ProcData.VB_Invoke_Property = "Characteristics"
Columns = m_Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
Dim iTmp As Integer

iTmp = Max(New_Columns, 0)

If New_Columns > 0 Then
   
   If (CLng(New_Columns) * 600) > ListBox.Width Then
      iTmp = Int(ListBox.Width / 600)
   End If

End If

If iTmp <> m_Columns Then
   m_Columns = iTmp
   bRefreshControls = True
   Refresh
End If

PropertyChanged "Columns"
End Property

Public Property Get ListStyle() As enumListStyle
ListStyle = m_ListStyle
End Property

Public Property Let ListStyle(ByVal New_ListStyle As enumListStyle)
If m_ListStyle <> New_ListStyle Then
   m_ListStyle = New_ListStyle
   bRefreshControls = True
   Refresh
   PropertyChanged "ListStyle"
End If
End Property

Public Property Get ItemHeight() As Integer
Attribute ItemHeight.VB_ProcData.VB_Invoke_Property = "Characteristics"
ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)
Dim iTmp As Integer

iTmp = IIf(New_ItemHeight > 0, Min(Max(New_ItemHeight, 100), 800), 0)

If iTmp <> m_ItemHeight Then
   m_ItemHeight = iTmp
   bRefreshControls = True
   Refresh
   PropertyChanged "ItemHeight"
End If

End Property

Public Property Get BackColorUnselected() As OLE_COLOR
BackColorUnselected = m_BackColorUnselected
End Property

Public Property Let BackColorUnselected(ByVal New_BackColorUnselected As OLE_COLOR)
m_BackColorUnselected = New_BackColorUnselected
ListBox.BackColor = m_BackColorUnselected
PropertyChanged "BackColorUnselected"
End Property

Public Property Get ForeColorDisabeled() As OLE_COLOR
ForeColorDisabeled = m_ForeColorDisabeled
End Property

Public Property Let ForeColorDisabeled(ByVal New_ForeColorDisabeled As OLE_COLOR)
m_ForeColorDisabeled = New_ForeColorDisabeled
PropertyChanged "ForeColorDisabeled"
End Property

Public Property Get BackColorDisabeled() As OLE_COLOR
BackColorDisabeled = m_BackColorDisabeled
End Property

Public Property Let BackColorDisabeled(ByVal New_BackColorDisabeled As OLE_COLOR)
m_BackColorDisabeled = New_BackColorDisabeled
PropertyChanged "BackColorDisabeled"
End Property

Public Property Get ForeColorSelected() As OLE_COLOR
ForeColorSelected = m_ForeColorSelected
End Property

Public Property Let ForeColorSelected(ByVal New_ForeColorSelected As OLE_COLOR)
m_ForeColorSelected = New_ForeColorSelected
PropertyChanged "ForeColorSelected"
End Property

Public Property Get BackColorSelected() As OLE_COLOR
BackColorSelected = m_BackColorSelected
End Property

Public Property Let BackColorSelected(ByVal New_BackColorSelected As OLE_COLOR)
m_BackColorSelected = New_BackColorSelected
PropertyChanged "BackColorSelected"
End Property

Public Property Get FontDisabeled() As Font
Set FontDisabeled = m_FontDisabeled
End Property

Public Property Set FontDisabeled(ByVal New_FontDisabeled As Font)
Set m_FontDisabeled = New_FontDisabeled
bRefreshControls = True
PropertyChanged "FontDisabeled"
End Property

Public Property Get FontUnselected() As Font
Set FontUnselected = m_FontUnselected
End Property

Public Property Set FontUnselected(ByVal New_FontUnselected As Font)
Set m_FontUnselected = New_FontUnselected
bRefreshControls = True
PropertyChanged "FontUnselected"
End Property

Public Property Get Appearance() As enumAppearance
Appearance = ListBox.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As enumAppearance)
ListBox.Appearance() = New_Appearance
PropertyChanged "Appearance"
End Property

Public Property Get ListIndex() As Variant
ListIndex = lCurrListItem - 1
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Variant)
If Not m_MultiSelect And lCurrListItem > 0 Then
   List.item(lCurrListItem).Selected = False
End If
If VarType(New_ListIndex) = vbString Then 'seek on key
   lCurrListItem = aScan(List.item(New_ListIndex).Text)
Else
   New_ListIndex = New_ListIndex + 1
   lCurrListItem = New_ListIndex 'numeric
End If
If lCurrListItem > 0 Then
   If List.item(New_ListIndex).Enabeled And List.item(New_ListIndex).Visible Then
      lCurrListItem = New_ListIndex
      List.item(lCurrListItem).Selected = True
   Else
      lCurrListItem = 0
   End If
End If
If VScrollBar.Visible Then
   UpdateBars Max(Min(lCurrListItem - 1, VScrollBar.Max), 0)
ElseIf HScrollBar.Visible Then
   UpdateBars Max(Min(lCurrListItem - 1, HScrollBar.Max), 0)
End If
UpdateView lCurrListItem - 1
End Property

Public Property Get FontSelected() As Font
Set FontSelected = m_FontSelected
End Property

Public Property Set FontSelected(ByVal New_FontSelected As Font)
Set m_FontSelected = New_FontSelected
bRefreshControls = True
PropertyChanged "FontSelected"
End Property

Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_ProcData.VB_Invoke_Property = "Characteristics"
CausesValidation = m_CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
m_CausesValidation = New_CausesValidation
PropertyChanged "CausesValidation"
End Property

Public Property Get Alignment() As enumAlignMent
Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As enumAlignMent)
m_Alignment = New_Alignment
bRefresh = True
PropertyChanged "Alignment"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "Characteristics"
Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
m_Enabled = New_Enabled
bRefresh = True
Refresh
PropertyChanged "Enabled"
End Property

Public Property Get ExactMatch() As Boolean
Attribute ExactMatch.VB_ProcData.VB_Invoke_Property = "Characteristics"
ExactMatch = m_ExactMatch
End Property

Public Property Let ExactMatch(ByVal New_ExactMatch As Boolean)
m_ExactMatch = New_ExactMatch
PropertyChanged "ExactMatch"
End Property

Public Property Get ForeColorUnselected() As OLE_COLOR
ForeColorUnselected = m_ForeColorUnselected
End Property

Public Property Let ForeColorUnselected(ByVal New_ForeColorUnselected As OLE_COLOR)
m_ForeColorUnselected = New_ForeColorUnselected
PropertyChanged "ForeColorUnselected"
End Property

Public Sub SoftSeek(ByVal New_Text As String)
Dim lTmp As Long
lTmp = aScan(New_Text)
If lTmp = 0 Then Exit Sub
lCurrPos = lTmp
Do While lCurrPos > 0
   If Not List.item(lCurrPos).Enabeled Then
      lCurrPos = aScan(New_Text, lCurrPos + 1)
   Else
      Exit Do
   End If
Loop
If List.Count > iMaxlblItems Then
   If m_Columns = 0 Then
      VScrollBar.Value = lCurrPos
   Else
      HScrollBar.Value = lCurrPos
   End If
End If
UpdateView lCurrPos - 1
End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "1234"
If lCurrListItem > 0 Then Text = List.item(lCurrListItem).Text
End Property

Public Property Let Text(ByVal New_Text As String)
Dim l As Long

If m_SelectOnMatch And lCurrListItem > 0 And Not m_MultiSelect Then
   List.item(lCurrListItem).Selected = False
End If

l = aScan(New_Text)

Do While l > 0
   
   If Not List.item(l).Enabeled Then
      l = aScan(New_Text, l + 1)
   Else
      Exit Do
   End If

Loop

lCurrPos = l

If List.Count > iMaxlblItems Then
   If m_Columns = 0 Then
      VScrollBar.Value = lCurrPos
   Else
      HScrollBar.Value = lCurrPos
   End If
End If

If l > 0 And m_SelectOnMatch Then
   List.item(l).Selected = True
   lCurrListItem = l

End If
If ListBox.Visible Then UpdateView l - 1

End Property

Public Property Get BackColorPicture() As OLE_COLOR
BackColorPicture = m_BackColorPicture
End Property

Public Property Let BackColorPicture(ByVal New_BackColorPicture As OLE_COLOR)
m_BackColorPicture = New_BackColorPicture
PropertyChanged "BackColorPicture"
End Property

Public Property Get SelectedAppearence() As enumAppearance
SelectedAppearence = m_SelectedAppearence
End Property

Public Property Let SelectedAppearence(ByVal New_SelectedAppearence As enumAppearance)
m_SelectedAppearence = New_SelectedAppearence
bRefresh = True
PropertyChanged "SelectedAppearence"
End Property

Public Property Get ListExitOnSelection() As Boolean
Attribute ListExitOnSelection.VB_ProcData.VB_Invoke_Property = "Characteristics"
ListExitOnSelection = m_ListExitOnSelection
End Property

Public Property Let ListExitOnSelection(ByVal New_ListExitOnSelection As Boolean)
m_ListExitOnSelection = New_ListExitOnSelection
PropertyChanged "ListExitOnSelection"
End Property

Public Property Get MousePreSelector() As Boolean
Attribute MousePreSelector.VB_ProcData.VB_Invoke_Property = "Characteristics"
MousePreSelector = m_MousePreSelector
End Property

Public Property Let MousePreSelector(ByVal New_MousePreSelector As Boolean)
m_MousePreSelector = New_MousePreSelector
PropertyChanged "MousePreSelector"
End Property

Public Property Get CopyColorItemsToIcons() As Boolean
Attribute CopyColorItemsToIcons.VB_ProcData.VB_Invoke_Property = "Characteristics"
CopyColorItemsToIcons = m_CopyColorItemsToIcons
End Property

Public Property Let CopyColorItemsToIcons(ByVal New_CopyColorItemsToIcons As Boolean)
m_CopyColorItemsToIcons = New_CopyColorItemsToIcons
PropertyChanged "CopyColorItemsToIcons"
End Property

Public Sub RefreshControls()
bRefreshControls = True
Call Refresh
End Sub

Public Property Get AutoSelectItem() As Boolean
Attribute AutoSelectItem.VB_ProcData.VB_Invoke_Property = "Characteristics"
AutoSelectItem = m_AutoSelectItem
End Property

Public Property Let AutoSelectItem(ByVal New_AutoSelectItem As Boolean)
m_AutoSelectItem = New_AutoSelectItem
PropertyChanged "AutoSelectItem"
End Property

Public Property Get Sort() As enumSort
Sort = m_Sort
End Property

Public Property Let Sort(ByVal New_Sort As enumSort)
m_Sort = New_Sort
PropertyChanged "Sort"
End Property

Public Property Get NoDeSelect() As Boolean
Attribute NoDeSelect.VB_ProcData.VB_Invoke_Property = "Characteristics"
NoDeSelect = m_NoDeSelect
End Property

Public Property Let NoDeSelect(ByVal New_NoDeSelect As Boolean)
m_NoDeSelect = New_NoDeSelect
PropertyChanged "NoDeSelect"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = "Characteristics"
Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
m_Locked = New_Locked
PropertyChanged "Locked"
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
ListBox.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As Integer
MousePointer = ListBox.MousePointer
End Property

Public Property Get MouseIcon() As Picture
Set MouseIcon = ListBox.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
Set ListBox.MouseIcon = New_MouseIcon
PropertyChanged "MouseIcon"
End Property

Public Property Get TopIndex() As Long
TopIndex = lTopIndex
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Long)
lTopIndex = Max(Min(New_TopIndex, List.Count), 0)
lCurrPos = lTopIndex + 1
If Ambient.UserMode Then
   UpdateBars lTopIndex
   UpdateView lTopIndex
End If
PropertyChanged "TopIndex"
End Property


Private Sub SliderMouseDown(Index As Integer)
Dim MousePoint    As POINTAPI
Dim iMinimum      As Integer
Dim iMaximum      As Integer
Dim iTmp          As Integer
Dim itmp2         As Integer
Dim i             As Integer

If List.Count = 0 Then Exit Sub

If Index = 0 Then
   iMinimum = 0
Else
   iMinimum = Slider(Index - 1).Left + 60
End If

If Index = Slider.Count - 1 Then
   iMaximum = ListBox.Width - 100
Else
   iMaximum = Slider(Index + 1).Left - 70
End If

GetCursorPos MousePoint
Call ScreenToClient(ListBox.hwnd, MousePoint)

Slider(Index).Move Min(Max(MousePoint.X * Screen.TwipsPerPixelX - 25, iMinimum), iMaximum)

If Index = 0 Then
   iMinimum = 0
Else
   iMinimum = Index * iMaxInColumn
End If

If Index = Slider.Count - 1 Then
   iMaximum = iMaxlblItems - 1
Else
   iMaximum = (Index + 2) * iMaxInColumn
End If

If Index > 0 Then
   iTmp = Slider(Index).Left - Slider(Index - 1).Left
Else
   iTmp = Slider(Index).Left
End If

For i = iMinimum To (iMinimum + iMaxInColumn) - 1 Step 1
   Select Case m_ListStyle
      Case PictureBox
         picItems(i).Width = Min(iTmp, lblItems(0).Height)
         lblItems(i).Width = Max(0, (iTmp - picItems(i).Width - 50))

      Case ImageBox
         imgItems(i).Width = Min(iTmp, 200)
         lblItems(i).Width = Max(0, iTmp - 25)
      Case CheckBox
         chkItems(i).Width = Min(iTmp, 200)
         lblItems(i).Width = Max(0, iTmp - 250)
      Case Else
         lblItems(i).Width = Max(0, iTmp - 50)
   End Select
Next

If Index = Slider.Count - 1 Then
   iTmp = ListBox.Width - Slider(Index).Left
Else
   iTmp = Slider(Index + 1).Left - Slider(Index).Left
End If

For i = iMinimum + iMaxInColumn To iMaximum - IIf(Index = Slider.Count - 1, 0, 1) Step 1
   Select Case m_ListStyle
      Case PictureBox
         picItems(i).Width = Min(iTmp, lblItems(0).Height)
         picItems(i).Move Slider(Index).Left + 35
         lblItems(i).Width = Max(0, iTmp - (picItems(i).Width + 90))
         lblItems(i).Move Slider(Index).Left + 35 + picItems(i).Width + 25
      Case ImageBox
         imgItems(i).Width = Min(iTmp, 200)
         imgItems(i).Move Slider(Index).Left + 50
         lblItems(i).Width = Max(iTmp - 60, 0)
         lblItems(i).Move Slider(Index).Left + 35
      Case CheckBox
         chkItems(i).Width = Min(iTmp, 200)
         chkItems(i).Move Slider(Index).Left + 35
         lblItems(i).Width = Max(0, iTmp - 260)
         lblItems(i).Move Slider(Index).Left + 35 + chkItems(i).Width + 10
      Case Else
         lblItems(i).Width = Max(0, iTmp - 60)
         lblItems(i).Move Slider(Index).Left + 35
   End Select
Next
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If PropertyName = "DisplayName" Then Call ShowDispName
End Sub


Private Sub UserControl_Show()
Dim l       As Long

Const WS_EX_TOOLWINDOW = &H80
Const GWL_EXSTYLE = -20
Const GWL_STYLE = -16

If bRefresh Or bRefreshControls Then Refresh

SetParent ToolTip.hwnd, 0

l = GetWindowLong(ToolTip.hwnd, GWL_EXSTYLE)
Call SetWindowLong(ToolTip.hwnd, GWL_EXSTYLE, l Or WS_EX_TOOLWINDOW)

SetWindowPos ToolTip.hwnd, ToolTip.hwnd, 0, 0, 0, 0, 39 '&H1 + &H2 + &H4 + &H20
Call SetWindowLong(ToolTip.hwnd, -8, ListBox.hwnd)

Call ShowDispName

End Sub

Public Function Version() As String
Version = AppVersion()
End Function

Public Property Get Updated() As Boolean
Updated = m_Updated
End Property

Public Property Let Updated(ByVal New_Updated As Boolean)
m_Updated = New_Updated
End Property

Private Sub UnLoadControls()
Dim i As Integer

For i = 1 To lblItems.Count - 1 Step 1
   Unload lblItems(i)
Next
For i = 1 To chkItems.Count - 1 Step 1
   Unload chkItems(i)
Next
For i = 1 To picItems.Count - 1 Step 1
   Unload picItems(i)
Next
For i = 1 To imgItems.Count - 1 Step 1
   Unload imgItems(i)
Next
For i = 1 To Slider.Count - 1 Step 1
   Unload Slider(i)
Next

End Sub

Public Sub ShowSplash()
'Call SetWindowPos(frmSplash.hwnd, -1, 0, 0, 0, 0, 3)
'Load frmSplash
'frmSplash.Show
End Sub

Public Property Get SelectOnMatch() As Boolean
Attribute SelectOnMatch.VB_ProcData.VB_Invoke_Property = "Characteristics"
SelectOnMatch = m_SelectOnMatch
End Property

Public Property Let SelectOnMatch(ByVal New_SelectOnMatch As Boolean)
m_SelectOnMatch = New_SelectOnMatch
PropertyChanged "SelectOnMatch"
End Property

Public Property Get CurrPosItem() As Long
If iIndex > -1 And iIndex <= iMaxVisible Then
   CurrPosItem = lblItems(iIndex).Tag
End If
End Property

Public Property Let CurrPosItem(ByVal New_ListItem As Long)
Dim l As Long
If New_ListItem <= List.Count And New_ListItem > -1 Then
   l = CurrPosItem
   If New_ListItem = l + 1 Then
      UserControl_KeyDown vbKeyDown, 0
   ElseIf New_ListItem = l - 1 Then
      UserControl_KeyDown vbKeyUp, 0
   Else
      Call Selector(New_ListItem)
   End If
End If
End Property

