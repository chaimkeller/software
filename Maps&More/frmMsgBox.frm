VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAction 
      Height          =   375
      Index           =   0
      Left            =   345
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't ask me again"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblMsg 
      Caption         =   "Msg"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////EK 061422////////////////
'Note: if background window has a setposition to topmost, then it must be set to notopmost before
'calling this procedure, otherwise the message box disappears and the top window freezes
'///////////////////////////////////////////////////////////////////

'************************************************************************************
'* frmMsgBox.frm
'* By John R Parrish
'* Copyright 2006
'* All Rights reserved
'************************************************************************************
'* You are free to use this form in your own personal and commercial projects
'* Provided that
'*  a. You do not sell the source code to a 3rd party
'*  b. You accept ALL responsibility for any problems
'*     that may be caused by using this form
'************************************************************************************
'* This form makes as many buttons as you need for the options you
'* want to give to the user.
'* It can also display and return the value of a "Don't ask me again" checkbox
'*
'* It has two functions that can be called
'*      frmMsgBox.Msg: works similar to a normal MsgBox, but it's easier to use
'*                     Only the Prompt is required
'*      frmMsgBox.MsgCstm: Allows you to select your own button captions
'*                         and add as many as you want.
'*                         A ParramArray is used for the button names
'*                         This prevents using Optional parameters :/
'*
'* Both functions return the index number of the button that was clicked (1 based)
'*      '0' always indicates that the user closed the box without hitting a button
'* You can also access the user selected options through the form's two
'*      Global variables: frmMsgBox.g_lBtnClicked and frmMsgBox.g_bDontAsk
'*
'************************************************************************************
'* Example uses
'*      Converting a standard MsgBox to frmMsgBox.Msg:
'*          lOption = MsgBox("Test", vbYesNoCancel Or vbCritical Or vbDefaultButton2)
'*              Use:
'*          lOption = frmMsgBox.Msg("Test", mbYesNoCancel, mbCritical, 2)
'*
'*      Custom MsgBox use:
'*          frmMsgBox.MsgCstm "Want a cup of coffee?", "Coffee?", mbQuestion, 1, True, _
'*                            "Yes", "No", "Maybe", "Ask me later"
'*          Select Case frmMsgBox.g_lBtnClicked
'*             Case 0 ' 0 always indicates that the user closed the box without hitting a button
'*             Case 1 'the 1st button in your list was clicked
'*             Case 2 'the 2nd button in your list was clicked
'*             Case 3 'ect.
'*             Case 4
'*          End Select
'*          bDontAsk = frmMsgBox.g_bDontAsk
'*
'************************************************************************************
'* Tips:
'*  a. This looks great when a Manifest File is used
'*  b. You can match the form to your skinned project by:
'*      1. Delete the command button "btnAction"
'*      2. Add your own custom button and name it "btnAction"
'*      3. Set it's index to 0 and adjust it's height
'*  c. You can copy frmMsgBox.frm and frmMsgBox.frx to the
'*     VB Templates folder to make it accessible from the
'*     VB 'Project/Add Form' menu
'*      The template folder is usually located at:
'*      C:\Program Files\Microsoft Visual Studio\VB98\Template\Forms
'************************************************************************************

Public Enum eMsgIcon
    mbNone
    mbExclamation
    mbInformation
    mbCritical
    mbQuestion
End Enum

Public Enum StandardIconEnum
    IDI_ASTERISK = 32516&
    IDI_EXCLAMATION = 32515&
    IDI_HAND = 32513&
    IDI_QUESTION = 32514&
End Enum

Public Enum eBtns
    mbOKOnly
    mbOKCancel
    mbAbortRetryIgnore
    mbYesNoCancel
    mbYesNo
    mbRetryCancel
End Enum

Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" _
                                        (ByVal hInstance As Long, _
                                         ByVal lpIconNum As StandardIconEnum) As Long
    
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
                                                ByVal x As Long, _
                                                ByVal y As Long, _
                                                ByVal hIcon As Long) As Long

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Const MB_ICONEXCLAMATION = 49
Private Const MB_ICONHAND = 17
Private Const MB_ICONINFORMATION = 65

'varables used to return the users chossen options
Public g_lBtnClicked As Long
Public g_bDontAsk As Boolean

Public Function Msg(ByRef Promt As String, _
                    Optional ByRef Buttons As eBtns = mbOKOnly, _
                    Optional ByRef Title As String, _
                    Optional ByRef MsgIcon As eMsgIcon = mbNone, _
                    Optional ByRef DefaultBtn As Long = 1, _
                    Optional ByRef ShowDontAsk As Boolean = False) As Long
'If Title is left blank, App.Path is used
'If you don't want any title just set Title = " "
'Btn captions are made automatically
    Dim sBtnText() As String
    
    SetTitleBar Title
    LoadIcon MsgIcon
    SetLabelWidth Promt, MsgIcon
    
    Select Case Buttons
        Case mbOKOnly
            sBtnText = Split("Ok")
        Case mbOKCancel
            sBtnText = Split("Ok|Cancel", "|")
        Case mbAbortRetryIgnore
            sBtnText = Split("Abort|Retry|Ignore", "|")
        Case mbYesNoCancel
            sBtnText = Split("Yes|No|Cancel", "|")
        Case mbYesNo
            sBtnText = Split("Yes|No", "|")
        Case mbRetryCancel
            sBtnText = Split("Retry|Cancel", "|")
        Case Else
            sBtnText = Split("Ok")
    End Select
    SetBtns sBtnText, MsgIcon
    SetDontAsk ShowDontAsk
    PositionForm DefaultBtn 'ParentForm, Me.Width
    Msg = g_lBtnClicked
End Function

Public Function MsgCstm(ByRef Promt As String, _
                        ByRef Title As String, _
                        ByRef MsgIcon As eMsgIcon, _
                        ByRef DefaultBtn As Long, _
                        ByRef ShowDontAsk As Boolean, _
                        ParamArray btnText()) As Long
'sets the user msg, Title and number and text on the buttons
'If Title is left blank, App.Path is used
    Dim lB As Long
    Dim sBtnText() As String

    SetTitleBar Title
    LoadIcon MsgIcon
    SetLabelWidth Promt, MsgIcon
    
    ReDim sBtnText(UBound(btnText))
    For lB = 0 To UBound(btnText)
        sBtnText(lB) = CStr(btnText(lB))
    Next
    SetBtns sBtnText, MsgIcon
    SetDontAsk ShowDontAsk
    PositionForm DefaultBtn 'ParentForm, Me.Width
    MsgCstm = g_lBtnClicked
End Function

Private Sub btnAction_Click(Index As Integer)
    'add one to the Btn#, 0 indicates they closed without hitting any btn
    g_lBtnClicked = Index + 1
    g_bDontAsk = CBool(chkDontAsk.value)
    Unload Me
End Sub

Private Sub form_load()
    g_lBtnClicked = 0
    g_bDontAsk = False
    Me.Font = btnAction(0).Font
    lblMsg.Font = Me.Font
    Call sCenterForm(Me)  'added by EK on 030524
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim obj As Object

    On Error Resume Next
    Me.Cls
    For Each obj In Me 'frm
        Unload obj
        Set obj = Nothing
    Next
    On Error GoTo 0
End Sub

Private Sub LoadIcon(ByRef MsgIcon As eMsgIcon)
    Dim hIcon As Long

    If MsgIcon Then
        'show the icon and play the right sound
        Select Case MsgIcon
            Case mbExclamation
                hIcon = LoadStandardIcon(0&, IDI_EXCLAMATION)
                MessageBeep MB_ICONEXCLAMATION '49
            Case mbInformation
                hIcon = LoadStandardIcon(0&, IDI_ASTERISK)
                MessageBeep MB_ICONINFORMATION '65
            Case mbCritical
                hIcon = LoadStandardIcon(0&, IDI_HAND)
                MessageBeep MB_ICONHAND '17
            Case mbQuestion
                hIcon = LoadStandardIcon(0&, IDI_QUESTION)
        End Select
        Call DrawIcon(Me.hdc, 9&, 10&, hIcon)
    Else
        Me.Cls
    End If
End Sub

Private Sub PositionForm(ByRef DefaultBtn As Long)

    Dim ret As Long

    If DefaultBtn > btnAction.count Then
        DefaultBtn = btnAction.count
    ElseIf DefaultBtn < 1 Then
        DefaultBtn = 1
    End If
    btnAction(DefaultBtn - 1).TabIndex = 0
    Me.Show vbModal

End Sub

Private Sub SetBtns(ByRef btnText() As String, _
                    ByRef MsgIcon As eMsgIcon)
    Dim lX As Long
    Dim lUb As Long
    Dim lWidth As Long
    Dim lRightMost As Long
    Dim lRowWidth() As Long
    Dim lBtnsInRow() As Long
    Dim lR As Long
    Dim lCnt As Long
    Dim lBtnTop As Long
    Dim lMaxWidth As Long
    Dim lHeight As Long
    
    lBtnTop = lblMsg.Height + (lblMsg.Top * 2)
    If lBtnTop < 900 Then
        If MsgIcon Then
            lBtnTop = 900
        End If
    End If
    btnAction(0).Top = lBtnTop
    
    Select Case Me.Width
        Case Is < Screen.Width / 4
            lMaxWidth = Screen.Width / 4
        Case Is < Screen.Width / 2
            lMaxWidth = Screen.Width / 2
        Case Else ' Is < Screen.Width / 4
            lMaxWidth = Screen.Width * 0.75
    End Select
    
    ReDim lRowWidth(0)
    ReDim lBtnsInRow(0)
    lUb = UBound(btnText)
    For lX = 0 To lUb
        With btnAction(lX)
            If lX Then
                'dynamically load the needed buttons
                Load btnAction(lX)
                .Top = btnAction(lX - 1).Top
                .Left = btnAction(lX - 1).Left + btnAction(lX - 1).Width + 120
            End If
            'set the button width and text
            .Width = Me.TextWidth(btnText(lX) & "WW") '"WW" is a buffer to make extra room on the button
            .Caption = btnText(lX)
            .Visible = True
            'wrap the buttons if needed
            If .Width + .Left + 120 > lMaxWidth Then
                lR = lR + 1
                ReDim Preserve lRowWidth(lR)
                ReDim Preserve lBtnsInRow(lR)
                .Left = btnAction(0).Left
                .Top = btnAction(lX - 1).Top + btnAction(lX - 1).Height + 120
                If btnAction(lX - 1).Left + btnAction(lX - 1).Width > _
                    btnAction(lRightMost).Left + btnAction(lRightMost).Width Then
                    lRightMost = lX - 1
                End If
            End If
            lRowWidth(lR) = lRowWidth(lR) + .Width + 120
            lBtnsInRow(lR) = lBtnsInRow(lR) + 1
        End With
    Next
    
    'adjust the width of the msg box
    lWidth = Me.Width
    If lRightMost = 0 Then
        lRightMost = lUb
    End If
    If btnAction(lRightMost).Left + btnAction(lRightMost).Width + btnAction(0).Left > lWidth Then
        lWidth = btnAction(lRightMost).Left + btnAction(lRightMost).Width + btnAction(0).Left
    End If
    Me.Width = lWidth
    
    'center the button rows
    For lUb = 0 To UBound(lRowWidth)
        lWidth = lRowWidth(lUb) - 120
        lWidth = ((Me.Width - lWidth) / 2) - 30
        For lR = 0 To lBtnsInRow(lUb) - 1
            If lR = 0 Then
                btnAction(lCnt).Left = lWidth
            Else
                btnAction(lCnt).Left = btnAction(lCnt - 1).Left + btnAction(lCnt - 1).Width + 120
            End If
            lCnt = lCnt + 1
        Next
    Next
End Sub

Private Sub SetDontAsk(ByRef ShowDontAsk As Boolean)
    Dim lUb As Long
    
    lUb = btnAction.count - 1
    'set the height of the form
    If ShowDontAsk Then
        chkDontAsk.value = 0
        chkDontAsk.Top = btnAction(lUb).Top + btnAction(lUb).Height + 120
        chkDontAsk.Visible = True
        Me.Height = chkDontAsk.Top + chkDontAsk.Height + 630 '585 '645
    Else
        chkDontAsk.Visible = False
        Me.Height = btnAction(lUb).Top + btnAction(lUb).Height + 630 '585 '645
    End If
End Sub

Private Sub SetLabelWidth(ByRef Promt As String, ByRef MsgIcon As eMsgIcon)
'Make sure that the Promt Label doesn't cause the form to be wider that the screen.
    
    lblMsg.Caption = Promt
    lblMsg.Width = Me.TextWidth(Promt)
    If MsgIcon Then
        lblMsg.Left = 780
    Else
        lblMsg.Left = 180
    End If
    Me.Width = lblMsg.Left + lblMsg.Width + 240 '120
    If Me.Width < 3330 Then Me.Width = 3330
    lblMsg.Height = Me.TextHeight(Promt)
    If lblMsg.Left + lblMsg.Width + 240 > Screen.Width * 0.75 Then
        lblMsg.AutoSize = True
        lblMsg.WordWrap = True
        lblMsg.Width = (Screen.Width * 0.75) - (lblMsg.Left + 120)
        Me.Width = Screen.Width
    Else
        lblMsg.WordWrap = False
    End If
End Sub

Private Sub SetTitleBar(ByRef Title As String)
    If Len(Title) Then
        Me.Caption = Title
    Else
        Me.Caption = App.Title 'Path
    End If
End Sub

