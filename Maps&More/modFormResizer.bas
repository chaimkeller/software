Attribute VB_Name = "modFormResizer"
Option Explicit

Private Type CONTROLINFO
    Name As String
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private OriginalFormWidth As Single
Private OriginalFormHeight As Single
Private Controls() As CONTROLINFO
Private ControlCount As Long

'===========================================
'  Controls to EXCLUDE from resizing
'===========================================
Private Function IsResizableControl(c As Control) As Boolean

    ' NEVER resize the form itself
    If TypeOf c Is Form Then
        IsResizableControl = False
        Exit Function
    End If

    ' NEVER resize the viewport
    If c.Name = "picView" Then
        IsResizableControl = False
        Exit Function
    End If

    ' NEVER resize the source image buffer
    If c.Name = "picSource" Then
        IsResizableControl = False
        Exit Function
    End If

    ' NEVER resize CommonDialog
    If TypeOf c Is CommonDialog Then
        IsResizableControl = False
        Exit Function
    End If

    ' NEVER resize ImageList
    If TypeOf c Is ImageList Then
        IsResizableControl = False
        Exit Function
    End If

    ' Everything else is OK
    IsResizableControl = True
End Function




'===========================================
'  Save / Restore Window Position
'===========================================

Public Sub SaveWindowState(frm As Form)
    On Error Resume Next

    If frm.WindowState <> vbNormal Then Exit Sub

    SaveSetting App.Title, "WindowState", "Left", CStr(frm.Left)
    SaveSetting App.Title, "WindowState", "Top", CStr(frm.Top)
    SaveSetting App.Title, "WindowState", "Width", CStr(frm.Width)
    SaveSetting App.Title, "WindowState", "Height", CStr(frm.Height)
    SaveSetting App.Title, "Picture", "FileName", CStr(PictureFileName)
End Sub

Public Sub RestoreWindowState(frm As Form)
    On Error Resume Next

    Dim L As Long, T As Long, W As Long, H As Long

    L = CLng(GetSetting(App.Title, "WindowState", "Left", frm.Left))
    T = CLng(GetSetting(App.Title, "WindowState", "Top", frm.Top))
    W = CLng(GetSetting(App.Title, "WindowState", "Width", frm.Width))
    H = CLng(GetSetting(App.Title, "WindowState", "Height", frm.Height))

    frm.Move L, T, W, H
    
    PictureFileName = GetSetting(App.Title, "Picture", "FileName")
End Sub

'===========================================
'  Initialize Resizer
'===========================================

Public Sub InitResizer(frm As Form)
    Dim c As Control
    Dim i As Long

    OriginalFormWidth = frm.ScaleWidth
    OriginalFormHeight = frm.ScaleHeight

    ControlCount = 0

    ' Count only resizable controls
    For Each c In frm.Controls
        If IsResizableControl(c) Then
            ControlCount = ControlCount + 1
        End If
    Next

    If ControlCount = 0 Then Exit Sub

    ReDim Controls(1 To ControlCount)

    i = 1
    For Each c In frm.Controls
        If IsResizableControl(c) Then
            On Error Resume Next
            Controls(i).Name = c.Name
            Controls(i).Left = c.Left
            Controls(i).Top = c.Top
            Controls(i).Width = c.Width
            Controls(i).Height = c.Height
            Controls(i).FontSize = c.Font.size
            i = i + 1
        End If
    Next
End Sub

'===========================================
'  Resize Controls
'===========================================

Public Sub ResizeControls(frm As Form)
    If OriginalFormWidth = 0 Or OriginalFormHeight = 0 Then Exit Sub
    If frm.WindowState = vbMinimized Then Exit Sub

    Dim xRatio As Double
    Dim yRatio As Double
    Dim c As Control
    Dim i As Long
    Dim cc As Integer, dd As Integer
    
'    On Error GoTo errhand

    xRatio = frm.ScaleWidth / OriginalFormWidth
    yRatio = frm.ScaleHeight / OriginalFormHeight

    i = 1
    For Each c In frm.Controls

        If c.Name = "picView" Then GoTo SkipControl
        If c.Name = "picSource" Then GoTo SkipControl
        If TypeOf c Is CommonDialog Then GoTo SkipControl
        If TypeOf c Is ImageList Then GoTo SkipControl
        If TypeOf c Is Form Then GoTo SkipControl
    
        If IsResizableControl(c) Then
            On Error Resume Next
    
            c.Left = Controls(i).Left * xRatio
            c.Top = Controls(i).Top * yRatio
            c.Width = Controls(i).Width * xRatio
            c.Height = Controls(i).Height * yRatio
    
            If Controls(i).FontSize > 0 Then
                c.Font.size = Controls(i).FontSize * ((xRatio + yRatio) / 2)
            End If
            
            If c.Name = "lstCP" Then 'shift it to the picview's end
               c.Left = frmViewer.picView.Left + frmViewer.picView.Width - frmViewer.lstCP.Width - 10
               End If
               
            If c.Name = "Toolbar1" Then
               frmViewer.Toolbar1.Refresh 'without this, the toolbar1 of the frmViewer is overpainted by the map
               End If
    
            i = i + 1
        End If
    
SkipControl:
    Next

'    For Each c In frm.Controls
'        If IsResizableControl(c) Then
'
''            If c.Name = "picView" Then GoTo SkipControl
'
'            Select Case c.Name
'                Case "picView"
'                    cc = 1
'                Case "Toolbar1"
'                     cc = 2
'               Case "progsearch"
'                     cc = 3
'               Case "lstAverage"
'                     cc = 4
'               Case "StatusBar1"
'                     cc = 5
'               Case "frmCoords"
'                     cc = 6
'               Case "lblCoords"
'                     cc = 7
'               Case "lstCP"
'                    cc = 8
'                Case "chkLstCP"
'                    cc = 9
'                Case "picSource"
'                    cc = 10
'                Case "ImageList1"
'                    cc = 11
'                Case "comdlg"
'                    cc = 12
'                Case "frmViewer"
'                    cc = 13
'                Case "picSource"
'                    cc = 14
'                Case Else
'                    cc = 14
'             End Select
'
'            c.Left = Controls(i).Left * xRatio
'            c.Top = Controls(i).Top * yRatio
'            c.Width = Controls(i).Width * xRatio
'            c.Height = Controls(i).Height * yRatio
'
'            If Controls(i).FontSize > 0 Then
'                c.Font.Size = Controls(i).FontSize * ((xRatio + yRatio) / 2)
'            End If
'
'            i = i + 1
'        End If
'
'SkipControl:
'    Next
'Exit Sub
'
'errhand:
'    dd = cc
'    Resume Next
End Sub

