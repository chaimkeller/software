Attribute VB_Name = "modGestures"
Option Explicit

'==========================================
'   PUBLIC GESTURE DEFINITIONS
'==========================================
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_GESTURE As Long = &H119
Public Const WM_GESTURENOTIFY As Long = &H11A
Public Const WM_POINTERDOWN As Long = &H246
Public Const WM_POINTERUPDATE As Long = &H245
Public Const WM_POINTERUP As Long = &H247

Public Const GID_BEGIN = 1
Public Const GID_END = 2
Public Const GID_ZOOM = 3
Public Const GID_PAN = 4
Public Const GID_ROTATE = 5
Public Const GID_TWOFINGERTAP = 6
Public Const GID_PRESSANDTAP = 7

' Some implementations use GF_BEGIN flag = &H1, but
' we’ll just test dwFlags = 1 for "begin" as you had.
Public Const GF_BEGIN As Long = &H1

Public Type GESTUREINFO
    cbSize As Long
    dwFlags As Long
    dwID As Long
    hwndTarget As Long
    ptsLocation As Currency
    dwInstanceID As Long
    dwSequenceID As Long
    ullArguments(1) As Long
    cbExtraArgs As Long
End Type

'==========================================
'   POINTER INPUT (touchscreens)
'==========================================
Public Type POINTER_INFO
    pointerType As Long
    pointerId As Long
    frameId As Long
    pointerFlags As Long
    sourceDevice As Long
    hwndTarget As Long
    ptPixelLocationX As Long
    ptPixelLocationY As Long
    ptHimetricLocationX As Long
    ptHimetricLocationY As Long
    ptPixelLocationRawX As Long
    ptPixelLocationRawY As Long
    ptHimetricLocationRawX As Long
    ptHimetricLocationRawY As Long
    dwTime As Long
    historyCount As Long
    inputData As Long
    dwKeyStates As Long
    PerformanceCount As Currency
    ButtonChangeType As Long
End Type

Public Declare Function GetPointerInfo Lib "user32" _
    (ByVal pointerId As Long, ByRef info As POINTER_INFO) As Long
    
'==========================================
'   POINTER STATE
'==========================================
Public PointerCount As Long
Public Pointer1X As Long, Pointer1Y As Long
Public Pointer2X As Long, Pointer2Y As Long
Public LastDistance As Double

Public Type POINTF
    x As Double
    y As Double
End Type

Public Type GEOF
    lat As Double
    lon As Double
End Type

Public Declare Function GetGestureInfo Lib "user32" ( _
    ByVal hGestureInfo As Long, _
    ByRef pGestureInfo As GESTUREINFO) As Long

Public Declare Function CloseGestureInfoHandle Lib "user32" ( _
    ByVal hGestureInfo As Long) As Long

'==========================================
'   SUBCLASSING API
'==========================================
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
Public Declare Function RegisterPointerInputTarget Lib "user32" _
    (ByVal hWnd As Long, ByVal pointerType As Long) As Long

Public Const PT_TOUCH = 2

Public Const GWL_WNDPROC As Long = -4

Public Const AngleStep As Double = 5 'degrees of rotation for each rotation step

'Public PI As Double 'already declared in Maps
'Public cd As Double   'conversion from degrees to radians 'already declared in Maps
Public Const EARTH_RADIUS As Double = 6371000#   ' meters

'---------------------------
' Globals
'---------------------------
Public OldWndProc As Long
Public PictureFileName As String
Public MainForm As Form
Public g_gx(0 To 3) As Double, g_gy(0 To 3) As Double

' Viewer state
Public ZoomFactor As Double
Public PanX As Double
Public PanY As Double
Public Dragging As Boolean
Public lastX As Long
Public lastY As Long
Public firstX As Long
Public firstY As Long
Public CurMouseX As Long
Public CurMouseY As Long

' Gesture state
Public LastGestureZoom As Double
Public LastGestureTime As Long

' For reading 64-bit zoom argument from ullArguments()
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
    
'==========================================
'   CENTRAL WINDOW PROCEDURE
'==========================================
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long) As Long
                        
'////////////////////debugging/////////////////////////
'If uMsg = WM_POINTERDOWN Or uMsg = WM_POINTERUPDATE Or uMsg = WM_POINTERUP Then
'        Debug.Print "WM_POINTER message:", uMsg
'    End If


    Select Case uMsg
    '------------------------------
    ' Mouse wheel ? zoom
    '------------------------------
    Case WM_MOUSEWHEEL
        
        Dim delta As Long
        delta = (wParam \ &H10000) And &HFFFF&
        If delta > 32767 Then delta = delta - 65536
        
        If delta > 0 Then
            MainForm.ZoomAtCursor 1.1
        ElseIf delta < 0 Then
            MainForm.ZoomAtCursor 1 / 1.1
        End If

        WndProc = 0
        Exit Function

    '------------------------------
    ' Touch gestures
    '------------------------------
    Case WM_GESTURE
        HandleGesture wParam, lParam
        WndProc = 0
        Exit Function

    End Select
    
    WndProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, lParam)
End Function

'==========================================
'   GESTURE DISPATCH
'==========================================
Public Sub HandleGesture(ByVal wParam As Long, ByVal lParam As Long)
    Dim gi As GESTUREINFO
    gi.cbSize = Len(gi)
    
'    Debug.Print "ID:", gi.dwID, "Flags:", gi.dwFlags, "Args:", gi.ullArguments(0), gi.ullArguments(1)

    If GetGestureInfo(lParam, gi) Then

        Select Case gi.dwID

        Case GID_BEGIN
            GestureBegin gi

        Case GID_END
            GestureEnd gi

        Case GID_PAN
            PanGesture gi

        Case GID_ZOOM
            ZoomGesture gi

        Case GID_ROTATE
            ' Optional: MainForm.RotateGesture gi

        Case GID_TWOFINGERTAP
            ' Optional: MainForm.TwoFingerTap gi

        Case GID_PRESSANDTAP
            ' Optional: MainForm.PressAndTap gi

        End Select

    End If

    CloseGestureInfoHandle lParam
End Sub
Public Sub GestureBegin(gi As GESTUREINFO)
    ' your code
End Sub

Public Sub GestureEnd(gi As GESTUREINFO)
    ' your code
End Sub

Public Sub ZoomGesture(gi As GESTUREINFO)
    Static LastDistance As Double
    Dim curDistance As Double

    ' ullArguments holds the zoom distance as a 64-bit value
    CopyMemory curDistance, gi.ullArguments(0), 8

    If (gi.dwFlags And GF_BEGIN) <> 0 Then
        LastDistance = curDistance
        Exit Sub
    End If

    If LastDistance = 0 Then
        LastDistance = curDistance
        Exit Sub
    End If

    Dim factor As Double
    factor = curDistance / LastDistance

    LastDistance = curDistance

    If factor > 0 Then
        MainForm.ZoomAtCursor factor
    End If
End Sub

Public Sub PanGesture(gi As GESTUREINFO)
    Static lastX As Long, lastY As Long
    Static LastDistance As Double

    Dim curX As Long, curY As Long
    curX = CLng(gi.ptsLocation / 65536@)
    curY = CLng((gi.ptsLocation And &HFFFF&) / 65536@)

    ' Begin gesture
    If (gi.dwFlags And GF_BEGIN) <> 0 Then
        lastX = curX
        lastY = curY
        LastDistance = 0
        Exit Sub
    End If

    ' Movement delta
    Dim dx As Long, dy As Long
    dx = curX - lastX
    dy = curY - lastY

    ' Detect two-finger gesture by checking ullArguments
    Dim distance As Double
    CopyMemory distance, gi.ullArguments(0), 8

    If distance > 0 Then
        ' Two-finger gesture ? zoom
        If LastDistance = 0 Then LastDistance = distance

        Dim factor As Double
        factor = distance / LastDistance

        ' Clamp zoom factor
        If factor < 0.5 Then factor = 0.5
        If factor > 2# Then factor = 2#

        MainForm.ZoomAtCursor factor

        LastDistance = distance
    Else
        ' One-finger gesture ? pan
        PanX = PanX + dx
        PanY = PanY + dy
        MainForm.RedrawView
    End If

    lastX = curX
    lastY = curY
End Sub

'==========================================
'   TOUCHSCREEN POINTER HANDLERS
'==========================================
Public Sub HandlePointerDown(pointerId As Long)
    Dim info As POINTER_INFO
    If GetPointerInfo(pointerId, info) = 0 Then Exit Sub
    
    Debug.Print "POINTERDOWN id=", info.pointerId, "count=", PointerCount + 1

    PointerCount = PointerCount + 1

    If PointerCount = 1 Then
        Pointer1X = info.ptPixelLocationX
        Pointer1Y = info.ptPixelLocationY

    ElseIf PointerCount = 2 Then
        Pointer2X = info.ptPixelLocationX
        Pointer2Y = info.ptPixelLocationY

        LastDistance = Sqr((Pointer2X - Pointer1X) ^ 2 + (Pointer2Y - Pointer1Y) ^ 2)
    End If
End Sub


Public Sub HandlePointerUpdate(pointerId As Long)
    Dim info As POINTER_INFO
    If GetPointerInfo(pointerId, info) = 0 Then Exit Sub
    
    Debug.Print "POINTERUPDATE id=", info.pointerId, "count=", PointerCount

    ' One finger ? pan
    If PointerCount = 1 Then
        MainForm.PointerPan info.ptPixelLocationX - Pointer1X, info.ptPixelLocationY - Pointer1Y
        Pointer1X = info.ptPixelLocationX
        Pointer1Y = info.ptPixelLocationY
        Exit Sub
    End If

    ' Two fingers ? zoom
    If PointerCount = 2 Then

        ' Update whichever finger moved
        If info.pointerId = 1 Then
            Pointer1X = info.ptPixelLocationX
            Pointer1Y = info.ptPixelLocationY
        Else
            Pointer2X = info.ptPixelLocationX
            Pointer2Y = info.ptPixelLocationY
        End If

        Dim dist As Double
        dist = Sqr((Pointer2X - Pointer1X) ^ 2 + (Pointer2Y - Pointer1Y) ^ 2)

        If LastDistance > 0 Then
            Dim factor As Double
            factor = dist / LastDistance

            ' Clamp factor to avoid overflow
            If factor < 0.5 Then factor = 0.5
            If factor > 2# Then factor = 2#

            MainForm.PointerZoom factor
        End If

        LastDistance = dist
    End If
End Sub


Public Sub HandlePointerUp(pointerId As Long)
    PointerCount = PointerCount - 1
    If PointerCount < 0 Then PointerCount = 0
End Sub

Public Function LoadPNG(strPath As String) As StdPicture
    Dim img As Object
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile strPath
    Set LoadPNG = img.FileData.Picture
End Function

