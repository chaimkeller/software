Attribute VB_Name = "rotate_text"
Option Explicit

'****************************************************************************************************
'
' Name          : api_text
' Author        : Dennis Burns
' Email         : nextlemming@aol.com
' Date          : Sept 26, 2001
' Description   : function and api calls to print rotated text.
'
'
' Notice        : This code is open to the public domain,
'                   just give credit where credit is due.
'
'****************************************************************************************************


Private Const LF_FACESIZE = 32
Private Const SYSTEM_FONT = 13

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long


Public Function TextRotate(hdc As Long, Output As String, x As Long, y As Long, Rotation As Single) As Boolean
    ' 8.26.01 Dennis Burns
    ' Function to print rotated where the target is anything that you can get
    '  an hdc for.
    '
    ' Parameters
    '   hdc       -   a handle to a device context.
    '   Output    -   Standard vb string
    '   x         -   horizontal position for text in logical units
    '   y         -   vertical position for text in logical units
    '   Rotation  -   Angle to rotate text to. Decimal Degrees 0 = normal,
    '                   90 = vertical, 180 = upside down.
    ' Return true on sucess

On Error GoTo ErrorTextRotate

    Dim lf As LOGFONT       'Logical font structure
    Dim FontToUse As Long   'Reference to font created with rotated text
    Dim oldFont As Long     'Reference to curent font for object
    Dim Result As Long      'Holds result of api calls, used for error checking
    
    'First, get a reference to the current font, yse the system font
    '   is the font for the VB objects.
    oldFont = SelectObject(hdc, GetStockObject(SYSTEM_FONT))
    'oldFont = SelectObject(hdc, tempFont)
    If oldFont = 0 Then GoTo ErrorTextRotate

    'Second, get the font structure with data from current font.
    Result = GetObjectAPI(oldFont, Len(lf), lf)
    If Result = 0 Then GoTo ErrorTextRotate
    
    ' Change the Escapement to rotate font (too bad they did not call it
    '   rotation so it could be found in a search)
    '   10 = 1 degree CCW from horizontal
    lf.lfEscapement = Rotation * 10
    
    'Create a new font the same as current font but rotated
    FontToUse = CreateFontIndirect(lf)
    If FontToUse = 0 Then GoTo ErrorTextRotate
    
    'Select font into object - returns reference to current font
    '   must be saved to restore.
    oldFont = SelectObject(hdc, FontToUse)
    If oldFont = 0 Then GoTo ErrorTextRotate
    
    ' this is where we actually output.
    Result = TextOut(hdc, x, y, Output, Len(Output))
    If Result = 0 Then GoTo ErrorTextRotate
    
    'Restore old font
    Result = SelectObject(hdc, oldFont)
    If Result = 0 Then GoTo ErrorTextRotate
    
    ' MUST delete font or there will be a resource leek
    Result = DeleteObject(FontToUse)
    If Result = 0 Then GoTo ErrorTextRotate


    TextRotate = True
    Exit Function
    

ErrorTextRotate:
    TextRotate = False
    
End Function

'///////////////////////////////////////////////////////////////
' ----------------------------------------
' The Rotator class module
'
' A class for printing rotated text to a
' Form, PictureBox or the Printer
'
' Author: Timm Dickel (Tim.Dic@t-online.de)
'
' ----------------------------------------
' Usage:
'     Dim rotTest As New Rotator
'     Set rotTest.Device = Printer
'     ' set all font attributes as required, e.g.
'     Printer.Font.Size = 12
'     'Label strings at a variety of angles
'     For nA = 0 To 359 Step 15
'        rotTest.Angle = nA
'        rotTest.PrintText Space(10) & Printer.Font.Name & Str(nA)
'     Next
'     Printer.EndDoc
'
' ----------------------------------------

Option Explicit

'API constants
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSY = 90

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As _
    Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
    "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As _
    Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, _
    ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long

'Module-level private variables
Private mobjDevice As Object
Private mfSX1 As Single
Private mfSY1 As Single
Private mfXRatio As Single
Private mfYRatio As Single
Private lfFont As LOGFONT
Private mnAngle As Integer

'~~~Angle
Property Let Angle(nAngle As Integer)
    mnAngle = nAngle
End Property
Property Get Angle() As Integer
    Angle = mnAngle
End Property

'~~~PrintText
Public Sub PrintText(sText As String)
    Dim lFont As Long
    Dim lOldFont As Long
    Dim lRes As Long
    Dim byBuf() As Byte
    Dim nI As Integer
    Dim sFontName As String
    Dim mobjDevicehdc As Long
    Dim mobjDeviceCurrentX As Single
    Dim mobjDeviceCurrentY As Single
    
    mobjDevicehdc = mobjDevice.hdc
    mobjDeviceCurrentX = mobjDevice.CurrentX
    mobjDeviceCurrentY = mobjDevice.CurrentY
    
    'Prepare font name, decoding from Unicode
    sFontName = mobjDevice.Font.Name
    byBuf = StrConv(sFontName & Chr$(0), vbFromUnicode)
    For nI = 0 To UBound(byBuf)
        lfFont.lfFaceName(nI) = byBuf(nI)
    Next nI
    
    'Convert known font size to required units
    lfFont.lfHeight = mobjDevice.Font.Size * GetDeviceCaps(mobjDevicehdc, _
        LOGPIXELSY) \ 72
    
    'Set Italic or not
    If mobjDevice.Font.Italic = True Then
        lfFont.lfItalic = 1
    Else
        lfFont.lfItalic = 0
    End If
    'Set Underline or not
    If mobjDevice.Font.Underline = True Then
        lfFont.lfUnderline = 1
    Else
        lfFont.lfUnderline = 0
    End If
    'Set Strikethrough or not
    If mobjDevice.Font.Strikethrough = True Then
        lfFont.lfStrikeOut = 1
    Else
        lfFont.lfStrikeOut = 0
    End If
    'Set Bold or not (use font's weight)
    lfFont.lfWeight = mobjDevice.Font.Weight
    'Set font rotation angle
    lfFont.lfEscapement = CLng(mnAngle * 10#)
    lfFont.lfOrientation = lfFont.lfEscapement
    
    'Build temporary new font and output the string
    lFont = CreateFontIndirect(lfFont)
    lOldFont = SelectObject(mobjDevicehdc, lFont)
    lRes = TextOut(mobjDevicehdc, XtoP(mobjDeviceCurrentX), _
        YtoP(mobjDeviceCurrentY), sText, Len(sText))
    lFont = SelectObject(mobjDevicehdc, lOldFont)
    DeleteObject lFont
End Sub

'~~~Device
Property Set Device(objDevice As Object)
    Dim fSX2 As Single
    Dim fSY2 As Single
    Dim fPX2 As Single
    Dim fPY2 As Single
    Dim nScaleMode As Integer
    Set mobjDevice = objDevice
    With mobjDevice
        'Grab current scaling parameters
        nScaleMode = .ScaleMode
        mfSX1 = .ScaleLeft
        mfSY1 = .ScaleTop
        fSX2 = mfSX1 + .ScaleWidth
        fSY2 = mfSY1 + .ScaleHeight
        'Temporarily set pixels mode
       .ScaleMode = vbPixels
    '   .ScaleMode = vbMillimeters
       'Grab pixel scaling parameters
        fPX2 = .ScaleWidth
        fPY2 = .ScaleHeight
        'Reset user's original scale
        If nScaleMode = 0 Then
            mobjDevice.Scale (mfSX1, mfSY1)-(fSX2, fSY2)
        Else
            mobjDevice.ScaleMode = nScaleMode
        End If
        'Calculate scaling ratios just once
        mfXRatio = fPX2 / (fSX2 - mfSX1)
        mfYRatio = fPY2 / (fSY2 - mfSY1)
    End With
End Property

'Scales X value to pixel location
Private Function XtoP(fX As Single) As Long
    XtoP = (fX - mfSX1) * mfXRatio
End Function

'Scales Y value to pixel location
Private Function YtoP(fY As Single) As Long
    YtoP = (fY - mfSY1) * mfYRatio
End Function


