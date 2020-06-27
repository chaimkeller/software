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
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpstring As String, ByVal nCount As Long) As Long
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
