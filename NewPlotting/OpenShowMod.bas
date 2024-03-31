Attribute VB_Name = "OpenShowMod"

'This module replacess the VB CommonDialog control with a
'MultiSelect GetOpenFileName Common Dialog API
'(This code is from "www.mvps.org/vbnet/code/")

Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000&
Public Const OFN_NOTESFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFS_MAXPATHNAME As Long = 260

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS
             
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY
             
Public Type OPENFILENAME
   nStructSize As Long
   hWndOwner As Long
   hInstance As Long
   sFilter As String
   sCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   sFile As String
   nMaxFile As Long
   sFileTitle As String
   nMaxTitle As Long
   sInitialDir As String
   sDialogTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   sDefFileExt As String
   nCustData As Long
   fnHook As Long
   sTemplateName As String
End Type

Public OFN As OPENFILENAME

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

'API declarations for keystroke handling
Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Public Const VK_DOWN = &H28
'Public Const VK_UP = &H26
'Public Const KEYEVENTF_KEYUP = &H2
'Public Const VK_TAB = &H9



Public Function StripDelimitedItem(startStrg As String, _
                                    delimiter As String) As String
    'This function takes a string separated by nulls,
    'splilts off 1 item, and shorten the string
    'so that the next item is ready for removal
    
    Dim pos As Long
    Dim item As String
                                    
    pos = InStr(1, startStrg, delimiter)
    If pos Then
       StripDelimitedItem = Mid$(startStrg, 1, pos)
       startStrg = Mid$(startStrg, pos + 1, Len(startStrg))
       
    End If
    
                                    
End Function

Public Function TrimNull(item As String) As String
   'removes a null from the end of the string
   
   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   If pos Then
      TrimNull = Left$(item, pos - 1)
   Else
      TrimNull = item
   End If
End Function

Public Function FileRoot(item As String) As String

    'return the file name without the directory information
    Dim Prefixs() As String
    
    Prefixs = Split(item, "\")
    If UBound(Prefixs) > 0 Then
        FileRoot = Prefixs(UBound(Prefixs))
        End If
        
    'now check that there are no extraneous characters at the end
    pos% = InStr(Len(FileRoot) - 5, FileRoot, ".")
    If Len(FileRoot) - pos% > 3 Then
       FileRoot = Mid$(FileRoot, 1, Len(FileRoot) - 1)
       End If
    
End Function
Public Function RemoveTermination(item As String) As String
    pos% = InStr(item, ".")
    If Len(item) - pos% > 3 Then
       RemoveTermination = Mid$(item, 1, Len(item) - 1)
       End If
End Function

Public Function RootDir(item As String) As String

    'return the root direcotry information
    Dim Prefixs() As String
    
    Prefixs = Split(item, "\")
    If UBound(Prefixs) > 0 Then
        RootDir = Prefixs(0)
        End If
    
End Function

