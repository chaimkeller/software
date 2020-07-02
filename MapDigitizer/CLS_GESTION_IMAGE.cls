VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGestionImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================================================
'Nom du contrôle : DKS_GESTION_IMAGE
'Version : 1.0.2
'Auteur : Teillet nicolas
'Environnement de développement : Visual basic 6.0
'Résumé : Librairie de gestion d'image pour un accès rapide
'====================================================================================================
'====================================================================================================
'Nom du fichier : CLS_Image.cls
'Crée le : 27/03/2004
'Rôle : classe principale de la librairie contenant l'objet d'image
'====================================================================================================

'déclaration des types privées pour les API :
'structure stockant les informations des en-têtes bitmap
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'structure stockant les informations des bits d'une image
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'stucture permettant de définir un bitmap
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'fonction permettant de récupèrer les bits d'une image
Private Declare Function GetDIBits Lib "gdi32" ( _
ByVal hdc As Long, _
ByVal hBitmap As Long, _
ByVal nStartScan As Long, _
ByVal nScanLines As Long, _
lpBits As Any, _
lpBitmapInfo As BITMAPINFO, _
ByVal wUsage As Long) As Long

'fonction permettant de définir les bits d'un bitmap dans un DC
Private Declare Function SetDIBits Lib "gdi32.dll" ( _
ByVal hdc As Long, _
ByVal hBitmap As Long, _
ByVal nStartScan As Long, _
ByVal nNumScans As Long, _
ByRef lpBits As Any, _
ByRef lpBI As BITMAPINFO, _
ByVal wUsage As Long) As Long
     
'fonction permettant de détruire un Device Context
Private Declare Function DeleteDC Lib "gdi32" ( _
ByVal hdc As Long) As Long

'fonction permettant de détruire l'objet sélectionné pour le tracé
Private Declare Function DeleteObject Lib "gdi32" ( _
ByVal hObject As Long) As Long

'déclaration des constantes privées pour les API :
Const BI_RGB = 0 'définit le type de compression d'un bitmap
Const DIB_RGB_COLORS = 0 'table de couleur en RGB
Private Const pixR As Integer = 3 'index du pixel de couleur rouge
Private Const pixG As Integer = 2 'index du pixel de couleur vert
Private Const pixB As Integer = 1 'index du pixel de couleur bleu

'déclaration des variables membres privées à l'objet
Private m_byImage() As Byte 'stocke le tableau de bits associés à l'image
Private m_pctPictureBox As PictureBox 'stocke le hDC de l'image
Private m_pBitmapInfo As BITMAPINFO 'stocke les info du bitmap associé à l'image


Public Sub Refresh()

'===========================================================================
'Permet de coller l'image stockée dans l'objet
'===========================================================================

'on active la routine de traitement d'erreur
On Error Resume Next
'l'erreur etant minime, on continue l'éxécution normalement
                
    ' Display the result.
    Call SetDIBits(m_pctPictureBox.hdc, m_pctPictureBox.Image, 0, m_pctPictureBox.ScaleHeight, m_byImage(1, 1, 1), m_pBitmapInfo, DIB_RGB_COLORS)
    Set m_pctPictureBox.Picture = m_pctPictureBox.Image

End Sub

Private Function BlueValue(ByVal lColor As Long) As Integer

'==================================================================
'Permet d'obtenir la valeur bleu d'une couleur RGB spécifiée
'
'COL_Couleur : Couleur de type RGB
'
'Renvoie la valeur bleu de la couleur envoyé sous forme d'un entier
'Renvoie "-1" en cas d'erreur
'===================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux du milieu
    Let BlueValue = ((lColor \ 65536) And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de bleu
'Solution : on renvoie une valeur d'erreur : "-1"

Let BlueValue = -1

End Function

Private Sub ExtractBits()

'===========================================================================
'Permet d'extraire le tableau de bits de l'image
'===========================================================================

'on active la routine de traitement d'erreur
On Error Resume Next
'l'erreur etant minime, on continue l'éxécution normalement

'    '--------------diagnostics------------------
'    ier = RecordSimpleBits
'    DigiGDIfailed = True
'    Exit Sub
'    '------------------------------------------

    ' Prepare the bitmap description.
    With m_pBitmapInfo.bmiHeader
        Let .biSize = Len(m_pBitmapInfo.bmiHeader)
        Let .biWidth = m_pctPictureBox.ScaleWidth
        Let .biHeight = -m_pctPictureBox.ScaleHeight
        Let .biPlanes = 1
        Let .biBitCount = 32
        Let .biCompression = BI_RGB
        Let .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) - (((.biWidth * .biBitCount) + 7) \ 8) * Abs(.biHeight)
    End With

    ' Load the bitmap's data using DIBs.
    GDMDIform.Toolbar1.Enabled = False
    GDMDIform.Toolbar1.Buttons(1).Enabled = False
    GDMDIform.Toolbar1.Buttons(2).Enabled = False
    
    ReDim m_byImage(1 To 4, 1 To Int(m_pctPictureBox.ScaleWidth / twipsx), 1 To Int(m_pctPictureBox.ScaleHeight / twipsy))
    ier = GetDIBits(m_pctPictureBox.hdc, m_pctPictureBox.Image, 0, m_pctPictureBox.ScaleHeight, m_byImage(1, 1, 1), m_pBitmapInfo, DIB_RGB_COLORS)
    If val(ier) = 0 And IgnoreAutoRedrawError% <> 1 Then 'DIBS failed
       'do it using simple API
       DigiGDIfailed = True 'use VB6 Point to record the colors before drawing a contour
       ReDim m_byImage(0) 'reclaim memory
       ier = RecordSimpleBits 'record the pixel rgb colors to a file residing on the hard disk
'       Call MsgBox("GDI failed mapping the colors probably since the chosen map is too big...." _
'                   & vbCrLf & "" _
'                   & vbCrLf & "If you want to continue with this map, some features of this program" _
'                   & vbCrLf & "will be disabled including erasing digitizing errors." _
'                   , vbExclamation Or vbDefaultButton1, "GDI Error")
'       GDMDIform.StatusBar1.Panels(1).Text = "GDI failed mapping the colors probably since the map is too big..."
'       GDMDIform.Toolbar1.Buttons(40).Enabled = False
       End If
       
    'now record byte array if not already exists, one byte for each pixel, to be used for editing the digitizing
    ier = RecordByteImage
    
    GDMDIform.Toolbar1.Enabled = True
    GDMDIform.Toolbar1.Buttons(1).Enabled = True
    GDMDIform.Toolbar1.Buttons(2).Enabled = True
    
    If ier < 0 Then
      Call MsgBox("Buffer file wasn't successfully created." _
                  & vbCrLf & "Some editing options will not be available." _
                  & vbCrLf & "" _
                  & vbCrLf & "(Hint: check if you are running the programming with" _
                  & vbCrLf & "administrator permission)" _
                  & vbCrLf & "" _
                  , vbExclamation, "File permission error")
      End If
    
End Sub

Public Function GetPixelN(ByVal lX As Long, ByVal lY As Long) As Long

'=============================================================================================
'Permet de récupèrer la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut récupèrer la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut récupèrer la couleur
'
'Renvoie la couleur du point en RGB
'Renvoie -1 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    If lX <= 0 Or lY <= 0 Or lX >= m_pctPictureBox.ScaleWidth Or lY >= m_pctPictureBox.ScaleHeight Then
    
        Let GetPixelN = -1
        Exit Function
    
    End If

    'on renvoie la couleur correspondante
    Let GetPixelN = rgb(m_byImage(pixR, lX + 1, lY + 1), m_byImage(pixG, lX + 1, lY + 1), m_byImage(pixB, lX + 1, lY + 1))
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On ne connaît pas la couleur du point
'Solution : On renvoie une valeur d'erreur : -1

Let GetPixelN = -1

End Function

Public Function GetPixelRGB(ByVal lX As Long, ByVal lY As Long, ByRef byRed As Byte, ByRef byGreen As Byte, ByRef byBlue As Byte) As Long

'=============================================================================================
'Permet de récupèrer la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut récupèrer la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut récupèrer la couleur
'BYT_Red : Composante rouge de la couleur
'BYT_Green : Composante verte de la couleur
'BYT_Blue : Composante bleue de la couleur
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    If lX < 0 Or lY < 0 Or lX > m_pctPictureBox.ScaleWidth Or lY > m_pctPictureBox.ScaleHeight Then
    
        Let GetPixelRGB = -1
        Exit Function
    
    End If

    'on renvoie la couleur correspondante
    Let byRed = m_byImage(pixR, lX + 1, lY + 1)
    Let byGreen = m_byImage(pixG, lX + 1, lY + 1)
    Let byBlue = m_byImage(pixB, lX + 1, lY + 1)
    Let GetPixelRGB = 0
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On ne connaît pas la couleur du point
'Solution : On renvoie une valeur d'erreur : 0

Let GetPixelRGB = -1
GDMDIform.StatusBar1.Panels(1).Text = "Error: can't erase since picture is too big and GDI failed"

End Function

Private Function GreenValue(ByVal lColor As Long) As Integer

'===================================================================
'Permet d'obtenir la valeur vert d'une couleur RGB
'
'lColor : Couleur de type RGB
'
'Renvoie la valeur vert de la couleur envoyé sous forme d'un entier
'Renvoie "-1" en cas d'erreur
'===================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux du milieu à côté de ceux du bleu
    Let GreenValue = ((lColor \ 256) And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de vert
'Solution : on renvoie une valeur d'erreur : "-1"

GreenValue = -1

End Function

Public Property Set PictureBox(ByRef pctPictureBox As PictureBox)

'==============================================================================================
'Permet de définir le hDC d'une image stockée dans l'objet ainsi que les dimensions de l'image
'==============================================================================================

    Set m_pctPictureBox = pctPictureBox
    
    'on extrait les bits de l'image
    Call ExtractBits
    
End Property

Private Function RedValue(ByVal lColor As Long) As Integer

'===================================================================
'Permet d'obtenir la valeur rouge d'une couleur RGB spécifiée
'
'COL_Couleur : Couleur de type RGB
'
'Revoie la valeur rouge de la couleur envoyé sous forme d'un entier
'Renvoie -1 en cas d'erreur
'====================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux de droite
    Let RedValue = (lColor And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de rouge
'Solution : on renvoie une valeur d'erreur : "-1"

Let RedValue = -1

End Function

Public Function SetPixelN(ByVal lX As Long, ByVal lY As Long, ByVal lColor As OLE_COLOR) As Long

'=============================================================================================
'Permet de définir la couleur d'un pixel de l'image
'
'lX : Coordonnée X en pixel du point de l'image dont on veut définir la couleur
'lY : Coordonnée Y en pixel du point de l'image dont on veut définir la couleur
'lColor : Couleur à affecter au pixel
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    If lX <= 0 Or lY <= 0 Or lX >= m_pctPictureBox.ScaleWidth Or lY >= m_pctPictureBox.ScaleHeight Then
    
        Let SetPixelN = 0
        Exit Function
    
    End If

    'on applique la couleur du pixel
    Let m_byImage(pixR, lX + 1, lY + 1) = RedValue(lColor)
    Let m_byImage(pixG, lX + 1, lY + 1) = GreenValue(lColor)
    Let m_byImage(pixB, lX + 1, lY + 1) = BlueValue(lColor)

'la fonction a réussie
SetPixelN = 1
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu définir la couleur du pixel
'Solution : On renvoie une valeur d'erreur : 0

Let SetPixelN = 0

End Function

Public Function SetPixelRGB(ByVal lX As Long, ByVal lY As Long, ByVal byRed As Byte, ByVal byGreen As Byte, ByVal byBlue As Byte) As Long

'=============================================================================================
'Permet de définir la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut définir la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut définir la couleur
'BYT_Red : composante rouge du pixel à appliquer
'BYT_Green : composante verte du pixel à appliquer
'BYT_Blue : composante bleue du pixel à appliquer
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
On Error GoTo erreur0

    If lX <= 0 Or lY <= 0 Or lX >= m_pctPictureBox.ScaleWidth Or lY >= m_pctPictureBox.ScaleHeight Then
    
        Let SetPixelRGB = -1
        Exit Function
    
    End If

    'on applique la couleur du pixel
    Let m_byImage(pixR, lX + 1, lY + 1) = byRed
    Let m_byImage(pixG, lX + 1, lY + 1) = byGreen
    Let m_byImage(pixB, lX + 1, lY + 1) = byBlue

'la fonction a réussie
Let SetPixelRGB = 1
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu définir la couleur du pixel
'Solution : On renvoie une valeur d'erreur : 0

Let SetPixelRGB = 0

End Function

Private Sub Class_Initialize()

    'on redimensionne le tableau de bits de l'image
    ReDim m_byImage(0)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : RecordSimpleBits
' Author    : Dr-John-K-Hall
' Date      : 2/24/2015
' Purpose   : Write binary file containing the pixel colors
'---------------------------------------------------------------------------------------
'
Private Function RecordSimpleBits() As Integer
   On Error GoTo RecordSimpleBits_Error

   Dim i As Long, j As Long
   Dim R As Long
   Dim blue As Double
   Dim green As Double
   Dim red As Double
   
   Dim ier As Integer
   
   Dim NewColorEnum As ColorEnum
   oldFileBytes = -1
   
    ier = 0
    
'   pos% = InStr(picnam$, ".")
'   picext$ = Mid$(picnam$, pos% + 1, 3)
   picfile$ = App.Path & "\" & RootName(picnam$) & "-PIC" & ".bin"
    
   If Not DigiPicFileOpened Then 'check if file is complete, and open for reading
      If Dir(picfile$) <> sEmpty Then
         Picfilnum% = FreeFile
         Open picfile$ For Random Access Read As #Picfilnum% Len = Len(NewColorEnum)
         'try reading the last pixel, if doesn't succeed reload the data
         ier = TestPicFile
         If ier = 0 Then 'file is OK, otherwise read it over again
            DigiPicFileOpened = True
            RecordSimpleBits = ier
            Exit Function
            End If
         End If
   Else 'file was opened once and for some reason not closed
      If Picfilnum% > 0 Then
         'do the read test
         ier = TestPicFile
         If ier = 0 Then 'file is OK, otherwise read it over again
            DigiPicFileOpened = True
            RecordSimpleBits = ier
            Exit Function
            End If
      Else
         Picfilnum% = 0
         Close #Picfilnum%
         DigiPicFileOpened = False
         End If
      End If
      
   Call MsgBox("The map file is too big to use the fastest means of recording the pixels." _
               & vbCrLf & "Instead, a image of the file needs to be written onto the hard disk." _
               & vbCrLf & "This might take a bit of time, please wait......." _
               , vbInformation, "Storing pixel information")
               
   Picfilnum% = FreeFile
   Open picfile$ For Random Access Write As #Picfilnum% Len = Len(NewColorEnum)
             
   With GDMDIform
      '------fancy progress bar settings---------
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
   End With
   pbScaleWidth = 100

   Call UpdateStatus(GDMDIform, 1, 0)
   GDMDIform.StatusBar1.Panels(1).Text = "Storing pixel colors, please wait....this could take a while...."
   
   For j = 1 To m_pctPictureBox.ScaleHeight
   
       For i = 1 To m_pctPictureBox.ScaleWidth
          
           R = GetPixel(m_pctPictureBox.hdc, i, j)
           If R <> -1 Then
              'inline recupcouleur to make it faster
              blue = Fix((R / 256) / 256)
              green = Fix((R - ((blue * 256) * 256)) / 256)
              red = Fix(R - ((blue * 256) * 256) - (green * 256))

              NewColorEnum.RedColor = Abs(red)
              NewColorEnum.GreenColor = Abs(green)
              NewColorEnum.BlueColor = Abs(blue)
              
           Else
              'make it blank, i.e., white
              NewColorEnum.RedColor = 255
              NewColorEnum.GreenColor = 255
              NewColorEnum.BlueColor = 255
              End If
              
           Put #Picfilnum%, , NewColorEnum
           DoEvents
       Next i
       Call UpdateStatus(GDMDIform, 1, 100 * j / m_pctPictureBox.ScaleHeight)
       DoEvents
   Next j
   
   Close #Picfilnum%
   'now reopen for reading
   Picfilnum% = FreeFile
   Open picfile$ For Random Access Read As #Picfilnum% Len = Len(NewColorEnum)
   DigiPicFileOpened = True
   
   GDMDIform.picProgBar.Visible = False
   GDMDIform.StatusBar1.Panels(1) = sEmpty
   GDMDIform.StatusBar1.Panels(2) = sEmpty
   
   RecordSimpleBits = ier
   
   On Error GoTo 0
   Exit Function

RecordSimpleBits_Error:

    ier = -1
    GDMDIform.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(1) = sEmpty
    ReDim m_byImage(0)
    RecordSimpleBits = ier
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecordSimpleBits of Class Module CGestionImage"

End Function
'---------------------------------------------------------------------------------------
' Procedure : RecordByteImage
' Author    : Dr-John-K-Hall
' Date      : 4/28/2015
' Purpose   : record blank array the size of the picture
'---------------------------------------------------------------------------------------
'
Private Function RecordByteImage() As Integer

   On Error GoTo RecordByteImage_Error

   Dim PointImage As Byte, ier As Integer, pos%
   Dim ImageFile$, picext$, TmpImageFile$
   Dim Byte0 As Byte, Byte1 As Byte
   Dim i As Long, j As Long
   Dim Xpix As Long, Ypix As Long
   Dim RecNum&
   
   On Error GoTo RecordByteImage_Error
   
   Byte0 = 0
   Byte1 = 1
   
   ier% = 0
   
'   pos% = InStr(picnam$, ".")
'   picext$ = Mid$(picnam$, pos% + 1, 3)
   ImageFile$ = App.Path & "\" & RootName(picnam$) & "-IMG" & ".bin"
   TmpImageFile$ = App.Path & "\" & RootName(picnam$) & "-IMG" & ".buf"
   
   If Dir(ImageFile$) = sEmpty Then
   
       filnumImage% = FreeFile
       Open ImageFile$ For Random Access Write As #filnumImage% Len = Len(PointImage)
       
       With GDMDIform
          '------fancy progress bar settings---------
          .picProgBar.AutoRedraw = True
          .picProgBar.BackColor = &H8000000B 'light grey
          .picProgBar.DrawMode = 10
        
          .picProgBar.FillStyle = 0
          .picProgBar.ForeColor = &H400000 'dark blue
          .picProgBar.Visible = True
       End With
       pbScaleWidth = 100
    
       Call UpdateStatus(GDMDIform, 1, 0)
       GDMDIform.StatusBar1.Panels(1).Text = "Creating point digitizing buffer, please wait..."
       
       For j = 1 To pixhi
       
           For i = 1 To pixwi
           
               RecNum& = i + (j - 1) * pixwi
           
               Put #filnumImage%, RecNum&, Byte0
               
    '           DoEvents
           Next i
           Call UpdateStatus(GDMDIform, 1, 100 * j / pixhi)
           DoEvents
       Next j
               
       Close #filnumImage%
   
       GDMDIform.picProgBar.Visible = False
       GDMDIform.StatusBar1.Panels(1) = sEmpty
       GDMDIform.StatusBar1.Panels(2) = sEmpty
       
       End If
   
   'now copy this file to the temporary buffer
   If Dir(ImageFile$) <> sEmpty Then
   
      'remove old buffer file if it exists
      If Dir(TmpImageFile$) <> sEmpty Then
         Kill TmpImageFile$
         End If
         
      'copy blank buffer template to buffer file
      FileCopy ImageFile$, TmpImageFile$
   
     'open buffer file for future read and write operations
      filnumImage% = FreeFile
      Open TmpImageFile$ For Random Access Read Write As #filnumImage% Len = Len(PointImage)
    
      'flag the program that the buffer file is available for read and write
      ImagePointFile = True
      
      ier = 0
      
   Else
      
      ier = -1 'error code, buffer file couldn't be created
      
      End If
      
   RecordByteImage = ier

   On Error GoTo 0
   Exit Function

RecordByteImage_Error:
    ier = -1
    RecordByteImage = ier
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecordByteImage of Class Module CGestionImage"
End Function
