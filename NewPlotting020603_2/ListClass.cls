VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                                Eric de Decker secr VB Belgie
'                                E.B. our API Killer  ProFinance Woerden
'                                Willem secr VBgroup Ned vbg@vbgroup.nl

'This sources may be used freely without the intention of commercial distribution. For all
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.


Public Text As String
'local variable(s) to hold property value(s)
Private mvarItemForeColor As Long '<<<change
Private mvarItemBackColor As Long '<<<change
Private mvarUnderlayingValue As Variant 'local copy
Private mvarEnabeled As Boolean 'local copy
Private mvarToolTipText As String 'local copy
Private mvarSelected As Boolean
Private mvarPicture As Picture
Private mvarDownPicture As Picture
Private mvarVisible As Boolean


Public Property Let Visible(ByVal bVisible As Boolean)
mvarVisible = bVisible
End Property

Public Property Get Visible() As Boolean
Visible = mvarVisible
End Property

Public Property Let DownPicture(ByVal oData As Picture)
Set mvarDownPicture = oData
End Property

Public Property Get DownPicture() As Picture
Set DownPicture = mvarDownPicture
End Property

Public Property Let ItemPicture(ByVal oData As Picture)
Set mvarPicture = oData
End Property

Public Property Get ItemPicture() As Picture
Set ItemPicture = mvarPicture
End Property

Public Property Let ToolTipText(ByVal sData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToolTipText = 5
    mvarToolTipText = sData
End Property

Public Property Get ToolTipText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToolTipText
    ToolTipText = mvarToolTipText
End Property

Public Property Let Selected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Selected = True
    mvarSelected = vData
End Property


Public Property Get Selected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Selected
    Selected = mvarSelected
End Property

Public Property Let Enabeled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Enabeled = 5
    mvarEnabeled = vData
    'If Not vData Then mvarSelected = False
End Property


Public Property Get Enabeled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Enabeled
    Enabeled = mvarEnabeled
End Property

Public Property Let UnderlayingValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnderlayingValue = 5
    mvarUnderlayingValue = vData
End Property

Public Property Get UnderlayingValue() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnderlayingValue
   UnderlayingValue = mvarUnderlayingValue
End Property

Public Property Let ItemForeColor(ByVal lData As Long) '<<<change
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemForeColor = RGB(R,G,B)
    mvarItemForeColor = lData
End Property

Public Property Get ItemForeColor() As Long '<<<change
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemForeColor
    ItemForeColor = mvarItemForeColor
End Property

Public Property Let ItemBackColor(ByVal lData As Long) '<<<change
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemBackColor = RGB(R,G,B)
    mvarItemBackColor = lData
End Property

Public Property Get ItemBackColor() As Long '<<<change
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemBackColor
    ItemBackColor = mvarItemBackColor
End Property



