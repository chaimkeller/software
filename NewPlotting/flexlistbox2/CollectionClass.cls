VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Class1"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                                Eric de Decker secr VB Belgie
'                                E.B. our API Killer  ProFinance Woerden
'                                Willem secr VBgroup Ned vbg@vbgroup.nl

'This sources may be used freely without the intention of commercial distribution. For all
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.


'local variable to hold collection
Private mCol As Collection

Friend Function Add(key As String, Optional sKey As String, Optional lBefore As Long) As List
    'create a new object
    Dim objNewMember As List
    Set objNewMember = New List

    'set the properties passed into the method
    objNewMember.Text = key
    If Len(sKey) = 0 Then
      If lBefore Then
        mCol.Add objNewMember, , lBefore
      Else
         mCol.Add objNewMember
      End If
    Else
      If lBefore Then
        mCol.Add objNewMember, sKey, lBefore
      Else
         mCol.Add objNewMember, sKey
      End If
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function


Public Property Get item(vntIndexKey As Variant) As List
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  Set item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Friend Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

