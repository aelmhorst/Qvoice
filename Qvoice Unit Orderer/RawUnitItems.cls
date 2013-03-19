VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RawUnitItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RawUnitItem"
Attribute VB_Ext_KEY = "Member0" ,"RawUnitItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function AddNew(Item As OrderItem, Optional sKey As String, Optional Before As Integer = -1) As RawUnitItem
    'create a new object
    Dim objNewMember As RawUnitItem
    Set objNewMember = New RawUnitItem

    'set the properties passed into the method
    objNewMember.Init Item
    
    
    If Len(sKey) = 0 Then
        AddItem objNewMember, Before
    Else
        AddItemWithKey objNewMember, Before, sKey
    End If


    'return the object created
    Set AddNew = objNewMember
    Set objNewMember = Nothing


End Function

Public Sub Add(Item As RawUnitItem, Optional sKey As String, Optional Before As Integer = -1)
    'create a new object
 
    
    
    If Len(sKey) = 0 Then
        AddItem Item, Before
    Else
        AddItemWithKey Item, Before, sKey
    End If



End Sub

Private Sub AddItem(Item As RawUnitItem, Before As Integer)
    If Before > -1 Then
        mCol.Add Item, Before:=Before
    Else
        mCol.Add Item
    End If
End Sub

Private Sub AddItemWithKey(Item As RawUnitItem, Before As Integer, sKey As String)
    If Before > -1 Then
        mCol.Add Item, sKey, Before:=Before
    Else
        mCol.Add Item, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As RawUnitItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
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

