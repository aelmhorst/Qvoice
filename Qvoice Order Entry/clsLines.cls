VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Line"
Attribute VB_Ext_KEY = "Member0" ,"Line"
Option Explicit

'local variable to hold collection
Private mcol                    As Collection
Private mcolDeletedLineIDs      As New Collection
Private WithEvents mOrder    As Order
Attribute mOrder.VB_VarHelpID = -1

Public Sub Init(pobjOrder As Order)
    Set mOrder = pobjOrder
End Sub

Public Function Add(plngQuantity As Long, pcurLength As Currency, _
             pcurWidth As Currency, plngSlabID As Long, _
             plngLaminateID As Long)
   
   'create a new object
    Dim objNewMember As qkorder.Line
    Set objNewMember = New qkorder.Line
    
    'set the property passed into the method
    objNewMember.Add mOrder, (mcol.Count + 1), plngQuantity, pcurLength, _
            pcurWidth, plngSlabID, plngLaminateID
    
        mcol.Add objNewMember
Set Add = objNewMember

Set objNewMember = Nothing
End Function
Public Function Load(iSerialID As Long, rs As Recordset)
    'create a new object
    Dim objNewMember As qkorder.Line
    Set objNewMember = New qkorder.Line
    
    'set the property passed into the method
    objNewMember.Load mOrder, iSerialID, rs
    
    mcol.Add objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Line
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mcol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mcol.Count
End Property

Friend Property Get DeletedLines() As Collection
    Set DeletedLines = mcolDeletedLineIDs
End Property
Public Sub Remove(vntIndexKey As Variant)
    mcolDeletedLineIDs.Add Item(vntIndexKey).SerialID
    mcol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mcol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mcol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mcol = Nothing
    If Not mOrder Is Nothing Then Set mOrder = Nothing
End Sub

Private Sub mOrder_CustomerChanged()
Dim lLine As qkorder.Line
    For Each lLine In mcol
        lLine.Customer_Changed
    Next
End Sub

