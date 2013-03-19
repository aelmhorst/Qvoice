VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Charges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Charge"
Attribute VB_Ext_KEY = "Member0" ,"Charge"
Option Explicit

'local variable to hold collection
Private mcol As Collection
Public Sub Add(plngChargeID As Long, _
             pstrAbbrev As String, pstrDesc As String, _
             pcurPrice As Currency, plngQuantity As Long, _
             pbolLineal As Boolean, _
            pcurLengthAdj As Currency, Optional plngLineChargeID As Long = 0)
            
    'create a new object
    Dim objNewMember As Charge
    Set objNewMember = New Charge


    'set the properties passed into the method
    objNewMember.Init plngChargeID, plngQuantity, _
        pstrAbbrev, pcurPrice, pstrDesc, pbolLineal, pcurLengthAdj, plngLineChargeID
    
    mcol.Add objNewMember
   
    Set objNewMember = Nothing
End Sub

Public Property Get Item(vntIndexKey As Variant) As Charge
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


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mcol.Remove vntIndexKey
End Sub

Public Property Get ExtPrice() As Currency

End Property


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
End Sub
Friend Sub Clear()
    Do While mcol.Count > 0
     mcol.Remove 1
    Loop
End Sub

