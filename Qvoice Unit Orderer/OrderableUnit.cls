VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderableUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mLengthAllowedInInches As Currency
Private mSlabLength As Integer
Private mPFC As Currency

'// The quasi-constructor for this class
'// must be called before use
Public Sub OrderableUnitItem(inSlabLength As Integer, inLengthAllowedInInches As Currency, inPFC As Currency)
    mSlabLength = inSlabLength
    mLengthAllowedInInches = inLengthAllowedInInches
    mPFC = inPFC
End Sub

Public Property Get LengthAllowedInInches() As Currency
    LengthAllowedInInches = mLengthAllowedInInches
End Property

Public Property Get SlabLength() As Integer
    SlabLength = mSlabLength
End Property

Public Property Get Cost() As Currency
    Cost = mPFC * SlabLength
End Property

