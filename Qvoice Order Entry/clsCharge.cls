VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Charge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mbolChanged         As Boolean
Private mbolLineal          As Boolean
Private mlngLineChargeID    As Long
Private mlngChargeID        As Long
Private mstrChargeAbbrev    As String
Private mstrChargeDesc      As String
Private mcurChargePrice     As Currency
Private mcurAdjustment      As Currency
Private mlngQuantity        As Long

Public Property Get LengthAdjustment() As Currency
    LengthAdjustment = mcurAdjustment
End Property
Public Property Get iChargeID() As Long
    iChargeID = mlngChargeID
End Property
Public Property Get TotalCharge() As Currency
    TotalCharge = Price * Quantity
End Property
Public Sub Init(plngChargeID As Long, _
            plngQuantity As Long, pstrAbbrev As String, _
            pcurPrice As Currency, pstrDesc As String, _
            pbolLineal As Boolean, pcurLengthAdjustment As Currency, _
            Optional plngChargeLineID As Long = 0)

mbolLineal = pbolLineal
mlngChargeID = plngChargeID
mlngQuantity = plngQuantity
mstrChargeAbbrev = pstrAbbrev
mstrChargeDesc = pstrDesc
mcurChargePrice = pcurPrice
mcurAdjustment = pcurLengthAdjustment

If plngChargeLineID = 0 Then
    mbolChanged = True
Else
    mlngLineChargeID = plngChargeLineID
    mbolChanged = False
End If

End Sub

Public Property Get iLineChargeID() As Long
    iLineChargeID = mlngLineChargeID
End Property
Public Property Let iLineChargeID(plngChargeID As Long)
    mlngLineChargeID = plngChargeID
End Property

Public Property Get Quantity() As Long
    Quantity = mlngQuantity
End Property
Public Property Let Quantity(plngQuantity As Long)
If mlngQuantity <> plngQuantity Then
    mlngQuantity = plngQuantity
    mbolChanged = True
End If
End Property

Public Property Get Abbrev() As String
    Abbrev = mstrChargeAbbrev
End Property
Public Property Get Description() As String
    Description = mstrChargeDesc
End Property
Public Property Get Price() As Single
    Price = mcurChargePrice
End Property

Public Property Get Lineal() As Boolean
    Lineal = mbolLineal
End Property
