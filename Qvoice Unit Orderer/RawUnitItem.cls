VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RawUnitItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private m_SlabIndex As Integer
Private m_Child As RawUnitItem
Private m_InitialItem As OrderItem

Public Sub Init(in_InitialItem As OrderItem)
    Set m_InitialItem = in_InitialItem
End Sub

Public Property Get OrderableUnit() As OrderableUnit
    If m_SlabIndex = -1 Then m_SlabIndex = ModAlgorithms.GetSlabIndex(Me.Length)
    Set OrderableUnit = g_OrderableItems(m_SlabIndex)
End Property

Public Property Get InitialLength() As Currency
    InitialLength = m_InitialItem.Length
End Property

Public Property Get SerialID() As Long
    SerialID = m_InitialItem.SerialID
End Property

Public Property Get ItemNumber() As Byte
    ItemNumber = m_InitialItem.ItemNumber
End Property

Public Property Get OrderItem() As OrderItem
    Set OrderItem = m_InitialItem
End Property

Public Property Get Cost() As Currency
    Cost = OrderableUnit.Cost
End Property

Public Property Get Tightness() As Currency
    Tightness = OrderableUnit.LengthAllowedInInches - Me.Length
End Property

Public Property Get Length() As Currency
    If m_Child Is Nothing Then
        Length = InitialLength
    Else
        '// TODO: Determine if we should fudge on an extra 1/8th of an inch for each additional slab?
        '// For now, we will do an extra 1/16th
        Length = InitialLength + m_Child.Length + ModAlgorithms.g_ItemSeperationFactor
    End If
End Property

Public Sub AppendChild(in_Child As RawUnitItem)
If m_Child Is Nothing Then
    Set m_Child = in_Child
Else
    m_Child.AppendChild in_Child
End If
    m_SlabIndex = -1
End Sub

Public Sub ReplaceChild(in_Child As RawUnitItem)
    Set m_Child = in_Child
    m_SlabIndex = -1
End Sub


Public Property Get Child() As RawUnitItem
    If m_Child Is Nothing Then
        Set Child = Nothing
    Else
        Set Child = m_Child
    End If
End Property


Private Sub Class_Initialize()
  m_SlabIndex = -1
End Sub

Private Sub Class_Terminate()
    If Not m_Child Is Nothing Then
        Set m_Child = Nothing
    End If
    If Not m_InitialItem Is Nothing Then
        Set m_InitialItem = Nothing
    End If
        
End Sub
