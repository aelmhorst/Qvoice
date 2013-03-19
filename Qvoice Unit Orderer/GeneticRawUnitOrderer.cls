VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneticRawUnitOrderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim m_MaxLength As Currency

Public Function OrderUnits(in_Items() As OrderItem, in_Units() As OrderableUnit) As RawUnitItems

Dim a_Counter As Integer, aUbound As Integer, a_SlabCandidatePos As Integer, aOriginalArrayLength As Integer
Dim a_Items As New RawUnitItems

Dim a_CurrentCost As Currency
Dim a_CurrentTightness As Currency

Dim a_CandidateCost As Currency
Dim a_CandidateTightness As Currency

Dim a_Finished As Boolean
Dim a_IsComplete As Boolean

Dim a_SlabItem As OrderableUnit
Dim a_Item As OrderItem

'// Get the Max Length ITem
m_MaxLength = 0
aUbound = UBound(in_Units)
For a_Counter = 1 To aUbound
    If in_Units(a_Counter).LengthAllowedInInches > m_MaxLength Then m_MaxLength = in_Units(a_Counter).LengthAllowedInInches
Next

'// Find out if any slabs are too long
aUbound = UBound(in_Items)
aOriginalArrayLength = aUbound

a_Counter = 1
Do
    If in_Items(a_Counter).Length > m_MaxLength Then
        ' This is pretty inefficient, but this code won't be hit too much
        aUbound = aUbound + 1
        ReDim Preserve in_Items(1 To aUbound)
        Set a_Item = New OrderItem
        a_Item.Length = in_Items(a_Counter).Length - m_MaxLength
        If a_Item.Length > m_MaxLength Then a_Item.Length = m_MaxLength
        a_Item.SerialID = in_Items(a_Counter).SerialID
        a_Item.ItemNumber = in_Items(a_Counter).ItemNumber
        Set in_Items(aUbound) = a_Item
        in_Items(a_Counter).Length = in_Items(a_Counter).Length - a_Item.Length
    Else
        a_Counter = a_Counter + 1
    End If
    If a_Counter > aOriginalArrayLength Then Exit Do
Loop



SelectionSort in_Items

g_OrderableItems = in_Units

' Setup our original collection. This one is sorted
For a_Counter = 1 To aUbound
    a_Items.AddNew in_Items(a_Counter)
Next


While a_Finished = False
    a_SlabCandidatePos = 0
    a_CandidateCost = 100000000@
    a_CandidateTightness = 32767
    
    For a_Counter = 2 To a_Items.Count
    
        EvaluateCombinationWithX a_Items, a_Counter, a_CurrentCost, a_CurrentTightness, a_IsComplete
        If Not a_IsComplete Then
            If a_CurrentCost < a_CandidateCost Then
                a_SlabCandidatePos = a_Counter
                a_CandidateCost = a_CurrentCost
                a_CandidateTightness = a_CurrentTightness
            ElseIf a_CurrentCost = a_CandidateCost And a_CurrentTightness < a_CandidateTightness Then
                a_SlabCandidatePos = a_Counter
                a_CandidateCost = a_CurrentCost
                a_CandidateTightness = a_CurrentTightness
            End If
        End If
        
    Next
    If a_SlabCandidatePos > 0 Then
        MoveItem a_SlabCandidatePos, a_Items
    Else
        a_Finished = True
    End If
    Debug.Print a_Items.Count
Wend
Set OrderUnits = a_Items


End Function

Private Sub EvaluateCombinationWithX(in_Items As RawUnitItems, in_CurrentPos As Integer, ByRef out_Cost As Currency, _
                            ByRef out_Tightness As Currency, ByRef out_IsComplete As Boolean)
    Dim a_PrimaryItem As RawUnitItem
    Dim a_CandidateItem As New RawUnitItem
    Dim a_PrimaryOrderItem As New OrderItem
    
    Dim a_Counter As Integer
    
    out_IsComplete = False
    out_Cost = 0@
    
    Set a_PrimaryItem = in_Items(1)
    a_PrimaryOrderItem.Length = a_PrimaryItem.Length
    
    a_CandidateItem.Init a_PrimaryOrderItem
    
    a_CandidateItem.AppendChild in_Items(in_CurrentPos)
    If a_CandidateItem.Length > m_MaxLength Then
        out_IsComplete = True
        Exit Sub
    End If
    out_Cost = a_CandidateItem.Cost
    out_Tightness = a_CandidateItem.Tightness
    For a_Counter = 2 To in_Items.Count
        If a_Counter <> in_CurrentPos Then
            Set a_CandidateItem = in_Items(a_Counter)
            out_Cost = out_Cost + a_CandidateItem.Cost
        End If
    Next
End Sub

'// Specifies the amount that should be factored into a raw item's length if it is split in half
'// Allows for adjustments based on blade
Public Property Let ItemSeperationFactor(inFactor As Currency)
    g_ItemSeperationFactor = inFactor
End Property



Private Sub MoveItem(in_Index As Integer, in_Collection As RawUnitItems)
    
   'Move the item from position 1 to the desired index
    Dim a_TargetItem As RawUnitItem
    Set a_TargetItem = in_Collection(1)
    a_TargetItem.AppendChild in_Collection(in_Index)
    in_Collection.Remove in_Index
    in_Collection.Remove 1
    
    Dim a As Integer
    Dim aItem As RawUnitItem
   
    For a = 1 To in_Collection.Count
        Set aItem = in_Collection(a)
        If a_TargetItem.Length < aItem.Length Then
            Exit For
        End If
    Next
    
    If a < in_Collection.Count Then
        in_Collection.Add a_TargetItem, Before:=a
    Else
        in_Collection.Add a_TargetItem
    End If
    
End Sub

