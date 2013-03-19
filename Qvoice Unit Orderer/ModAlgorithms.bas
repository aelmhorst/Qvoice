Attribute VB_Name = "ModAlgorithms"
Option Explicit

Public g_OrderableItems() As OrderableUnit
Public g_ItemSeperationFactor As Currency


Public Function GetSlabIndex(in_LengthInInches As Currency) As Integer
    Dim aUbound As Integer
    Dim aCounter As Integer
    Dim a_SlabFound As Integer
    Dim aDif As Currency
    aUbound = UBound(g_OrderableItems)
    a_SlabFound = aUbound
    For aCounter = 1 To aUbound
        aDif = g_OrderableItems(aCounter).LengthAllowedInInches - in_LengthInInches
        If aDif > 0 Then
             If aDif < g_OrderableItems(a_SlabFound).LengthAllowedInInches - in_LengthInInches Then
                 a_SlabFound = aCounter
             End If
        ElseIf aDif = 0 Then
            a_SlabFound = aCounter
        End If
    Next
    GetSlabIndex = a_SlabFound
End Function



Public Sub SelectionSort(in_Items() As OrderItem)
Dim aIndex As Long, aSlot As Long, aMinSlot As Long, aUb As Long
Dim aMinItem As OrderItem
    'Selection Sort
    '----------------
    'The idea is to search the array for the smallest item
    'Then swap that item with the one at the top of the Array
    'Next find the smallest remaining item
    'And swap it with the second item in the Array
    'Continue until every item has been swapped into its final position
    
    aUb = UBound(in_Items)
    
    'Make each slot in turn available for the next smallest value
    For aIndex = 1 To aUb
     
        'Initialise the Min with the next item off the array
        Set aMinItem = in_Items(aIndex)
        aMinSlot = aIndex
        'Check all the items below the current slot for the smallest remaining value
        For aSlot = aIndex + 1 To aUb
            If in_Items(aSlot).Length < aMinItem.Length Then
                'Take a copy of the smallest value (this frees its slot)
                Set aMinItem = in_Items(aSlot)
                aMinSlot = aSlot
            End If
        Next
        'Swap the two values
        'First put the old value into the slot occupied by the smallest value
        Set in_Items(aMinSlot) = in_Items(aIndex)
        'Now copy the smallest value into the newly free Index
        Set in_Items(aIndex) = aMinItem
    Next

End Sub
