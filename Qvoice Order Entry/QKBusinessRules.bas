Attribute VB_Name = "QKBusinessRules"
Option Explicit
Option Compare Text


'*****************************************************************************
'   modQKBusinessRules
'   Created By:     Andy Elmhorst
'   Purpose:        Establishes the QK-specific rules on the various aspects
'                   of pricing out countertops
'*****************************************************************************



Public Function QKGetSpecialCharge(OrderLines As Lines) As Charge
    Dim lcharge             As New Charge
    Dim lclsline            As Line
    Dim aDctLaminate             As Dictionary
    Dim aDctSlab            As Dictionary
    Dim lstrSlabKey         As String
    Dim lvarItem            As Variant
    Dim lintCounter         As Integer
    Dim lcurCharge          As Currency
    Dim lstrDesc             As String
    
    Set aDctLaminate = New Dictionary
    Set aDctSlab = New Dictionary
    
    If OrderLines Is Nothing Then
        Set QKGetSpecialCharge = New Charge
        Exit Function
    End If
    
    For Each lclsline In OrderLines
        With lclsline
            If .LinealFeet > 0 Then
                'Check for the Laminate Brand Upcharge (Wausau Homes)
                If .LamTopUpcharge > 0 Then
                    If aDctLaminate.Exists(.LaminateBrand) Then
                        'increment the item count
                        lvarItem = aDctLaminate(.LaminateBrand)
                        lvarItem(0) = lvarItem(0) + .Ordered
                        aDctLaminate(.LaminateBrand) = lvarItem
                    Else
                        'add this item into the mix
                        aDctLaminate.Add .LaminateBrand, Array(.Ordered, .LamTopUpcharge, .LamJobUpcharge)
                    End If
                End If
                'Check for the minimum footage upcharge
                lstrSlabKey = .SlabID & "-" & .LaminateID
                If aDctSlab.Exists(lstrSlabKey) Then
                    'increment the lineal feet
                    lvarItem = aDctSlab(lstrSlabKey)
                    lvarItem(0) = lvarItem(0) + (.LinealFeet * .Ordered)
                    aDctSlab(lstrSlabKey) = lvarItem
                Else
                    'add this item into the mix
                    aDctSlab.Add lstrSlabKey, Array((.LinealFeet * .Ordered), .SlabDesc, .LaminateDesc, .PerFootCharge, .SlabMinimum)
                End If
            End If
        End With
    Next
    
    'Now go through both collections and see what we need to charge extra for.
    If aDctLaminate.Count > 0 Then
        For lintCounter = 0 To (aDctLaminate.Count - 1)
            lvarItem = aDctLaminate.Items(lintCounter)
            lstrDesc = lstrDesc & aDctLaminate.Keys(lintCounter) & " upcharge. "
            If (lvarItem(0) * lvarItem(1)) > lvarItem(2) Then
                lcurCharge = lcurCharge + lvarItem(2)
            Else
                lcurCharge = lcurCharge + (lvarItem(0) * lvarItem(1))
            End If
        Next
    End If
    
    If aDctSlab.Count > 0 Then
        For lintCounter = 0 To (aDctSlab.Count - 1)
            lvarItem = aDctSlab.Items(lintCounter)
            If lvarItem(0) < lvarItem(4) Then
                lstrDesc = lstrDesc & lvarItem(4) & " foot minumum upcharge for " & _
                    lvarItem(1) & " in " & lvarItem(2) & ". "
                lcurCharge = lcurCharge + ((lvarItem(4) - lvarItem(0)) * lvarItem(3))
            End If
        Next
    End If
    
    Set lcharge = New Charge
    If lcurCharge > 0 Then lcharge.Init 0, 1, "MISC", lcurCharge, lstrDesc, False, 0@
    Set QKGetSpecialCharge = lcharge

End Function
Public Function DimensionstoFeet(Dmstring As String) As Long
    Dim lsngDim1            As Long
    Dim lsngDim2            As Long
    Dim lintSepPosition     As Integer
    
    lintSepPosition = InStr(1, Dmstring, "x", vbTextCompare)
    
    lsngDim1 = QKInchesToFeet(CSng(Left$(Dmstring, _
                        lintSepPosition - 1)))
    lsngDim2 = QKInchesToFeet(CSng(Mid$(Dmstring, _
                        lintSepPosition + 1)))
    DimensionstoFeet = lsngDim1 * lsngDim2
End Function

    
'*****************************************************************************
'   Public Function InchesToLineal( lsngInches As Single,  lintCustomerList As Integer, Optional lintSlabAvail As Integer) As Single
'   Created By:     Andy Elmhorst
'   Date:           09/27/1998
'   Purpose:        Takes a specified length in inches and converts it to lineal feet, following
'                   QK-specific guidelines
'*****************************************************************************
Public Function InchesToLineal(lcurInches As Currency, _
                pvarAvailableLengths As Variant) As Currency

    Dim lcurLineal As Currency
    Dim lintCounter As Integer
    Dim lintMax As Integer
    
    If lcurInches = 0 Then
        InchesToLineal = 0
        Exit Function
    End If
    
    lintMax = UBound(pvarAvailableLengths)
    
    For lintCounter = 0 To lintMax
        If ((pvarAvailableLengths(lintCounter) * 12) + 0.25) >= lcurInches Then
            lcurLineal = pvarAvailableLengths(lintCounter)
            Exit For
        End If
    Next
    If lcurLineal = 0# Then lcurLineal = QKInchesToFeet(lcurInches)
    
    InchesToLineal = lcurLineal

End Function


Private Function QKInchesToFeet(Inches As Currency) As Long
    Dim llngLineal As Long
    Dim lcurInches As Currency
    
    llngLineal = Fix(Inches / 12@)
    lcurInches = Inches - (llngLineal * 12&)
    If lcurInches > 0.25@ Then llngLineal = llngLineal + 1&
    
    QKInchesToFeet = llngLineal
End Function

Public Function GetOrderableUnitLengthInInches(inFootLength As Integer) As Currency
    If inFootLength < 8 Then
        GetOrderableUnitLengthInInches = (inFootLength * 12@) + 0.25@
    Else
        GetOrderableUnitLengthInInches = (inFootLength * 12@) + 0.5@
    End If
End Function
