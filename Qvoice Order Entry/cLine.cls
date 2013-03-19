VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Charges"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Enum tAreaType
    Lineal = 1
    SquareEdge = 2
    Laminate = 3
End Enum


Private mOrder              As Order
Public Ordered              As Long             'The quantity ordered
Public Posted               As Long             'The actual quantity posted to an invoice
Public ShippedNow           As Long             'The amount being entered in the posting screen
Public SlabCode             As String
Public SlabLength           As Currency
Public SlabWidth            As Currency
Public SlabDesc             As String
Public SlabPrice            As Currency
Public LaminateDesc         As String
Public LineNumber           As Long

Private mlngSlabID          As Long
Private mlngSerialID        As Long
Private mlngLaminateID      As Long
Private mlngShipped         As Long             'The total quantity shipped

Private mcurLaminateUpch    As Currency
Private mcurLamJobCharge    As Currency
Private mcurLamTopCharge    As Currency
Private mcurUnitPrice            As Currency
Private mcurLineal          As Currency
Private mcurMinimumLength   As Currency

Private mstrLamCode         As String
Private mstrLamBrand        As String
Private mvarAvailability    As Variant
Private mvarCharges         As Charges

'Flags
Private mtLinealType        As tAreaType
Private mbolChanged         As Boolean
Private mstrDescription As String



Public Property Get Shipped() As Long
    Shipped = mlngShipped
End Property
Public Property Let Shipped(in_Shipped As Long)
    Changed = True
    set_Shipped in_Shipped
End Property
Private Sub set_Shipped(in_Shipped As Long)
    mlngShipped = in_Shipped
End Sub

Friend Property Get LamJobUpcharge() As Currency
    LamJobUpcharge = mcurLamJobCharge
End Property
Friend Property Get LaminateCode() As String
    LaminateCode = mstrLamCode
End Property
Friend Property Get LamTopUpcharge() As Currency
    LamTopUpcharge = mcurLamTopCharge
End Property
Public Property Get LaminateBrand() As String
    LaminateBrand = mstrLamBrand
End Property
Public Property Get LaminateUpcharge() As Currency
    LaminateUpcharge = mcurLaminateUpch
End Property
Friend Property Get SlabMinimum() As Currency
    SlabMinimum = mcurMinimumLength
End Property
Public Property Get Charges() As Charges
    Dim rs As Recordset
    If mvarCharges Is Nothing Then
        Set mvarCharges = New Charges
        If mlngSerialID > 0 Then
            'Load the charges
            #If DBUG Then
                PerfLog "Loading Charges For " & mlngSerialID, 3
            #End If
            Set rs = DataCenter.GetChargesOnLine(mlngSerialID)
            With rs
                Do Until .EOF
                    If Not IsNull(!tiLineal) Then
                        mvarCharges.Add !iChargeID, !vchChargeCode, !vchChargeDesc, _
                              !flchargeprice, !iQuantity, (!tiLineal = 1), 0, !iLineChargeID
                    Else
                        mvarCharges.Add !iChargeID, !vchChargeCode, !vchChargeDesc, _
                              !flchargeprice, !iQuantity, False, 0, !iLineChargeID
                    End If
                    .MoveNext
                Loop
            End With
            rs.Close
            Set rs = Nothing
            #If DBUG Then
                PerfLog "Completed Loading Charges For " & mlngSerialID, 3
            #End If
        End If
    End If
    Set Charges = mvarCharges
End Property

Private Sub Class_Terminate()
    Set mvarCharges = Nothing
    If Not mOrder Is Nothing Then Set mOrder = Nothing
End Sub

Public Property Get Unposted() As Long
Unposted = Shipped - Posted
End Property
Public Property Get Unshipped() As Long
Unshipped = Ordered - Shipped
End Property

Public Property Get SerialID() As Long
    SerialID = mlngSerialID
End Property
Public Property Let SerialID(plngSerialID As Long)
    mlngSerialID = plngSerialID
End Property

Public Sub Load(pobjOrder As Order, iSerialID As Long, rs As Recordset)

Set mOrder = pobjOrder
mlngSerialID = iSerialID

'Set rs = DataCenter.GetLine(iSerialID)
With rs
    LineNumber = !iLineNumber
    Ordered = !iOrdered
    set_Shipped !iShipped
    Posted = !iPosted
    set_SlabID !iSlabid
    SlabLength = !flSlabLength
    SlabWidth = !flSlabWidth
    set_LaminateID !iLaminateID
    mcurLineal = !flSlabArea
    mcurUnitPrice = CCur(!flLineExtPrice) / CCur(Ordered) / CCur(mOrder.Customer.Multiplier)
    mstrDescription = !txtLineDesc
End With
'rs.Close
'Set rs = Nothing

Changed = False
End Sub
Public Sub Customer_Changed()
    'Refresh the properties on our Slab and Laminate ID's
    SlabID = mlngSlabID
    LaminateID = mlngLaminateID
    Changed = True
End Sub

Public Property Let SlabID(in_SlabID As Long)
    If in_SlabID <> mlngSlabID Then
        Changed = True
    End If
    set_SlabID in_SlabID
End Property

Private Sub set_SlabID(in_SlabID As Long)
Dim rs As Recordset
mlngSlabID = in_SlabID
If in_SlabID > 0 Then
    Set rs = DataCenter.GetSlab(in_SlabID, mOrder.PriceListID)
    With rs
        SlabCode = rs!vchSlabCode
        SlabDesc = rs!vchSlabDesc & ""
        mvarAvailability = Split(!vchAvailableLengths, ",")
        mtLinealType = !tiCalcType
        SlabPrice = !flPrice
        mcurMinimumLength = !flMinimumLength
        .Close
    End With
    Set rs = Nothing
End If
End Sub

Public Property Get SlabID() As Long
SlabID = mlngSlabID
End Property
Private Property Get Changed() As Boolean
    Changed = mbolChanged
End Property
Private Property Let Changed(pbolChanged As Boolean)
    mbolChanged = pbolChanged
    If pbolChanged Then
        mcurUnitPrice = 0@
        mstrDescription = ""
    End If
End Property

Public Property Let LaminateID(in_LaminateID As Long)
If mlngLaminateID <> in_LaminateID Then
    Changed = True
End If
set_LaminateID in_LaminateID
End Property
Private Sub set_LaminateID(in_LaminateID As Long)
    Dim rs As Recordset
    If in_LaminateID > 0 Then
        Set rs = DataCenter.GetLaminate(in_LaminateID, mOrder.PriceListID, mlngSlabID)
            mlngLaminateID = in_LaminateID
            mstrLamCode = rs!vchLaminateCode
            LaminateDesc = rs!vchLaminateDesc
            mcurLaminateUpch = rs!mSlabUpCharge
            mcurLamJobCharge = rs!flJobUpcharge
            mcurLamTopCharge = rs!flTopUpcharge
            mstrLamBrand = rs!vchBrandDescription
        rs.Close
        Set rs = Nothing
    End If
End Sub

Public Sub Save(iOrderId As Long)
Dim lobjCharge As Charge

If Changed Or mlngSerialID = 0 Then
    'Save this line
    DataCenter.SaveLine iOrderId, Me
End If
End Sub
Public Property Get LaminateID() As Long
    LaminateID = mlngLaminateID
End Property
Public Property Get MeasureLineal() As Boolean
    MeasureLineal = (mtLinealType = Lineal)
End Property
Public Property Get PerFootCharge() As Single
    PerFootCharge = SlabPrice + LaminateUpcharge
End Property

Public Property Get LinealFeet() As Currency
    Dim lcharge As Charge
    
    If mcurLineal = 0 Then
        If mtLinealType = Lineal Then
           mcurLineal = InchesToLineal(Me.OrderableLengthInInches, mvarAvailability)
        Else
             If SlabLength > 0 Then
                If mtLinealType = SquareEdge Then
                    mcurLineal = InchesToArea(SlabLength) * InchesToArea(SlabWidth)
                Else
                    mcurLineal = SlabLength / 144 * SlabWidth
                End If
            Else
                mcurLineal = 0
            End If
        End If
    End If
    LinealFeet = mcurLineal
End Property

Public Property Get OrderableLengthInInches() As Currency
    
    Dim aActualLength As Currency
    aActualLength = SlabLength
    
    If (mtLinealType <> Lineal) Or (aActualLength = 0) Then
        OrderableLengthInInches = 0
        Exit Property
    End If
    
    Dim lcharge As Charge
    For Each lcharge In Charges
        If Abs(lcharge.LengthAdjustment) > 0 Then
            aActualLength = aActualLength + (lcharge.Quantity * lcharge.LengthAdjustment)
        End If
    Next
    OrderableLengthInInches = aActualLength
End Property

Private Function InchesToArea(pcurInches As Currency) As Currency
Dim pcurArea As Currency
pcurArea = Fix(pcurInches / 12@)
If pcurInches > (pcurArea * 12) Then pcurArea = pcurArea + 1
InchesToArea = pcurArea
End Function


Public Property Get LineDescription() As String
Dim lcharge As Charge
Dim lintCounter As Integer

If Len(mstrDescription) > 0 Then
    LineDescription = mstrDescription
Else
    If SlabLength > 0 Then
        If MeasureLineal Then
            mstrDescription = Format$(SlabLength, "###.000") & " inch " & Trim$(SlabDesc)
        Else
            mstrDescription = Format$(SlabLength, "###.000") & " X " & Format$(SlabWidth, "###.000") & " inch " & Trim$(SlabDesc)
        End If
        lintCounter = 1
    End If
    
    For Each lcharge In Me.Charges
        lintCounter = lintCounter + 1
        If lcharge.Quantity > 1 Then
            mstrDescription = mstrDescription & _
                IIf(lintCounter > 1, ", ", "") & ToWords(lcharge.Quantity) & Space$(1) & _
                Trim$(lcharge.Description) & Chr$(115)
        Else
            mstrDescription = mstrDescription & IIf(lintCounter > 1, ", ", "") & Trim$(lcharge.Description)
        End If
    Next
    
    If mlngLaminateID > 0 Then
        mstrDescription = mstrDescription & ", Color:(" & _
             Trim$(mstrLamCode) & ") " & Trim$(LaminateDesc)
            
    End If
    
    If LinealFeet > 0 Then
        If MeasureLineal Then
            mstrDescription = mstrDescription & " [" & LinealFeet & " LF]"
        Else
            mstrDescription = mstrDescription & " [" & LinealFeet & " SF]"
        End If
    End If
    LineDescription = mstrDescription
End If
End Property
Public Property Get ChargeDescriptions() As String
Dim lstrworkstring As String
Dim lcharge As Charge
Dim lintIndex As Integer
Dim lstrComma As String

For Each lcharge In Me.Charges
    lintIndex = lintIndex + 1
    If lintIndex > 1 Then lstrComma = ","
    If lcharge.Quantity > 1 Then
        lstrworkstring = lstrworkstring & _
          lstrComma & ToWords(lcharge.Quantity) & Space$(1) & _
            Trim$(lcharge.Description) & Chr$(115)
    Else
        lstrworkstring = lstrworkstring & lstrComma & Trim$(lcharge.Description)
    End If
Next

ChargeDescriptions = lstrworkstring
End Property
Public Property Get List() As Currency
    List = CCur(UnitPrice * Ordered) / mOrder.Customer.Multiplier
End Property

Public Property Get UnitPrice() As Currency
If mcurUnitPrice = 0@ Then
    CalculatePrice
End If
    UnitPrice = CCur(CLng(mcurUnitPrice * mOrder.Customer.Multiplier * 100) / 100)
End Property
Public Property Get Total() As Currency
    Total = UnitPrice * Ordered
End Property


Public Sub Add(pOrder As Order, plngLineNumber As Long, plngQuantity As Long, pcurLength As Currency, _
            pcurewidth As Currency, plngSlabID As Long, plngLaminateID As Long)

Set mOrder = pOrder

LineNumber = plngLineNumber
Ordered = plngQuantity
SlabLength = pcurLength
SlabWidth = pcurewidth
SlabID = plngSlabID
LaminateID = plngLaminateID
Changed = True

' Reset the lineal feet so the it will get recalculated
mcurLineal = 0


End Sub

Private Sub CalculatePrice()
Dim lcharge     As Charge

mcurUnitPrice = 0@

For Each lcharge In Charges
    If lcharge.Lineal Then
        mcurUnitPrice = mcurUnitPrice + (lcharge.TotalCharge * LinealFeet)
    Else
        mcurUnitPrice = mcurUnitPrice + (lcharge.TotalCharge)
    End If
Next
    mcurUnitPrice = mcurUnitPrice + (LinealFeet * PerFootCharge)
End Sub
