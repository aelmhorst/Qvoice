Attribute VB_Name = "POEditing"
Option Explicit

Public Enum POAlgorithmAction
    CreatePOLineItemsSingleBrand
    CombinePOLineItems
End Enum
Public Const STOCK_BRAND As String = "Z"


'// The ID used for items that are "on the next bulk order"
Private Const PENDING_BULK_PO_ID As Integer = -1
'// The ID used when creating a slab order from multiple orders
Private Const BULK_PO_BATCHID As Integer = -2

Private mOrderer As GeneticRawUnitOrderer

Public Function CanCreatePurchaseOrders(inOrder As Order) As Boolean
   CanCreatePurchaseOrders = MainModule.GetOrderType(inOrder.JobType).CanPost
End Function

Public Sub HandlePOCreation(in_Order As Order, inUser1 As String, inUser2 As String)
    Dim aController As New OrderPOCreationController
    aController.OrderPOCreationController in_Order.OrderID, inUser1, inUser2
End Sub

Public Sub EditPendingBulkPO()
    Dim aController As New POPurchaseOrderEditController
    aController.POPurchaseOrderEditController "Stock", -1
End Sub

' Orders caps for a specific laminate brand on an order
' used by the OrderPOCreationController
Public Sub OrderCapsForPO(in_OrderID As Long, in_BrandCode As String, in_POID As Long)
    Dim aRs As Recordset
    Set aRs = DataCenter.GetChargesForBrandCodeOnOrder(in_OrderID, in_BrandCode)
    Do While Not aRs.EOF
        DataCenter.InsertPODetailLine in_POID, aRs!vchLaminateCode, aRs!vchVendorCode, aRs!vchVendorDesc, 0, aRs!Quantity, 1, 0@, 0
        aRs.MoveNext
    Loop
    aRs.Close
    Set aRs = Nothing
End Sub

Public Sub CombineLineItemsToPO(in_POID As Long, inList As dbPOSlabList, in_Action As POAlgorithmAction, Optional in_Listener As iListener)
    Dim aCurrentBrandCode As String
    Dim aCurrentColorCode As String
    Dim aCurrentVendorCode As String
    Dim aCurrentVendorDesc As String
    Dim aCurrentVendorTypeID As Long
    Dim aCol As Collection
    Dim aCounter As Integer
    Dim aItem As OrderItem
   
    aCurrentBrandCode = inList.BrandCode
    If mOrderer Is Nothing Then Set mOrderer = New GeneticRawUnitOrderer
       
    Do While Not inList.EOF
        If ((in_Action = CreatePOLineItemsSingleBrand) And (aCurrentBrandCode <> inList.BrandCode)) Then Exit Do
        aCurrentColorCode = inList.LaminateCode
        aCurrentBrandCode = inList.BrandCode
        aCurrentVendorTypeID = inList.VendorSlabType
        aCurrentVendorCode = inList.VendorCode
        aCurrentVendorDesc = inList.VendorDescription
        
        ' Clear the Collection
        Set aCol = New Collection
        
        Do While Not inList.EOF
            If aCurrentColorCode <> inList.LaminateCode Then Exit Do
            If aCurrentVendorCode <> inList.VendorCode Then Exit Do
            If Len(aCurrentVendorCode) = 0 Then
                DataCenter.InsertPODetailLine in_POID, aCurrentColorCode, "#", "", 0, inList.Ordered, 1, inList.LengthInInches, inList.SerialID
            Else
            For aCounter = 1 To inList.Ordered
                Set aItem = New OrderItem
                aItem.Length = inList.LengthInInches
                aItem.SerialID = inList.SerialID
                aItem.ItemNumber = aCounter
                aCol.Add aItem
            Next
        End If
            inList.MoveNext
        Loop
        
        If aCol.Count > 0 Then
            If in_Action = CombinePOLineItems Then
                CombinePODetailLines aCol, aCurrentVendorTypeID
            Else
            If Not IsMissing(in_Listener) Then in_Listener.Receive "On Record " & inList.AbsolutePosition & " out of " & inList.RecordCount
            AddOrderItemsToPO aCol, aCurrentColorCode, aCurrentVendorCode, aCurrentVendorTypeID, aCurrentVendorDesc, in_POID
            End If
        End If
    Loop
End Sub

Private Sub CombinePODetailLines(in_OrderItems As Collection, in_VendorSlabTypeID As Long)
    
    Dim aUnitItems As RawUnitItems
    Set aUnitItems = DoGeneticOrder(in_OrderItems, in_VendorSlabTypeID)

    Dim aCurrentPOLineID As Long
    Dim aLineItem As RawUnitItem
    For Each aLineItem In aUnitItems
        '// Update the current line and then move to the new PO
        DataCenter.UpdatePurchaseOrderDetailOrderedSize aLineItem.SerialID, aLineItem.OrderableUnit.SlabLength
        aCurrentPOLineID = aLineItem.SerialID
        ' Move the detail lines from the children to the parent and then delete the children
        While Not aLineItem.Child Is Nothing
            Set aLineItem = aLineItem.Child
            DataCenter.CombinePurchaseOrderDetailLines aCurrentPOLineID, aLineItem.SerialID
        Wend
    Next

End Sub

Private Sub AddOrderItemsToPO(in_OrderItems As Collection, in_GroupByCode As String, in_VendorCode As String, in_VendorSlabTypeID As Long, in_Description As String, in_PO As Long)
    Dim aUnitItems As RawUnitItems
    Set aUnitItems = DoGeneticOrder(in_OrderItems, in_VendorSlabTypeID)
    DataCenter.InsertPODetailLines aUnitItems, in_GroupByCode, in_VendorCode, in_Description, in_PO
End Sub

Private Function DoGeneticOrder(in_OrderItems As Collection, in_VendorSlabTypeID As Long) As RawUnitItems
    Dim aItemsToOrder() As OrderItem
    Dim aCounter As Long
    Dim aUnitItems As RawUnitItems
    Static s_LastVendorSlabTypeID As Long
    Static s_OrderableUnits() As OrderableUnit

    ReDim aItemsToOrder(1 To in_OrderItems.Count)
    For aCounter = 1 To in_OrderItems.Count
        Set aItemsToOrder(aCounter) = in_OrderItems(aCounter)
    Next
        
    'Load our list of slab lengths
    If s_LastVendorSlabTypeID <> in_VendorSlabTypeID Then
        s_LastVendorSlabTypeID = in_VendorSlabTypeID
        s_OrderableUnits = DataCenter.GetVendorOrderableUnits(in_VendorSlabTypeID)
    End If
    
    Set DoGeneticOrder = mOrderer.OrderUnits(aItemsToOrder, s_OrderableUnits)
End Function


Private Function Ceiling(in_Double As Double) As Integer
    Dim aReturn As Integer
    aReturn = CInt(in_Double)
    If in_Double > aReturn Then aReturn = aReturn + 1
    Ceiling = aReturn
End Function


Public Sub CreateGeneralPO(inParent As Form)
    frmPOCreationDialog.Show vbModal, inParent
End Sub

Public Sub GeneratePOFromStockOrderItems( _
            in_OrderIds As Variant, _
            in_PORequiredDate As Date, _
            inUser1 As String, _
            inUser2 As String)
    Dim aPOList As dbPOSlabList
    Dim aPODoc As String
    Dim aPOID As Long
    
    Dim aCounter As Integer
    
    DataCenter.CreatePO in_PORequiredDate, 24, inUser1, inUser2, aPODoc, aPOID
    
    '// Move everything on the current bulk order to the new PO
    DataCenter.MovePODetailLines PENDING_BULK_PO_ID, aPOID
    
    '// Move the detail lines to the POID of -2. This is the working batch
    For aCounter = 0 To UBound(in_OrderIds, 2)
        DataCenter.MovePODetailLines 0 - in_OrderIds(0, aCounter), BULK_PO_BATCHID
    Next aCounter
    

    
    '// Open a recordset that will give all the information the genetic algorithm needs for ordering
    Set aPOList = DataCenter.GetPendingPurchaseOrderItemsBulk(BULK_PO_BATCHID)
    
    CombineLineItemsToPO aPOID, aPOList, POAlgorithmAction.CombinePOLineItems
    
   '// Combine and move the Charges to the new PO
    ProcessAndCombineBulkCharges aPOID
    
    '// Move the remaining items
    DataCenter.MovePODetailLines BULK_PO_BATCHID, aPOID
    
    POEditing.UpdatePOStatusOnAllOpenOrders
    
    '// Open the PO Editing screen
    Dim aController As New POPurchaseOrderEditController
    aController.POPurchaseOrderEditController aPODoc, aPOID
    
End Sub

Private Sub ProcessAndCombineBulkCharges(in_PurchaseOrderID As Long)
    Dim rs As Recordset
    Dim aCurrentColor As String
    Dim aCurrentChargeCode As String
    Dim aDescription As String
    Dim aQuantity As Long
    
    Set rs = DataCenter.GetBulkBatchChargesInOrder
    If rs.EOF Then Exit Sub
    ' TODO: Fix this in the query
    rs.Filter = "vchVendorItemCode <> Null"
    
    aCurrentColor = rs!vchGroupByCode
    aCurrentChargeCode = rs!vchVendorItemCode
    aDescription = rs!vchItemDescription
    aQuantity = 0
    
    Do Until rs.EOF
        If ((aCurrentColor <> rs!vchGroupByCode) Or (aCurrentChargeCode <> rs!vchVendorItemCode)) Then
            
            '// Save this as a PO Detail Item
            DataCenter.InsertPODetailLine in_PurchaseOrderID, aCurrentColor, aCurrentChargeCode, aDescription, 0, aQuantity, 1, 0@, 0
            aCurrentColor = rs!vchGroupByCode
            aCurrentChargeCode = rs!vchVendorItemCode
            aDescription = rs!vchItemDescription
            aQuantity = 0
        
        End If
        If IsNull(rs!iQuantity) Then
            '// This can happen in cases where the user has accidentally
            '// deleted the quantity. We will fudge to a quantity of one, rather
            '// than blowing up in their face
            aQuantity = aQuantity + 1
        Else
            aQuantity = aQuantity + rs!iQuantity
        End If
        DataCenter.RemovePODetailLine rs!iPurchaseOrderLineId
       rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    If aQuantity > 0 Then
        DataCenter.InsertPODetailLine in_PurchaseOrderID, aCurrentColor, aCurrentChargeCode, aDescription, 0, aQuantity, 1, 0@, 0
    End If
  
End Sub


Public Sub GenerateBlankPO(in_VendorID As Long, in_PORequiredDate As Date, _
            inUser1 As String, inUser2 As String)
    Dim aController As New POPurchaseOrderEditController
    Dim aPODoc As String
    Dim aPOID As Long
    
    DataCenter.CreatePO in_PORequiredDate, in_VendorID, inUser1, inUser2, aPODoc, aPOID
    DataCenter.InsertPODetailLine aPOID, "", "", "", 0, 1, 1, 0@, 0
    aController.POPurchaseOrderEditController aPODoc, aPOID

End Sub

Public Sub PrintPurchaseOrder(inPurchaseOrderID As Long, inShowPreviewWindow As Boolean, inUpdatePOStatus As Boolean)
    PrintPODoc inPurchaseOrderID, inShowPreviewWindow, Constants.REPORT_NAME_PO, inUpdatePOStatus
End Sub

Public Sub PrintPurchaseOrderDetail(inPurchaseOrderID As Long, inShowPreviewWindow As Boolean)
    PrintPODoc inPurchaseOrderID, inShowPreviewWindow, Constants.REPORT_NAME_PO_DETAIL, False
End Sub

Public Sub PrintPurchaseOrderSlabLabels(inPurchaseOrderID As Long, inShowPreviewWindow As Boolean)
    ReportPrinter.PrintGenericDocument _
        "purchaseorderslablabels", _
        "{PurchaseOrder.iPurchaseOrderID} = " & inPurchaseOrderID & _
        " AND {PurchaseOrderDetail.tiIsSlab} = TRUE", _
        False, _
        inShowPreviewWindow, _
        "View PurchaseOrder", _
        Globals.MainDocumentWindow.hwnd, _
        True
End Sub

Private Sub PrintPODoc(in_POID As Long, in_PreviewWindow As Boolean, inDocName As String, inUpdatePOStatus As Boolean)
    ReportPrinter.PrintGenericDocument _
        inDocName, _
        "{PurchaseOrder.iPurchaseOrderID} = " & in_POID, _
        False, _
        in_PreviewWindow, _
        "View PurchaseOrder", _
        Globals.MainDocumentWindow.hwnd, _
        True
    
    If inUpdatePOStatus Then
        POEditing.UpdatePOStatus in_POID, POStatusOnPO
    End If
End Sub


Public Sub SavePOOrderLineItems(in_RS As Recordset, inDeletedItems As Collection)
    Dim llngRecordID As Variant
    For Each llngRecordID In inDeletedItems
        DataCenter.UpdateOrderLineOnPOToFalse CLng(llngRecordID)
    Next
    DataCenter.ReconnectAndUpdate in_RS
End Sub

Public Function GetOrderableLength(inFootage As Currency, inInches As Currency) As Currency
    If inFootage > 8 Or inInches = 0 Then
        GetOrderableLength = inFootage * 12@
    Else
        GetOrderableLength = inInches
    End If
End Function

Public Sub UpdatePOStatus(inPurchaseOrderID As Long, inPOStatus As POStatusEnum)
    Dim rs As Recordset
    Dim aRecordsAffected As Integer
    aRecordsAffected = DataCenter.UpdatePOStatus(inPurchaseOrderID, inPOStatus)
End Sub

Public Sub UpdatePOStatusOnAllOpenOrders()
    Dim rs As Recordset
    Dim aRemaining As Integer
    Dim aStatus As POStatusEnum
    aStatus = POStatusNew
    
    Set rs = DataCenter.GetOrderPOStatusResults()
    While Not rs.EOF
        If rs.Fields.Item(Constants.PO_DETAIL_ORDERED).Value > 0 Then
            aStatus = POStatusOnPO
            aRemaining = rs.Fields.Item(Constants.PO_DETAIL_ORDERED).Value - rs.Fields.Item(Constants.PO_DETAIL_RECEIVED).Value
            If aRemaining = 0 Then
                aStatus = POStatusReceived
            ElseIf rs.Fields.Item(Constants.PO_DETAIL_RECEIVED) > 0 Then
                aStatus = POStatusPartial
            End If
        End If
        If aStatus <> rs.Fields.Item(Constants.ORDER_HEADER_POSTATUSFLAG).Value Then
            DataCenter.UpdatePOStatusOnOrder rs.Fields.Item(Constants.ORDER_HEADER_ORDER_ID).Value, aStatus
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
End Sub

Public Function ReceivePODetailItem(inPurchaseOrderLineID As Long) As Boolean
    ReceivePODetailItem = DataCenter.UpdatePurchaseOrderDetailToReceived(inPurchaseOrderLineID)
End Function
