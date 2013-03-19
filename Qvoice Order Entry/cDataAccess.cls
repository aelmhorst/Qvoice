VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'   Module::DataCenter
'   Created By:     Andy Elmhorst
'   Date:           02/18/1999
'   Purpose:        All of the saving and loading of data is done in this
'                   central location for the purpose of encapsulation
'*************************************************************************
Option Explicit
Option Compare Text

'Public Numerics
Private mlngSelectedJobID      As Long
Public SelectedJobInfo          As String



'Objects
Private mRsOEReps               As Recordset
Private mrsSlabs                As Recordset
Private mChargecommand         As Command

Private mcolSettings        As Dictionary
#If DBUG Then
    Private WithEvents mCn      As ADODB.Connection
Attribute mCn.VB_VarHelpID = -1
#Else
    Private mCn                 As ADODB.Connection
#End If
Public Property Let SelectedJobID(plngID As Long)
    mlngSelectedJobID = plngID
    SelectedJobInfo = ""
End Property
Public Property Get SelectedJobID() As Long
    SelectedJobID = mlngSelectedJobID
End Property

Private Property Get Cn() As ADODB.Connection
    If mCn Is Nothing Then Set mCn = New ADODB.Connection
    If mCn.State = adStateClosed Then
        lConnect
    End If
    Set Cn = mCn
End Property
Public Sub SaveOrder(ByRef Order As Order)
    Dim rs As Recordset
    Dim mvarfields As Variant
    Dim mvardata As Variant
    
    mvarfields = Array("iCustomerID", "iPriceListID", "vchOrderNumber", "vchRefNumber", "dtEntryDate", "dtUpdateDate", "dtPostedDate", "iStatus", "iOrderType", _
            "dtRequestDate", "vchJobName", "vchPONumber", "vchShipAddress1", _
            "vchShipAddress2", "vchShipAddress3", "vchShipAddress4", "vchShipAddress5", "tiCartoned", "tiPadded", "tiSplined", _
            "vchTrackingCode", "iUserID", "mOrderTotal", "txtComment", "tiGroup", "iBatchID", Constants.ORDER_HEADER_ORDER_RUSHFLAG)
    With Order
        mvardata = Array(.Customer.UniqueID, .PriceListID, .OrderNumber, .GetProperty("vchRefNumber"), .Entrydate, Now(), .GetProperty("dtPostedDate"), .State, .JobType, .RequiredDate, .GetProperty("vchJobName"), .PO, .ShipTo(1), _
            .ShipTo(2), .ShipTo(3), .ShipTo(4), .ShipTo(5), .Cartoned, .Padded, .Splined, .Trackingcode, .UserID, .Total, .GetProperty("txtComment"), .GetProperty("tiGroup"), .GetProperty("iBatchID"), .Rush)
    End With
    Set rs = New Recordset
    If Order.OrderID = 0 Then
            rs.CursorLocation = adUseServer
            rs.Open "Select * from vorderheader where iOrderID = 0", Cn, adOpenDynamic, adLockPessimistic
            rs.AddNew mvarfields, mvardata
            rs.MoveLast
            Order.OrderID = rs(Constants.ORDER_HEADER_ORDER_ID)
    Else
            rs.CursorLocation = adUseServer
            rs.Open "Select * from vOrderHeader where iOrderID = " & Order.OrderID, Cn, adOpenDynamic, adLockPessimistic
            rs.Update mvarfields, mvardata
    End If
            
    rs.Close
    Set rs = Nothing
    If Order.Lines.DeletedLines.Count > 0 Then
        For Each mvardata In Order.Lines.DeletedLines
            Cn.Execute "qkdOrderLine " & mvardata
        Next
    End If
    
End Sub

Public Function GetCustomerList() As Recordset
    Set GetCustomerList = GetData("qkoegCustomerList")
End Function

Public Function GetAddress(in_Address As Integer) As AddressInfo
    Dim rs                  As Recordset
    
    Set rs = GetDataFromSingularCommand("qkoegAddress", "iAddressid", adInteger, 2, in_Address)
    
    Dim aAddress As New AddressInfo
    
    With rs
        aAddress.PreAddressInfo = !PreAddressInfo & ""
        aAddress.Name = !vchLocationName & ""
        aAddress.Address1 = !vchAddress1 & ""
        aAddress.Address2 = !vchAddress2 & ""
        aAddress.City = !vchCity & ""
        aAddress.State = !vchState & ""
        aAddress.Zip = !vchZip & ""
        aAddress.Phone = !vchTelephoneNumber & ""
        .Close
    End With
    
    Set GetAddress = aAddress
    
    Set rs = Nothing

End Function

Public Function GetCustomerShipLocations(in_CustomerID As Long) As Recordset

    Set GetCustomerShipLocations = GetDataFromSingularCommand("qkoegCustomerShipLocations", "iCustomerID", adInteger, 2, in_CustomerID)

End Function

Public Function GetSlabList() As Recordset

    If mrsSlabs Is Nothing Then
        Set mrsSlabs = GetData("select iSlabID,vchSlabCode,vchslabDesc from qkgslabinfo")
    Else
        mrsSlabs.MoveFirst
    End If
    
    Set GetSlabList = mrsSlabs
End Function


Public Function LoadOrderTypes() As Recordset
    Set LoadOrderTypes = GetData("Select * from qkgordertypes")
End Function

Public Function GetCustomer(iCustomerID As Long) As Recordset
    
    Set GetCustomer = GetDataFromSingularCommand("qkcgCustomer", "iCustomerID", adInteger, 2, iCustomerID)

End Function

Public Function GetColorMatches(lstrSearchStr As String, Optional pstrProcedureName As String = "qklgColorMatches") As Recordset
    Dim rs                  As Recordset
    
    On Error GoTo errhandler
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    Set rs = GetDataFromSingularCommand(pstrProcedureName, "vchSearchString", adVarChar, 30, lstrSearchStr)
    
    Set GetColorMatches = rs
    
    
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Function
errhandler:
        Screen.MousePointer = MousePointerConstants.vbDefault
        HandleError Err.Description, False
        
End Function

Public Sub Connect()
    Dim rs As Recordset
    
    If mcolSettings Is Nothing Then
        Set rs = GetData("Select * from gkgAppSettings")
        With rs
            Set mcolSettings = New Dictionary
            mcolSettings.CompareMode = TextCompare
            Do Until .EOF
                mcolSettings.Add !vchSettingKey.Value, !vchSetting.Value
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
    End If
    
    'Retrieve our path settings specific to our data files
'    If Len(mstrExtension) = 0 Then
'        mstrExtension = GetIniSetting("General", "JobDataExtension", "qll")
'        mstrDataPath = App.Path
'        mstrInvoicePath = GetSetting("QK", "General", "InvoiceFolder", "c:\")
'        'mstrOrderPath = GetIniSetting("General", "OrderPath", "")
'    End If

End Sub

Public Function Setting(pstrSettingKey As String) As String
    If mcolSettings.Exists(pstrSettingKey) Then
        Setting = mcolSettings(pstrSettingKey)
    Else
        Setting = ""
    End If
End Function
Private Sub lConnect()
    With mCn
        .ConnectionString = Settings.ConnectionString
        .Open
    End With
End Sub

Public Sub ShowOpen(pintOrderType As Integer, in_OpenReason As OpenOrderReason)
    frmLoadJobs.Init pintOrderType, in_OpenReason
    If SelectedJobID = 0 Then frmLoadJobs.Show vbModal
End Sub

Public Function GetOrders(inOrderType As Integer, inMinRequestDate As Date) As Recordset
    Dim aCmd   As Command
    Set aCmd = New Command
    With aCmd
        Set .ActiveConnection = Cn
        .CommandText = "qkogOrdersByType"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("iOrderType", adInteger, adParamInput, Value:=inOrderType)
        .Parameters.Append .CreateParameter("dtMinRequestDate", adDate, adParamInput, Value:=inMinRequestDate)
    End With
    Set GetOrders = GetDataFromCmd(aCmd)
End Function

Public Function GetOrdersWithoutPos(plngOrderType As Integer) As Recordset

    Set GetOrdersWithoutPos = GetData("qkpogOrdersWithoutPos")
End Function


Public Function GetOrderHeader(OrderID As Long) As Recordset

    Set GetOrderHeader = GetDataFromSingularCommand("qkogOrderHeader", Constants.ORDER_HEADER_ORDER_ID, adInteger, 2, OrderID)

End Function

Public Function GetOeLinesForPOLine(iPODetailLineID As Long) As Recordset
    
    Set GetOeLinesForPOLine = GetDataFromSingularCommand("qkPOGOELinesForPOLine", "iPurchaseOrderLineID", adInteger, 2, iPODetailLineID)

End Function


Public Sub SaveLine(iOrderId As Long, pLine As qkorder.Line)
    Dim rs              As Recordset
    Dim lbolNewLine     As Boolean
    Dim mvarfields      As Variant
    Dim mvardata        As Variant
    
    mvarfields = Array(Constants.ORDER_HEADER_ORDER_ID, "iSlabID", "txtLineDesc", "iLineNumber", "flSlabLength", "flSlabWidth", _
        "flSlabArea", "flSlabPrice", "flOrderableLength", "iLaminateID", "iOrdered", "iShipped", _
        "iPosted", "flLaminateExtPrice", _
        "flChargesExtPrice", "flLineExtPrice")
    With pLine
        mvardata = Array(iOrderId, .SlabID, .LineDescription, .LineNumber, .SlabLength, .SlabWidth, .LinealFeet, _
        .SlabPrice, .OrderableLengthInInches, .LaminateID, .Ordered, .Shipped, .Posted, _
       .LaminateUpcharge, .Charges.ExtPrice, .Total)
    End With
    
    Set rs = New Recordset
    If pLine.SerialID > 0 Then
        rs.Open "Select * from vorderline where iSerialID = " & pLine.SerialID, Cn, adOpenKeyset, adLockPessimistic
        rs.Update mvarfields, mvardata
    Else
        lbolNewLine = True
        rs.Open "select * from vorderline where iOrderID = " & iOrderId, Cn, adOpenKeyset, adLockPessimistic
        rs.AddNew mvarfields, mvardata
        pLine.SerialID = rs!iSerialID
    End If
    
    rs.Close
    Set rs = Nothing
    SaveCharges pLine, lbolNewLine

End Sub

Public Sub CreatePO(in_RequestedDate As Date, _
    in_VendorID As Long, _
    inUser1 As String, _
    inUser2 As String, _
    ByRef out_DocPONumber As String, _
    ByRef out_DocID As Long)
    Dim rs As New Recordset

    out_DocPONumber = NextIncrement("PO")
    rs.Open "PurchaseOrder", Cn, adOpenKeyset, adLockPessimistic, adCmdTable
    rs.AddNew Array( _
                Constants.PURCHASE_ORDER_NUMBER, _
                "iVendorID", _
                "dtRequested", _
                "dtDateOrdered", _
                Constants.PURCHASE_ORDER_STATUS, _
                Constants.PURCHASE_ORDER_USER1, _
                Constants.PURCHASE_ORDER_USER2), _
            Array( _
                out_DocPONumber, _
                in_VendorID, _
                in_RequestedDate, _
                Now, _
                0, _
                Left$(inUser1, 30), _
                Left$(inUser2, 30))
    out_DocID = rs!iPurchaseOrderID
    rs.Close

End Sub

Public Function UpdatePOStatus(inPurchaseOrderID As Long, inPOStatus As POStatusEnum) As Long
    Dim cmd As New Command
    Dim aDateField As String
    Dim aRecordsAffected As Long
    
    Select Case inPOStatus
        Case POStatusEnum.POStatusNew, POStatusEnum.POStatusOnPO
            aDateField = Constants.PURCHASE_ORDER_DATE_ORDERED
        Case POStatusEnum.POStatusPartial, POStatusEnum.POStatusReceived
            aDateField = Constants.PURCHASE_ORDER_DATE_RECEIVED
    End Select
            
    With cmd
        .CommandType = adCmdText
        .CommandText = "Update PurchaseOrder Set iPOStatus = " & inPOStatus & ", " & aDateField & " = #" & Now() & "# Where iPurchaseOrderID = " & inPurchaseOrderID & " AND iPOStatus <> " & inPOStatus
        Set .ActiveConnection = Cn
        .Execute aRecordsAffected, Options:=adExecuteNoRecords
    End With
    
   Set cmd = Nothing

   UpdatePOStatus = aRecordsAffected
End Function

Private Sub SaveCharges(pLine As qkorder.Line, pbolFirstSave As Boolean)
    Dim rs              As Recordset
    Dim mvarfields      As Variant
    Dim mvardata        As Variant
    Dim lcharge         As Charge
    Dim lstrUpdateString    As String
    
    mvarfields = Array("iSerialID", "iChargeID", "iQuantity", "vchChargeCode", "vchChargeDesc", "flChargePrice")
    Set rs = New Recordset
    
    rs.Open "Select * from vorderlinecharge where iSerialID = " & pLine.SerialID, Cn, adOpenKeyset, adLockPessimistic
    
    For Each lcharge In pLine.Charges
        With lcharge
            mvardata = Array(pLine.SerialID, .iChargeID, .Quantity, .Abbrev, Left$(.Description, 50), .Price)
            If .iLineChargeID > 0 Then
                rs.Filter = "iLineChargeID=" & .iLineChargeID
                If Not rs.EOF Then
                    rs.Update mvarfields, mvardata
                Else
                    rs.AddNew mvarfields, mvardata
                    .iLineChargeID = rs!iLineChargeID
                End If
            Else
                rs.AddNew mvarfields, mvardata
                .iLineChargeID = rs!iLineChargeID
            End If
            If Not pbolFirstSave Then lstrUpdateString = lstrUpdateString & .iLineChargeID & ","
        End With
    Next
    rs.Close
    Set rs = Nothing
    
    'Now delete any orphans, if any
    If Not pbolFirstSave And Len(lstrUpdateString) > 0 Then
        lstrUpdateString = "Delete From OrderLineCharge Where iSerialID = " & pLine.SerialID & _
                    " AND NOT iLineChargeID IN (" & Left$(lstrUpdateString, Len(lstrUpdateString) - 1) & ")"
        Cn.Execute lstrUpdateString, , adExecuteNoRecords
    End If
End Sub

Public Function GetSlab(iSlabid As Long, iPriceListID As Long) As Recordset
    
    Dim cmd         As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "gkigSlab"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("iSlabID", adInteger, adParamInput, 4, iSlabid)
        .Parameters.Append .CreateParameter("iPriceListID", adInteger, adParamInput, 4, iPriceListID)
    End With
    Set GetSlab = GetDataFromCmd(cmd)

End Function

Public Function GetPOItemsOnOrder(in_OrderID As Long) As dbPOSlabList
    Dim aReturnVal As New dbPOSlabList
    
    aReturnVal.POSlabList GetDataFromSingularCommand("qkpogPOItemsOnOrder", Constants.ORDER_HEADER_ORDER_ID, adInteger, 4, in_OrderID), SingleOrder
    Set GetPOItemsOnOrder = aReturnVal
End Function

Public Function GetPOsByStatus(inStatus As POStatusEnum) As Recordset
    Set GetPOsByStatus = GetDataFromSingularCommand("qkpogPurchaseOrderByStatus", "iPOStatus", adInteger, 4, inStatus)
End Function

Public Function getSpecialVendorItemsOnOrder(in_OrderID As Long) As Recordset
    Set getSpecialVendorItemsOnOrder = GetDataFromSingularCommand("qkpogSpecialVendorItemsOnOrder", Constants.ORDER_HEADER_ORDER_ID, adInteger, 4, in_OrderID)
End Function

Public Function GetPendingPurchaseOrderItemsBulk(in_POID As Long) As dbPOSlabList
    Dim aReturnVal As New dbPOSlabList
    
    aReturnVal.POSlabList GetDataFromSingularCommand("qkpovPendingPurchaseOrderItems", "iPurchaseOrderID", adInteger, 4, in_POID), POList
    Set GetPendingPurchaseOrderItemsBulk = aReturnVal
End Function

Public Function getPurchaseOrderList(in_OrderID As Long) As Recordset
    Set getPurchaseOrderList = GetDataFromSingularCommand("qkpogPurchaseOrdersOnOrder", Constants.ORDER_HEADER_ORDER_ID, adInteger, 4, in_OrderID)
End Function

Public Function GetCharge(pstrChargeCode As String, iPriceListID As Long) As Charge
    Dim rs          As Recordset
    Dim lcharge     As Charge
    
    Set lcharge = New Charge
    
    If mChargecommand Is Nothing Then
        Set mChargecommand = New Command
        With mChargecommand
            Set .ActiveConnection = Cn
            .CommandText = "qkigCharge"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("vchChargeCode", adVarChar, adParamInput, 20, pstrChargeCode)
            .Parameters.Append .CreateParameter("iPriceListID", adInteger, adParamInput, 4, iPriceListID)
            .Prepared = True
        End With
    Else
        mChargecommand!vchChargeCode = Left$(pstrChargeCode, 20)
        mChargecommand!iPriceListID = iPriceListID
    End If
    
    Set rs = New Recordset
    With rs
        .CursorLocation = adUseClient
        .Open mChargecommand, , adOpenForwardOnly, adLockReadOnly
    
        If Not .EOF Then lcharge.Init !iChargeID, 0, !vchChargeCode, _
            !flPrice, !vchChargeDesc, (!tiLineal = 1), !flLengthAdj
    End With
    rs.Close
    Set rs = Nothing
    Set GetCharge = lcharge

End Function
Public Function GetLaminate(iLaminateID As Long, iPriceListID As Long, iSlabid As Long) As Recordset

    Dim cmd         As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "qkigLaminate"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("iLaminateID", adInteger, adParamInput, 4, iLaminateID)
        .Parameters.Append .CreateParameter("iPriceListID", adInteger, adParamInput, 4, iPriceListID)
        .Parameters.Append .CreateParameter("iSlabID", adInteger, adParamInput, 4, iSlabid)
    End With
    
    Set GetLaminate = GetDataFromCmd(cmd)

End Function

Public Function GetLinesOnOrder(iOrderId As Long) As Recordset

    Set GetLinesOnOrder = GetDataFromSingularCommand("qkogLinesOnOrder", Constants.ORDER_HEADER_ORDER_ID, adInteger, 2, iOrderId)

End Function

Public Function GetChargesOnLine(iSerialID As Long) As Recordset
    Set GetChargesOnLine = GetData("qkcgChargeLine " & iSerialID)
End Function


Public Function GetLineCharges(in_SerialId As Long) As Recordset
    
    Set GetLineCharges = GetDataFromSingularCommand("qkogLineCharges", "iSerialID", adInteger, 2, in_SerialId)
    
End Function
Public Function NextIncrement(ByVal pvarOrderType As Variant) As String
    Dim rs              As Recordset
    Dim mlngInterval    As Long
    Dim lstr            As String
    
    Set rs = New Recordset
    With rs
        If IsNumeric(pvarOrderType) Then
            .Open "Select vchPrefix,iIncrement from DocumentType where iOrderType = " & pvarOrderType, Cn, adOpenForwardOnly, adLockPessimistic
        Else
            'This is used for Invoice Number and PO Number generation
            .Open "Select vchPrefix,iIncrement from DocumentType where vchPrefix = '" & pvarOrderType & "'", Cn, adOpenForwardOnly, adLockPessimistic
        End If
        If .EOF Then
            Err.Raise 1, App.ProductName, "Unable to locate document with type: " & pvarOrderType
            Exit Function
        End If
        mlngInterval = rs!iIncrement + 1&
        .Update "iIncrement", mlngInterval
        lstr = rs!vchPrefix & Format$(mlngInterval, "000000")
        .Close
    End With
    Set rs = Nothing
    NextIncrement = lstr
End Function


'*****************************************************************************
'   DataCenter::GetOEReps
'   Created By:     Andy Elmhorst
'   Date:           02/18/1999
'   Purpose:        Retrieves a recordset of all of the oereps
'                   The program expects to receive
'                   "vchUsername","vchUserID", and "iUserID"
'   Once this recordset is retrieved, it is cached in mrsOEReps
'*****************************************************************************
Public Function GetOEReps() As Recordset
    
    If mRsOEReps Is Nothing Then
        Set mRsOEReps = GetData("qkogOrderReps")
    Else
        mRsOEReps.MoveFirst
    End If
    
    Set GetOEReps = mRsOEReps
End Function

Private Sub Class_Initialize()
    Connect
End Sub

Private Sub Class_Terminate()
    If Not mRsOEReps Is Nothing Then
        mRsOEReps.Close
        Set mRsOEReps = Nothing
    End If
    If Not mrsSlabs Is Nothing Then
        mrsSlabs.Close
        Set mrsSlabs = Nothing
    End If
    If Not mChargecommand Is Nothing Then Set mChargecommand = Nothing
    
    If Not mCn Is Nothing Then
        Set mCn = Nothing
    End If
    If Not mcolSettings Is Nothing Then Set mcolSettings = Nothing
End Sub

Public Sub PostShipment(ByVal plngShipmentID As Long, ByVal pdtDeliveryDate As Date, _
                     ByRef plngOrdersAffected As Long, pListener As iListener)
    Dim rs          As Recordset
    Dim lrsInv      As Recordset
    Dim cmd         As Command
    Dim lstrDoc     As String
    Dim llngInvID   As Long
    Dim lOrder      As Order
    Dim lcharge     As Charge
    Dim llngCustomerID As Long
    Dim ldblCustomerTax As Currency
    Dim lstrCustomerName As String
    Dim ldblTax     As Currency
    Dim lcurInvtotal As Currency
    Dim lCustomer   As Customer
    
    On Error GoTo errhandler:
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    Set cmd = New Command
    With cmd
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = Cn
        .CommandText = "qkoegShipmentBatch"
        .Parameters.Append .CreateParameter("iBatchID", adInteger, adParamInput, 4, plngShipmentID)
    End With
    
    Set rs = New Recordset
    With rs
        .CursorLocation = adUseClient
        .Open cmd, , adOpenStatic, adLockReadOnly
    End With
    
    If rs.EOF Then
        pListener.Receive "No records found . . ."
        plngOrdersAffected = 0
    Else
        Do Until rs.EOF
            llngCustomerID = rs!iCustomerID
            lstrCustomerName = rs!vchCustomerName
            ldblCustomerTax = rs!flTaxRate
            
            
            '// factor out the common code here and get rid of the rest
            
            
            '// Begin a transaction for this invoice
            Cn.BeginTrans
            
            '//*******************************************
            '// This will need to change
            '// Only need to do multiple per customer if pref is set
            '//*******************************************
            Do Until llngCustomerID <> rs!iCustomerID
                If rs!iStatus = 0 Then
                    '// this is a new order
                    '// we're going to have to instantiate the order object and
                    '// process special charges, if any
                    
                    '// Factor this out into a common function that just does the order
                    '// Do the invoice header generation stuff first and then
                    '// Go through one or more orders, based on the pref
                    '// Do the order detail and update stuff, either with batch or order number
                    pListener.Receive "Processing Order " & rs!vchOrderNumber
                    Set lOrder = New Order
                    lOrder.Init rs!iOrderId
                    With lOrder
                        Set lcharge = .SpecialCharge
                        If lcharge.TotalCharge > 0@ Then
                            .Lines.Add 1, 0, 0, 0, 0
                            With .Lines(.Lines.Count)
                                .Charges.Add lcharge.iChargeID, _
                                        lcharge.Abbrev, lcharge.Description, _
                                        lcharge.Price, lcharge.Quantity, _
                                        lcharge.Lineal, lcharge.LengthAdjustment
                                .ShippedNow = 1
                            End With
                            Set lcharge = Nothing
                            .Save
                        End If
                    End With
                    Set lOrder = Nothing
                End If
                rs.MoveNext
                If rs.EOF Then Exit Do
            Loop
            
            '//Ok, we've handled the special charges for this customer
            '//Now, create the Invoice Header and detail records
            
            lstrDoc = NextIncrement("INV")
            pListener.Receive "Creating " & lstrDoc & " for " & lstrCustomerName
            Set lCustomer = New Customer
            lCustomer.Init llngCustomerID
            
            If lrsInv Is Nothing Then Set lrsInv = New Recordset
            lrsInv.Open "InvoiceHeader", Cn, adOpenKeyset, adLockPessimistic
            lrsInv.AddNew Array("vchInvoiceNumber", "iDocumentStatus", "iCustomerID", _
                "vchBillToLine1", "vchBillToLine2", "vchBillToLine3", "vchBillToLine4", _
                "dtShipDate", "mDeliveryCharge", "mInvoiceSubtotal", _
                "mTax", "mInvoiceTotal", "tiPrinted", "vchSignature"), _
                   Array(lstrDoc, 0, llngCustomerID, _
                lCustomer.AddressInfo.Line2, lCustomer.AddressInfo.Line3, lCustomer.AddressInfo.Line4, lCustomer.AddressInfo.Line5, _
                pdtDeliveryDate, 0#, 0#, _
                0, 0, False, "")
            llngInvID = lrsInv!iInvoiceID
            lrsInv.Close
            
            With cmd
                '// Insert invoice detail records
                
                '// *********************************
                '// Here is one that will need to be duplicated
                .CommandText = "qkoeiInvoiceDetail"
                '// *********************************
                Do Until .Parameters.Count = 0
                    .Parameters.Delete 0
                Loop
                .Parameters.Append .CreateParameter("iIvoiceID", adInteger, adParamInput, 4, llngInvID)
                .Parameters.Append .CreateParameter("iBatchID", adInteger, adParamInput, 4, plngShipmentID)
                .Parameters.Append .CreateParameter("iCustomerID", adInteger, adParamInput, 4, llngCustomerID)
                .Execute , , adExecuteNoRecords
                
                
                '//update the order records as finished
                
                
                '// *********************************
                '// Here is one that will need to be duplicated
                .CommandText = "qkoeuFinishOrder"
                '// *********************************
                .Parameters.Delete 0
                .Parameters.Append .CreateParameter("dtPosted", adDBTimeStamp, adParamInput, 8, pdtDeliveryDate)
                .Execute , , adExecuteNoRecords
                
                lrsInv.Open "qkoegInvtotal " & llngInvID, Cn, adOpenStatic, adLockReadOnly
                If IsNull(lrsInv!mInvoiceSubtotal) Then
                    lcurInvtotal = 0
                Else
                    lcurInvtotal = lrsInv!mInvoiceSubtotal
                End If
                lrsInv.Close
                If ldblCustomerTax = 0@ Then
                    ldblTax = 0@
                Else
                    ldblTax = CCur(CLng(lcurInvtotal * ldblCustomerTax * 100)) / 100
                End If
                
                '//Now, update the Invoice header record and we're done
                .CommandText = "qkoeuFinishInvoice"
                Do Until .Parameters.Count = 0
                    .Parameters.Delete 0
                Loop
                .Parameters.Append .CreateParameter("iInvoiceID", adInteger, adParamInput, 4, llngInvID)
                .Parameters.Append .CreateParameter("mSubTotal", adDouble, adParamInput, 8, lcurInvtotal)
                .Parameters.Append .CreateParameter("mTotal", adDouble, adParamInput, 8, lcurInvtotal + ldblTax)
                .Parameters.Append .CreateParameter("mDeliveryCharge", adDouble, adParamInput, 8, 0)
                .Parameters.Append .CreateParameter("mTax", adDouble, adParamInput, 8, ldblTax)
                .Execute , , adExecuteNoRecords
            End With
            
            '// We're done with this customer, commit the transaction
            Cn.CommitTrans
        Loop
        plngOrdersAffected = rs.RecordCount
        pListener.Receive "Completed. " & plngOrdersAffected & " orders posted."
    End If
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
errhandler:
        Screen.MousePointer = MousePointerConstants.vbDefault
        '// Rollback our transaction and exit
        Cn.RollbackTrans
        HandleError "An error has occurred during Shipment Posting." & vbCrLf & _
                   "This shipment posting will be cancelled." & vbCrLf & _
                 "Details: Error Number " & Err.Number & " - " & Err.Description, False
End Sub


Public Function PostOrder(pOrder As Order, pcurDeliveryCharge As Currency, _
        pdtDeliveryDate As Date, in_ChargeSalesTax As Boolean, _
        in_UseLocalTax As Boolean, pstrSignature As String) As String
    Dim lLine           As qkorder.Line
    Dim rs              As Recordset
    Dim cmd             As Command
    Dim lcharge         As Charge
    Dim llngInvID       As Long
    Dim lcurInvSubTotal    As Currency
    Dim lcurTax         As Currency
    Dim lbolFinished    As Boolean
    Dim lstrDocNumber   As String
    
    
    On Error GoTo errhandler:
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    lbolFinished = True
    
    'Determine the Sales Tax (Dependent on whether or not it is a local pickup)
    If in_ChargeSalesTax Then
        If in_UseLocalTax Then
            With pOrder.Customer
                If .Tax > 0@ Then .Tax = CCur(Setting("local_sales_tax"))
            End With
        End If
    Else
        pOrder.Customer.Tax = 0@
    End If
    
    
    'First, if this is a new order, add a line into the database to
    'hold the special charge, if present
    'Also, flip the status flag to "In Process"
        If pOrder.State = ejsNew Then
             With pOrder
                Set lcharge = .SpecialCharge
                If lcharge.TotalCharge > 0@ Then
                    .Lines.Add 1, 0, 0, 0, 0
                    With .Lines(.Lines.Count)
                        .Charges.Add lcharge.iChargeID, _
                                lcharge.Abbrev, lcharge.Description, _
                                lcharge.Price, lcharge.Quantity, _
                                lcharge.Lineal, lcharge.LengthAdjustment
                        .ShippedNow = 1
                    End With
                End If
                .State = ejsInProcess
                .Save
            End With
        End If
    
    'Create the Invoice Header Record
        lstrDocNumber = NextIncrement("INV")
        Set rs = New Recordset
        
        rs.Open "InvoiceHeader", Cn, adOpenKeyset, adLockPessimistic
        rs.AddNew Array("vchInvoiceNumber", "iDocumentStatus", "iCustomerID", _
        "vchBillToLine1", "vchBillToLine2", "vchBillToLine3", "vchBillToLine4", _
            "dtShipDate", "mDeliveryCharge", "mInvoiceSubtotal", _
            "mTax", "mInvoiceTotal", "tiPrinted", "vchSignature"), _
               Array(lstrDocNumber, 0, pOrder.Customer.UniqueID, _
               pOrder.Customer.AddressInfo.Line2, pOrder.Customer.AddressInfo.Line3, pOrder.Customer.AddressInfo.Line4, pOrder.Customer.AddressInfo.Line5, _
            pdtDeliveryDate, pcurDeliveryCharge, lcurInvSubTotal, _
            0, 0, False, pstrSignature)
        llngInvID = rs!iInvoiceID
        rs.Close
    
    
    'Now go through and set the correct shipped amounts on each line item and add an invDetail Record
    rs.Open "InvoiceDetail", Cn, adOpenKeyset, adLockPessimistic
    
    For Each lLine In pOrder.Lines
        With lLine
            If .ShippedNow > 0 Then
                rs.AddNew Array("iInvoiceID", Constants.ORDER_HEADER_ORDER_ID, "iSerialID", "iQtyShipped", "mLineTotal"), _
                        Array(llngInvID, pOrder.OrderID, .SerialID, .ShippedNow, (.ShippedNow * .UnitPrice))
                lcurInvSubTotal = lcurInvSubTotal + (.ShippedNow * .UnitPrice)
                .Shipped = .Shipped + .ShippedNow
                .ShippedNow = 0
            End If
            If .Shipped < .Ordered Then lbolFinished = False
        End With
    Next
    rs.Close
    Set rs = Nothing
    With pOrder
        If lbolFinished Then .State = ejsFinished
        .SetProperty "dtPostedDate", Now()
        .SetProperty "iBatchID", 0
        pOrder.Save
    End With
    
    'Now, finalize the transaction by updating the InvoiceHeader summary fields
    lcurInvSubTotal = lcurInvSubTotal + pcurDeliveryCharge
    lcurTax = CCur(CLng(lcurInvSubTotal * pOrder.Customer.Tax * 100) / 100)
    Set cmd = New Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "qkuInvHeader"
        Set .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("iInvoiceID", adInteger, adParamInput, 4, llngInvID)
        .Parameters.Append .CreateParameter("mDeliveryCharge", adDouble, adParamInput, 4, pcurDeliveryCharge)
        .Parameters.Append .CreateParameter("mInvoiceSubtotal", adDouble, adParamInput, 4, lcurInvSubTotal)
        .Parameters.Append .CreateParameter("mTax", adDouble, adParamInput, 8, lcurTax)
        .Parameters.Append .CreateParameter("mInvoiceTotal", adDouble, adParamInput, 8, lcurTax + lcurInvSubTotal)
        .Execute , , adExecuteNoRecords
    End With
    Set cmd = Nothing
    Set pOrder = Nothing
    Screen.MousePointer = MousePointerConstants.vbDefault
    PostOrder = lstrDocNumber
    Exit Function
errhandler:
        Screen.MousePointer = MousePointerConstants.vbDefault
        HandleError "Order Posting Failed. " & Err.Description, False
End Function

Public Function GetRecentSignatures(plngCustomerID As Long) As Recordset
    Set GetRecentSignatures = GetData("gqcgSignatureList " & plngCustomerID)
End Function

Public Function GetPOBindingRecordSet(plngPONumber As Long) As Recordset
    
    Set GetPOBindingRecordSet = GetData("Select *, [iQuantity] - [iQuantityReceived] as [iRemaining] from PurchaseOrderDetail where iPurchaseOrderID = " & plngPONumber, adLockBatchOptimistic)
    
End Function

Public Sub ReconnectAndUpdate(in_RS As Recordset)
    
    Set in_RS.ActiveConnection = Cn
    in_RS.UpdateBatch
    Set in_RS.ActiveConnection = Nothing
    
End Sub
Public Sub UpdatePOStatusOnOrder(inOrderId As Long, inPOStatus As POStatusEnum)
    Dim cmd As New Command
    With cmd
        .CommandType = adCmdText
        .CommandText = "Update OrderHeader Set tiPOStatusID = " & inPOStatus & " Where iOrderID = " & inOrderId
        Set .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
        Set .ActiveConnection = Nothing
    End With
    Set cmd = Nothing
End Sub

'Public Function GetOrderIDsOnPurchaseOrder(inPurchaseOrderID As Long) As Recordset
'    Set GetOrderIDsOnPurchaseOrder = GetDataFromSingularCommand("qkpogOrdersOnPurchaseOrder", "iPurchaseOrderID", adInteger, 4, inPurchaseOrderID)
'End Function

Public Function GetOrderPOStatusResults() As Recordset
    Set GetOrderPOStatusResults = GetData("qkpogOrderPOStatusResults")
End Function

Public Sub GenerateInvoiceBatch(llngBatchid As Long, llngGeneratedCount As Long)
    Dim cmd As Command
    Dim lstrNewBatchID As String
    
    lstrNewBatchID = NextIncrement("BAT")
    llngBatchid = CLng(Mid$(lstrNewBatchID, 4))
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        .CommandText = "qkuInvoiceBatch"
        .Parameters.Append .CreateParameter("iCurrentBatchID", adInteger, adParamInput, 4, 0)
        .Parameters.Append .CreateParameter("iUpdateBatchID", adInteger, adParamInput, 4, llngBatchid)
        .Parameters.Append .CreateParameter("iStatus", adInteger, adParamInput, 4, 1)
        .Execute llngGeneratedCount
    End With

End Sub
Public Sub FinishInvoiceBatch(plngBatchID As Long)
    Dim cmd As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        .CommandText = "qkuInvoiceBatch"
        .Parameters.Append .CreateParameter("iCurrentBatchID", adInteger, adParamInput, 4, plngBatchID)
        .Parameters.Append .CreateParameter("iUpdateBatchID", adInteger, adParamInput, 4, plngBatchID)
        .Parameters.Append .CreateParameter("iStatus", adInteger, adParamInput, 4, 2)
        .Execute , , adExecuteNoRecords
    End With
End Sub

Public Function GetBatchInfo(plngBatchID As Long) As Recordset
    Set GetBatchInfo = GetData("qkogBatchPrintingInfo " & plngBatchID)
End Function
Public Sub DeleteShipment(plngBatchID As Long)
    Cn.Execute "qkoeDeleteShipment " & plngBatchID, , adExecuteNoRecords
End Sub
Public Sub CreateShipment(pDtRequested As Date, plngRecordsAffected As Long, plngBatchID As Long, Optional plngCustomerID As Long = 0)

    Dim lstr As String
    lstr = NextIncrement("BAT")
    plngBatchID = CLng(Mid$(lstr, 4))
    
    Dim cmd As Command
        Set cmd = New Command
        With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("dtRequested", adDBTimeStamp, adParamInput, 8, Format$(pDtRequested, "YYYY/MM/DD"))
        .Parameters.Append .CreateParameter("iBatchID", adInteger, adParamInput, 4, plngBatchID)
        If plngCustomerID > 0 Then
            .CommandText = "qkoeCreateShipmentbyCust"
            .Parameters.Append .CreateParameter("iCustomerID", adInteger, adParamInput, 4, plngCustomerID)
        Else
            .CommandText = "qkoeCreateShipmentbyDate"
        End If
        .Execute plngRecordsAffected
    End With
End Sub

Public Function DeleteOrder(plngOrderID As Long) As String
    Dim rs As Recordset
    Dim aReturnString As String
    
    Set rs = Me.GetOrderHeader(plngOrderID)
        
    If rs!iStatus = 0 Then
        Cn.Execute "qkoedOrder " & plngOrderID, , adExecuteNoRecords
        aReturnString = "Order status was ""NEW""." & vbCrLf & _
                "Entire Order was deleted."
    Else
        Cn.Execute "qkoeCancelOrder " & plngOrderID, , adExecuteNoRecords
        aReturnString = "Order has been partially posted." & vbCrLf & _
                "Remaining Portion has been canceled."
    End If
    rs.Close
    Set rs = Nothing
    
    DeleteOrder = aReturnString
End Function



Private Function GetData(in_Query As String, Optional in_LockType As LockTypeEnum = adLockReadOnly) As Recordset
    Dim lrs As Recordset
    Set lrs = New Recordset
    With lrs
        .CursorLocation = adUseClient
        .Open in_Query, Cn, adOpenStatic, in_LockType
        Set .ActiveConnection = Nothing
    End With
    Set GetData = lrs
End Function

Private Function GetDataFromSingularCommand( _
        in_CommandText As String, _
        in_ParamName As String, _
        in_Type As DataTypeEnum, _
        in_ParamLength As Long, _
        in_ParamVal As Variant)
    
    Dim aCmd   As Command
    Set aCmd = New Command
    With aCmd
        Set .ActiveConnection = Cn
        .CommandText = in_CommandText
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(in_ParamName, in_Type, adParamInput, in_ParamLength, in_ParamVal)
    End With
    Set GetDataFromSingularCommand = GetDataFromCmd(aCmd)
    
End Function

Private Function ExecuteSingularCommand(in_CommandText As String, in_ParamName As String, _
                        in_Type As DataTypeEnum, in_ParamLength As Long, _
                        in_ParamVal As Variant)

    Dim aCmd   As Command
    Set aCmd = New Command
    With aCmd
        Set .ActiveConnection = Cn
        .CommandText = in_CommandText
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(in_ParamName, in_Type, adParamInput, in_ParamLength, in_ParamVal)
    End With
    
    aCmd.Execute Options:=adExecuteNoRecords
    
End Function

Private Function GetDataFromCmd(in_Cmd As Command) As Recordset
    Dim lrs As Recordset
    Set lrs = New Recordset
    With lrs
        .CursorLocation = adUseClient
        .Open in_Cmd, , adOpenForwardOnly, adLockReadOnly
        Set .ActiveConnection = Nothing
    End With
    Set GetDataFromCmd = lrs
End Function

Private Function ExecuteScalar(in_Query As String, Optional in_Default As Variant = Empty) As Variant
    Dim rs As Recordset
    Set rs = New Recordset
    rs.Open in_Query, Cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        ExecuteScalar = in_Default
    Else
        If IsNull(rs(0)) Then
            ExecuteScalar = in_Default
        Else
            ExecuteScalar = rs(0)
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function GetStockPOTotals() As Integer
     GetStockPOTotals = ExecuteScalar("qkPOGStockOrderTotals", 0)
End Function

Public Function GetPendingPOTotals(in_MaxRequestDate As Date) As Recordset
    Set GetPendingPOTotals = GetData("qkPOGPendingOrderReport '" & FormatMaxDate(in_MaxRequestDate) & "'", adLockBatchOptimistic)
End Function

Public Function getInvoicebyNumber(in_InvoiceNumber As String) As String
   getInvoicebyNumber = ExecuteScalar("qkoegInvoiceByNumber '" & in_InvoiceNumber & "'", "")
End Function
Public Function getInvoicebyOrder(in_OrderNumber As String) As String
    getInvoicebyOrder = ExecuteScalar("qkoegInvoiceByOrderNumber '" & in_OrderNumber & "'", "")
End Function
Public Function getPOID(in_PurchaseOrderNumber As String) As Long
    getPOID = ExecuteScalar("Select iPurchaseOrderID from PurchaseOrder where vchPONumber = '" & in_PurchaseOrderNumber & "'", 0)
End Function

'// getInvoiceList() as Long()
' Returns a list of invoice IDS that match the filter criteria
Public Function getInvoiceListByFilters(iCustomerID As Long, in_Date As Date, in_Filter As String) As String()
Dim rs As Recordset
Dim cmd As Command
Dim lintCounter As Integer
Dim lReturn() As String

Set cmd = New Command
With cmd
    .CommandText = "qkoegInvoicebyFilters"
    .CommandType = adCmdStoredProc
    .ActiveConnection = Cn
    .Parameters.Append .CreateParameter("iCustomerID", adInteger, adParamInput, 4, iCustomerID)
    .Parameters.Append .CreateParameter("dtOrderDate", adDBTimeStamp, adParamInput, Value:=in_Date)
    .Parameters.Append .CreateParameter("vchFilter", adVarChar, adParamInput, 255, Value:=in_Filter)
End With

Set rs = New Recordset
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
If (rs.RecordCount > 0) Then
    ReDim Preserve lReturn(rs.RecordCount - 1)
    Do Until rs.EOF
        lReturn(lintCounter) = CStr(rs(0))
        lintCounter = lintCounter + 1
        rs.MoveNext
    Loop
Else
    ReDim lReturn(0)
    lReturn(0) = "-1"
End If
getInvoiceListByFilters = lReturn
End Function

Public Sub GetLabelReportInfo(ByRef out_RecordsAffected As Long, _
        ByRef out_ReportSQL As String, _
        ByRef out_DBSQL As String, _
        in_OnlyDoPrinted As Boolean, _
        Optional ByVal in_OrderID As Long = 0, _
        Optional ByVal in_ShipmentID As Long = 0)
    Dim rs As Recordset
    
    If Not (in_OrderID = 0) Then
        out_ReportSQL = "{vOrderHeader.iOrderID} = " & in_OrderID & " and " & _
                "{vOrderHeader.iStatus} < 2.00 and " & _
                "{OrderLine.iLaminateID} <> 0"
    Else
        out_ReportSQL = "{vOrderHeader.iBatchID} = " & in_ShipmentID & " and " & _
                "{vOrderHeader.iStatus} < 2.00 and " & _
                "{OrderLine.iLaminateID} <> 0"
    End If
    If in_OnlyDoPrinted Then
        out_ReportSQL = out_ReportSQL & " and {OrderLine.tiLblPrinted} = false"
    End If
    
    out_DBSQL = Replace(out_ReportSQL, "{", "(")
    out_DBSQL = Replace(out_DBSQL, "}", ")")
    out_DBSQL = Replace(out_DBSQL, " and ", ") and (")
    
    'Run a query to get the totals
    Set rs = Cn.Execute("Select Count(*) from vOrderHeader INNER JOIN OrderLine " & _
            "ON vOrderHeader.iOrderID = OrderLine.iOrderID " & _
            "WHERE ((" & out_DBSQL & "))")
            
    out_RecordsAffected = rs(0)
    rs.Close
    
    'Now finish building the update query
    out_DBSQL = "UPDATE vOrderHeader INNER JOIN OrderLine " & _
            "ON vOrderHeader.iOrderID = OrderLine.iOrderID " & _
            "SET OrderLine.tiLblPrinted = True " & _
            "WHERE ((" & out_DBSQL & "))"

End Sub
Public Function UpdateLabels(in_Query As String) As Long
    Dim out_RecordsAffected As Long
    Cn.Execute in_Query, out_RecordsAffected, adExecuteNoRecords
    UpdateLabels = out_RecordsAffected
End Function


Public Function GetVendorOrderableUnits(in_VendorSlabTypeID As Long) As OrderableUnit()
    Dim rs As New Recordset
    Dim aReturnVal() As OrderableUnit
    Dim aCounter As Integer
    Dim aUnit As OrderableUnit
    
    rs.CursorLocation = adUseClient
    rs.Open "Select * from VendorSlabDetail where iVendorSlabTypeID = " & in_VendorSlabTypeID, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    aCounter = 1
    ReDim aReturnVal(1 To rs.RecordCount)
    While Not rs.EOF
        Set aUnit = New OrderableUnit
        aUnit.OrderableUnitItem rs!iLength, QKBusinessRules.GetOrderableUnitLengthInInches(rs!iLength), rs!cPrice
        Set aReturnVal(aCounter) = aUnit
        aCounter = aCounter + 1
        rs.MoveNext
    Wend
    
    rs.Close
    
    GetVendorOrderableUnits = aReturnVal
End Function

Public Sub InsertPODetailLines(in_Items As RawUnitItems, in_GroupByCode As String, in_VendorCode As String, in_Description As String, in_PO As Long)
    Dim aLineItem As RawUnitItem
    Dim aRs As Recordset
    Dim aCurrentPODetailID As Long
    
    Set aRs = New Recordset
    aRs.CursorLocation = adUseServer
    aRs.Open "PurchaseOrderDetail", Cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
    For Each aLineItem In in_Items
        aCurrentPODetailID = InsertPODetailLine(in_PO, in_GroupByCode, in_VendorCode, in_Description, (aLineItem.OrderableUnit.SlabLength), 1, aLineItem.ItemNumber, aLineItem.Length, aLineItem.SerialID, aRs)
        While Not aLineItem.Child Is Nothing
            Set aLineItem = aLineItem.Child
            AddDetailLineMapping aLineItem.SerialID, aCurrentPODetailID, aLineItem.ItemNumber
        Wend
    Next
    
    aRs.Close
    
End Sub

Public Function InsertPODetailLine(in_PO As Long, in_GroupByCode As String, _
                    in_VendorCode As String, in_Description As String, _
                    in_Length As Double, in_Quantity As Long, in_ItemNumber As Byte, _
                    inLengthUsed As Currency, in_SerialId As Long, _
                    Optional in_RS As Recordset) As Long
    Dim aReturnVal As Long
    Dim aExternal As Boolean
    
    If in_SerialId = 0 Then
        Cn.Execute "Insert Into PurchaseOrderDetail (iPurchaseOrderID, vchGroupByCode, vchVendorItemCode, vchItemDescription, flSize, iQuantity, tiIsSlab ) " & _
                "Values ( " & in_PO & ", '" & CleanDBString(in_GroupByCode, "'") & "', '" & in_VendorCode & "', '" & _
                CleanDBString(in_Description, "'") & "', " & in_Length & ", " & in_Quantity & ", False )"
    Else
        Dim aRs As Recordset
        If in_RS Is Nothing Then
            Set aRs = New Recordset
            aRs.CursorLocation = adUseServer
            aRs.Open "PurchaseOrderDetail", Cn, adOpenKeyset, adLockPessimistic, adCmdTable
        Else
            aExternal = True
            Set aRs = in_RS
        End If
        With aRs
            .AddNew Array("iPurchaseOrderID", "vchGroupByCode", "vchVendorItemCode", "vchItemDescription", "flSize", "iQuantity", "tiIsSlab", "dcLengthUsed"), _
                Array(in_PO, in_GroupByCode, in_VendorCode, in_Description, in_Length, in_Quantity, True, inLengthUsed)
            aReturnVal = !iPurchaseOrderLineId
            AddDetailLineMapping in_SerialId, aReturnVal, in_ItemNumber
            If Not aExternal Then
                .Close
            End If
        End With
    End If
    InsertPODetailLine = aReturnVal
End Function

Public Sub MovePODetailLines(in_OldID As Long, in_NewID As Long)
    Dim cmd As New Command
    
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "qkpouMoveLineItems"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("iOldID", adInteger, adParamInput, 4, in_OldID)
        .Parameters.Append .CreateParameter("iNewID", adInteger, adParamInput, 4, in_NewID)
        .Execute , , adExecuteNoRecords
    End With
    
    Set cmd = Nothing
End Sub

Public Function GetChargesForBrandCodeOnOrder(in_OrderID As Long, in_BrandCode As String) As Recordset
Dim cmd As New Command
    
    With cmd
        Set .ActiveConnection = Cn
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter(Constants.ORDER_HEADER_ORDER_ID, adInteger, adParamInput, 4, in_OrderID)
        If in_BrandCode = STOCK_BRAND Then
            .CommandText = "qkpogStockCodeCharges"
        Else
            .CommandText = "qkpogBrandCodeCharges"
            .Parameters.Append .CreateParameter("vchBrandCode", adVarChar, adParamInput, 4, in_BrandCode)
        End If
    End With
    Set GetChargesForBrandCodeOnOrder = GetDataFromCmd(cmd)
    
End Function

Public Function getChargesForSpecialVendorItemsOnOrder(in_OrderID As Long)
    Set getChargesForSpecialVendorItemsOnOrder = GetDataFromSingularCommand("qkpogSpecialVendorItemCharges", "in_OrderID", adInteger, 4, in_OrderID)
End Function

Public Sub AddDetailLineMapping(in_SerialId As Long, _
        in_PurchaseOrderLineId As Long, _
        in_ItemCountNumber As Byte)
    Cn.Execute "Insert Into PurchaseOrderDetailMapping (iSerialID, iPurchaseOrderLineID, iItemNumber) Select " & in_SerialId & " As Expr1," & in_PurchaseOrderLineId & " as Expr2, " & in_ItemCountNumber & " as Expr3"
    Cn.Execute "Update OrderLine set tiOnPo = True where iSerialID = " & in_SerialId
End Sub
Public Sub CombinePurchaseOrderDetailLines(inNewPurchaseOrderLineID As Long, inOldPurchaseOrderLineID As Long)
    Cn.Execute "Update PurchaseOrderDetailMapping Set iPurchaseOrderLineID = " & inNewPurchaseOrderLineID & " where iPurchaseOrderLineID = " & inOldPurchaseOrderLineID
    Me.RemovePODetailLine (inOldPurchaseOrderLineID)
End Sub

Public Sub UpdatePurchaseOrderDetailOrderedSize(in_PurchaseOrderLineId As Long, in_Size As Integer)
    Cn.Execute "Update PurchaseOrderDetail Set flSize = " & in_Size & " Where iPurchaseOrderLineID = " & in_PurchaseOrderLineId
End Sub

Public Function UpdatePurchaseOrderDetailToReceived(in_PurchaseOrderLineId As Long) As Boolean
    Dim aAffected As Integer
    Cn.Execute "Update PurchaseOrderDetail Set iQuantityReceived = iQuantity Where iPurchaseOrderLineID = " & in_PurchaseOrderLineId, aAffected, adExecuteNoRecords
    UpdatePurchaseOrderDetailToReceived = aAffected > 0
End Function

Public Sub UpdateOrderLineOnPOToFalse(in_PurchaseOrderLineId As Long)
    ExecuteSingularCommand "qkpoUpdatePOLineStatus ", "iPurchaseOrderLineId", adInteger, 4, in_PurchaseOrderLineId
End Sub

Public Function GetBulkBatchChargesInOrder() As Recordset
    Set GetBulkBatchChargesInOrder = GetData("qkpoGBatchChargesInOrder", adLockReadOnly)
End Function

Public Sub RemovePODetailLine(in_PurchaseOrderLineId As Long)
    Cn.Execute "Delete from PurchaseOrderDetail where iPurchaseOrderLineId = " & in_PurchaseOrderLineId
End Sub



Public Function GetSerialIDForLineNumber(in_OrderID As Long, in_LineNumber As Integer)
    GetSerialIDForLineNumber = ExecuteScalar("Select iSerialID from OrderLine Where iOrderid = " & _
                in_OrderID & " and iLineNumber = " & in_LineNumber, 0)
End Function

Public Function GetPendingPOFootage(inMaxDate As Date, inPODate As Date) As PendingPORequest
Dim aReturnVal As New PendingPORequest
Dim aFootage As Integer

    aReturnVal.MaxDate = inMaxDate
    aReturnVal.PODate = inPODate
    
    Dim aRs As Recordset
    
    aFootage = DataCenter.GetStockPOTotals
    
    Set aRs = DataCenter.GetPendingPOTotals(inMaxDate)
    If Not aRs.EOF Then
        Do While Not aRs.EOF
            aFootage = aFootage + aRs!TotalFootage
            aRs.MoveNext
        Loop
        aReturnVal.TotalFootage = aFootage
        Set aReturnVal.Data = aRs
    End If
    Set GetPendingPOFootage = aReturnVal
    
End Function

Private Function FormatMaxDate(in_Date As Date) As Date
    FormatMaxDate = FormatDateTime(in_Date, vbShortDate) & " 23:59:59"
End Function

'Public Function GetUser(inUserName As String, inPassword As String) As CurrentUser
'    Dim aCmd As New Command
'    Dim aRS As Recordset
'    Set aCmd.ActiveConnection = Cn
'    aCmd.CommandType = adCmdText
'    aCmd.CommandText = "Select * from OrderRep where vchUserName = '" & inUserName & "' and vchPassword = '" & inPassword & "'"
'    Set aRS = GetDataFromCmd(aCmd)
'    If aRS.RecordCount > 0 Then
'        Dim aCurrrentUser As CurrentUser
'        Set aCurrrentUser = New CurrentUser
'        aCurrrentUser.Init aRS!iRepID, aRS!vchFirstName, aRS!vchLastName, aRS!iUserEnableFlags
'        Set GetUser = aCurrrentUser
'    End If
'End Function

#If DBUG Then
    Private Sub mCn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
        Dim aParam As Parameter
        Debug.Print "Execute: " & pCommand.CommandText & " : " & RecordsAffected
        If pCommand.NamedParameters Then
            For Each aParam In pCommand.Parameters
                Debug.Print aParam.Name & " = " & aParam.Value,
            Next
            Debug.Print ""
        End If
    End Sub
    
    Private Sub mCn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
        Debug.Print "ERROR: " & pError.Description
    End Sub
#End If
