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
Private m_lngSelectedJobID      As Long
Public SelectedJobInfo          As String

'Strings
Private m_strDataPath            As String
Private m_strOrderPath           As String
Private m_strInvoicePath         As String
Private m_strExtension           As String

'Numerics
Private minthwndControl         As Integer

'Objects
Private mRsOEReps               As Recordset
Private mrsSlabs                As Recordset
Private m_LineCommand         As Command
Private m_Chargecommand       As Command

Private m_colSettings            As Dictionary

Private mCn                     As ADODB.Connection

Public Property Let SelectedJobID(plngID As Long)
    m_lngSelectedJobID = plngID
    SelectedJobInfo = ""
End Property
Public Property Get SelectedJobID() As Long
    SelectedJobID = m_lngSelectedJobID
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
Dim m_varfields As Variant
Dim m_vardata As Variant

m_varfields = Array("iCustomerID", "iPriceListID", "vchOrderNumber", "vchRefNumber", "dtEntryDate", "dtUpdateDate", "dtPostedDate", "iStatus", "iOrderType", _
        "dtRequestDate", "vchJobName", "vchPONumber", "vchShipAddress1", _
        "vchShipAddress2", "vchShipAddress3", "vchShipAddress4", "vchShipAddress5", "tiCartoned", "tiPadded", "tiSplined", _
        "vchTrackingCode", "iUserID", "mOrderTotal", "txtComment", "tiGroup", "iBatchID")
With Order
    m_vardata = Array(.Customer.UniqueID, .PriceListID, .OrderNumber, .GetProperty("vchRefNumber"), .Entrydate, Now(), .GetProperty("dtPostedDate"), .State, .JobType, .RequiredDate, .GetProperty("vchJobName"), .PO, .ShipTo(1), _
        .ShipTo(2), .ShipTo(3), .ShipTo(4), .ShipTo(5), .Cartoned, .Padded, .Splined, .Trackingcode, .UserID, .Total, .GetProperty("txtComment"), .GetProperty("tiGroup"), .GetProperty("iBatchID"))
End With
Set rs = New Recordset
If Order.OrderID = 0 Then
        rs.CursorLocation = adUseServer
        rs.Open "Select * from vorderheader where iOrderID = 0", Cn, adOpenDynamic, adLockPessimistic
        rs.AddNew m_varfields, m_vardata
        rs.MoveLast
        Order.OrderID = rs("iOrderID")
Else
        rs.CursorLocation = adUseServer
        rs.Open "Select * from vOrderHeader where iOrderID = " & Order.OrderID, Cn, adOpenDynamic, adLockPessimistic
        rs.Update m_varfields, m_vardata
End If
        
rs.Close
Set rs = Nothing
If Order.Lines.DeletedLines.Count > 0 Then
    For Each m_vardata In Order.Lines.DeletedLines
        Cn.Execute "qkdOrderLine " & m_vardata
    Next
End If
End Sub

Public Function GetCustomerList() As Recordset
Dim rs As Recordset

Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open "qkoegCustomerList", _
        Cn, adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
End With

Set rs.ActiveConnection = Nothing

Set GetCustomerList = rs
End Function

Public Function GetAddress(in_Address As Integer) As AddressInfo
Dim rs                  As Recordset
Dim cmd                 As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkoegAddress"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iAddressid", adInteger, adParamInput, 2, in_Address)
End With

Set rs = New Recordset

rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

Dim a_Address As New AddressInfo

With rs
    a_Address.PreAddressInfo = !PreAddressInfo & ""
    a_Address.Name = !vchLocationName & ""
    a_Address.Address1 = !vchAddress1 & ""
    a_Address.Address2 = !vchAddress2 & ""
    a_Address.City = !vchCity & ""
    a_Address.State = !vchState & ""
    a_Address.Zip = !vchZip & ""
    a_Address.Phone = !vchTelephoneNumber & ""
    .Close
End With

Set GetAddress = a_Address

Set rs = Nothing
Set cmd = Nothing
End Function

Public Function GetCustomerShipLocations(in_CustomerID As Long) As Recordset
Dim rs                  As Recordset
Dim cmd                 As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkoegCustomerShipLocations"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iCustomerid", adInteger, adParamInput, 2, in_CustomerID)
End With

Set rs = New Recordset

rs.Open cmd, , adOpenForwardOnly, adLockReadOnly


Set GetCustomerShipLocations = rs

Set cmd = Nothing

End Function

Public Function GetSlabList() As Recordset

If mrsSlabs Is Nothing Then
    Set mrsSlabs = New Recordset
    With mrsSlabs
        .CursorLocation = adUseClient
        .Open "select iSlabID,vchSlabCode,vchslabDesc from qkgslabinfo", _
                Cn, adOpenForwardOnly, adLockReadOnly
        Set .ActiveConnection = Nothing
    End With
Else
    mrsSlabs.MoveFirst
End If

Set GetSlabList = mrsSlabs
End Function


Public Function LoadOrderTypes() As Recordset
Dim lrsOrderTypes As Recordset

Set lrsOrderTypes = New Recordset
With lrsOrderTypes
    .CursorLocation = adUseClient
    .Open "Select * from qkgordertypes", _
            Cn, adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
End With

Set LoadOrderTypes = lrsOrderTypes
End Function

Public Function GetCustomer(iCustomerID As Long) As Recordset
Dim rs                  As Recordset
Dim cmd                 As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkcgCustomer"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iCustomerid", adInteger, adParamInput, 2, iCustomerID)
End With

Set rs = New Recordset

rs.Open cmd, , adOpenForwardOnly, adLockReadOnly



Set GetCustomer = rs

Set rs = Nothing
Set cmd = Nothing
End Function

Public Function GetColorMatches(lstrSearchStr As String, Optional pstrProcedureName As String = "qklgColorMatches") As Recordset
Dim rs                  As Recordset
Dim cmd                 As Command

On Error GoTo errhandler
Screen.MousePointer = vbHourglass

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = pstrProcedureName
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("vchSearchString", adVarChar, adParamInput, 30, lstrSearchStr)
End With

Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
End With

Set GetColorMatches = rs

Set rs = Nothing
Set cmd = Nothing

Screen.MousePointer = vbArrow
Exit Function
errhandler:
    Screen.MousePointer = vbArrow
    HandleError Err.Description, False
End Function

Public Sub Connect()
Dim rs As Recordset



If m_colSettings Is Nothing Then
    Set rs = New Recordset
    With rs
        .CursorLocation = adUseClient
        .Open "gkgappsettings", Cn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
        Set m_colSettings = New Dictionary
        m_colSettings.CompareMode = TextCompare
        Do Until .EOF
            m_colSettings.Add !vchSettingKey.Value, !vchSetting.Value
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
End If

'Retrieve our path settings specific to our data files
If Len(m_strExtension) = 0 Then
    m_strExtension = GetIniSetting("General", "JobDataExtension", "qll")
    m_strDataPath = App.Path
    m_strInvoicePath = GetSetting("QK", "General", "InvoiceFolder", "c:\")
    m_strOrderPath = GetIniSetting("General", "OrderPath", "")
End If

End Sub

Public Function Setting(pstrSettingKey As String) As String
    If m_colSettings.Exists(pstrSettingKey) Then
        Setting = m_colSettings(pstrSettingKey)
    Else
        Setting = ""
    End If
End Function
Private Sub lConnect()

With mCn
    .ConnectionString = GetIniSetting("General", "Connect", "DSN=QKORDER;")
    .Open
End With

End Sub


Public Sub ShowOpen(pintOrderType As Integer, Optional ForPosting As Boolean = False)
    frmLoadJobs.Init pintOrderType, ForPosting
    If SelectedJobID = 0 Then frmLoadJobs.Show vbModal
End Sub

Public Function GetOrders(plngOrderType As Integer) As Recordset
Dim rs                  As Recordset
Dim cmd                 As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkogOrdersByType"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iOrderType", adInteger, adParamInput, 2, plngOrderType)
End With

Set rs = New Recordset
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenDynamic, adLockReadOnly
Set rs.ActiveConnection = Nothing

Set GetOrders = rs

Set rs = Nothing
Set cmd = Nothing
End Function

Public Function GetOrderHeader(OrderID As Long) As Recordset
Dim rs                  As Recordset
Dim cmd                 As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkogOrderHeader"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iOrderID", adInteger, adParamInput, 2, OrderID)
End With
Set rs = New Recordset
rs.Open cmd, , adOpenForwardOnly, adLockReadOnly



Set GetOrderHeader = rs

Set rs = Nothing
Set cmd = Nothing
End Function

Public Function GetOeLinesForPOLine(iPODetailLineID As Long) As Recordset
    Dim rs                  As Recordset
    Dim cmd                 As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "qkPOGOELinesForPOLine"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("iPurchaseOrderLineID", adInteger, adParamInput, 2, iPODetailLineID)
    End With
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    
    Set GetOeLinesForPOLine = rs

    Set cmd = Nothing
End Function



'Public Function GetLine(plngSerialID As Long) As Recordset
'Dim rs          As Recordset
'
'If m_LineCommand Is Nothing Then
'    Set m_LineCommand = New Command
'    With m_LineCommand
'        Set .ActiveConnection = Cn
'        .CommandText = "qkogOrderLine"
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("iSerialID", adInteger, adParamInput, 2)
'    End With
'End If
'
'm_LineCommand!iSerialID = plngSerialID
'
'Set rs = New Recordset
'rs.Open m_LineCommand, , adOpenForwardOnly, adLockReadOnly
'
'
'Set GetLine = rs
'
'End Function

Public Sub SaveLine(iOrderId As Long, pLine As qkorder.Line)
Dim rs              As Recordset
Dim lbolNewLine     As Boolean
Dim m_varfields      As Variant
Dim m_vardata        As Variant

m_varfields = Array("iOrderID", "iSlabID", "txtLineDesc", "iLineNumber", "flSlabLength", "flSlabWidth", _
    "flSlabArea", "flSlabPrice", "iLaminateID", "iOrdered", "iShipped", _
    "iPosted", "flDiscountPercent", "flLaminateExtPrice", _
    "flChargesExtPrice", "flLineExtPrice")
With pLine
    m_vardata = Array(iOrderId, .SlabID, .LineDescription, .LineNumber, .SlabLength, .SlabWidth, .LinealFeet, _
    .SlabPrice, .LaminateID, .ordered, .Shipped, .Posted, 0, _
   .LaminateUpcharge, .Charges.ExtPrice, .Total)
End With

Set rs = New Recordset
If pLine.SerialID > 0 Then
    rs.Open "Select * from vorderline where iSerialID = " & pLine.SerialID, Cn, adOpenKeyset, adLockPessimistic
    rs.Update m_varfields, m_vardata
Else
    lbolNewLine = True
    rs.Open "select * from vorderline where iOrderID = " & iOrderId, Cn, adOpenKeyset, adLockPessimistic
    rs.AddNew m_varfields, m_vardata
    pLine.SerialID = rs!iSerialID
End If

rs.Close
Set rs = Nothing
SaveCharges pLine, lbolNewLine

End Sub

Public Sub CreatePO(in_RequestedDate As Date, in_VendorID As Long, ByRef out_DocPONumber As String, ByRef out_DocID As Long)
Dim rs As New Recordset

    out_DocPONumber = NextIncrement("PO")
    rs.Open "PurchaseOrder", Cn, adOpenKeyset, adLockPessimistic, adCmdTable
    rs.AddNew Array("vchPONumber", "iVendorID", "dtRequested", "dtDateOrdered"), Array(out_DocPONumber, in_VendorID, in_RequestedDate, Now)
    out_DocID = rs!iPurchaseOrderID
    rs.Close

End Sub

Private Sub SaveCharges(pLine As qkorder.Line, pbolFirstSave As Boolean)
Dim rs              As Recordset
Dim m_varfields      As Variant
Dim m_vardata        As Variant
Dim lcharge         As Charge
Dim lstrUpdateString    As String

m_varfields = Array("iSerialID", "iChargeID", "iQuantity", "vchChargeCode", "vchChargeDesc", "flChargePrice")
Set rs = New Recordset

rs.Open "Select * from vorderlinecharge where iSerialID = " & pLine.SerialID, Cn, adOpenKeyset, adLockPessimistic


For Each lcharge In pLine.Charges
    With lcharge
        m_vardata = Array(pLine.SerialID, .iChargeID, .Quantity, .Abbrev, Left$(.Description, 50), .Price)
        If .iLineChargeID > 0 Then
            rs.Filter = "iLineChargeID=" & .iLineChargeID
            If Not rs.EOF Then
                rs.Update m_varfields, m_vardata
            Else
                rs.AddNew m_varfields, m_vardata
                .iLineChargeID = rs!iLineChargeID
            End If
        Else
            rs.AddNew m_varfields, m_vardata
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
Dim rs          As Recordset
Dim cmd         As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "gkigSlab"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iSlabID", adInteger, adParamInput, 4, iSlabid)
    .Parameters.Append .CreateParameter("iPriceListID", adInteger, adParamInput, 4, iPriceListID)
End With
Set rs = New Recordset
rs.Open cmd, , adOpenForwardOnly, adLockReadOnly



Set GetSlab = rs

End Function

Public Function GetSpecialOrderItemsOnOrder(iOrderId As Long) As Recordset
Dim rs          As Recordset
Dim cmd         As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkpogSpecialColorItemsOnOrder"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iOrderID", adInteger, adParamInput, 4, iOrderId)
End With
Set rs = New Recordset
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

Set rs.ActiveConnection = Nothing

Set GetSpecialOrderItemsOnOrder = rs
End Function

'Note that this one will keep the cursor open on the server as the datasets could be large
Public Function GetGeneralPurchaseOrderItems(in_MaxRequiredDate As Date) As Recordset
    Dim rs          As Recordset
    Dim cmd         As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "qkpogGeneralPurchaseOrderItems"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("dtOrderRequriedDate", adDBTimeStamp, adParamInput, Value:=in_MaxRequiredDate)
    End With
    Set rs = New Recordset
    rs.CursorLocation = adUseServer
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set GetGeneralPurchaseOrderItems = rs
End Function

Public Function GetGeneralPurchaseOrderCharges(in_MaxRequiredDate As Date) As Recordset
    Dim rs          As Recordset
    Dim cmd         As Command
    
    Set cmd = New Command
    With cmd
        Set .ActiveConnection = Cn
        .CommandText = "qkpogGeneralPurchaseOrderCharges"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("dtOrderRequriedDate", adDBTimeStamp, adParamInput, Value:=in_MaxRequiredDate)
    End With
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    
    Set GetGeneralPurchaseOrderCharges = rs
End Function


Public Function GetCharge(pstrChargeCode As String, iPriceListID As Long) As Charge
Dim rs          As Recordset
Dim lcharge     As Charge

Set lcharge = New Charge

If m_Chargecommand Is Nothing Then
    Set m_Chargecommand = New Command
    With m_Chargecommand
        Set .ActiveConnection = Cn
        .CommandText = "qkigCharge"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("vchChargeCode", adVarChar, adParamInput, 20, pstrChargeCode)
        .Parameters.Append .CreateParameter("iPriceListID", adInteger, adParamInput, 4, iPriceListID)
    End With
Else
    m_Chargecommand!vchChargeCode = pstrChargeCode
    m_Chargecommand!iPriceListID = iPriceListID
End If

Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open m_Chargecommand, , adOpenForwardOnly, adLockReadOnly

    If Not .EOF Then lcharge.Init !iChargeID, 0, !vchChargeCode, _
        !flPrice, !vchChargeDesc, (!tiLineal = 1), !flLengthAdj
End With
rs.Close
Set rs = Nothing
Set GetCharge = lcharge

End Function
Public Function GetLaminate(iLaminateID As Long, iPriceListID As Long, iSlabid As Long) As Recordset
Dim rs          As Recordset
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
Set rs = New Recordset
rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

Set GetLaminate = rs

End Function

Public Function GetLinesOnOrder(iOrderId As Long) As Recordset
Dim rs          As Recordset
Dim cmd         As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkogLinesOnOrder"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iOrderID", adInteger, adParamInput, 2, iOrderId)
End With
Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
End With
Set GetLinesOnOrder = rs

End Function

Public Function GetChargesOnLine(iSerialID As Long) As Recordset
Dim rs          As Recordset

Set rs = New Recordset

rs.Open "qkcgChargeLine " & iSerialID, Cn, adOpenForwardOnly, adLockReadOnly


Set GetChargesOnLine = rs

End Function


Public Function GetLineCharges(iSerialID As Long) As Recordset
Dim rs          As Recordset
Dim cmd         As Command

Set cmd = New Command
With cmd
    Set .ActiveConnection = Cn
    .CommandText = "qkogLineCharges"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("iSerialID", adInteger, adParamInput, 2, iSerialID)
End With
Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
End With
Set GetLineCharges = rs
End Function
Public Function NextIncrement(ByVal pvarOrderType As Variant) As String
Dim rs              As Recordset
Dim m_lngInterval    As Long
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
    m_lngInterval = rs!iIncrement + 1&
    .Update "iIncrement", m_lngInterval
    lstr = rs!vchPrefix & Format$(m_lngInterval, "000000")
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
    Set mRsOEReps = New Recordset
    With mRsOEReps
        .CursorLocation = adUseClient
        .Open "qkogOrderReps", _
            Cn, adOpenForwardOnly, adLockReadOnly
        Set .ActiveConnection = Nothing
    End With
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
    If Not m_LineCommand Is Nothing Then Set m_LineCommand = Nothing
    If Not m_Chargecommand Is Nothing Then Set m_Chargecommand = Nothing
    If Not mCn Is Nothing Then
        Set mCn = Nothing
    End If
    If Not m_colSettings Is Nothing Then Set m_colSettings = Nothing
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
Screen.MousePointer = vbHourglass

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
Screen.MousePointer = vbArrow
Exit Sub
errhandler:
    Screen.MousePointer = vbArrow
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

Screen.MousePointer = vbHourglass

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
            rs.AddNew Array("iInvoiceID", "iOrderID", "iSerialID", "iQtyShipped", "mLineTotal"), _
                    Array(llngInvID, pOrder.OrderID, .SerialID, .ShippedNow, (.ShippedNow * .UnitPrice))
            lcurInvSubTotal = lcurInvSubTotal + (.ShippedNow * .UnitPrice)
            .Shipped = .Shipped + .ShippedNow
            .ShippedNow = 0
        End If
        If .Shipped < .ordered Then lbolFinished = False
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
Screen.MousePointer = vbArrow
PostOrder = lstrDocNumber
Exit Function
errhandler:
    Screen.MousePointer = vbArrow
    HandleError "Order Posting Failed. " & Err.Description, False
End Function

Public Function GetRecentSignatures(plngCustomerID As Long) As Recordset
Dim lrs As Recordset

Set lrs = New Recordset
With lrs
    .CursorLocation = adUseClient
    .Open "gqcgSignatureList " & plngCustomerID, Cn, adOpenStatic, adLockReadOnly
    Set .ActiveConnection = Nothing
End With

Set GetRecentSignatures = lrs
End Function

Public Function GetPOBindingRecordSet(plngPONumber As Long) As Recordset
Dim lrs As Recordset

Set lrs = New Recordset
With lrs
    .CursorLocation = adUseClient
    .Open "Select * from PurchaseOrderDetail where iPurchaseOrderID = " & plngPONumber, Cn, adOpenKeyset, adLockBatchOptimistic
    Set .ActiveConnection = Nothing
End With

Set GetPOBindingRecordSet = lrs
End Function

Public Sub ReconnectAndUpdate(in_RS As Recordset)
    Set in_RS.ActiveConnection = Cn
    in_RS.UpdateBatch
    Set in_RS.ActiveConnection = Nothing
End Sub


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
Dim rs As Recordset
Set rs = New Recordset
With rs
    .CursorLocation = adUseClient
    .Open "qkogBatchPrintingInfo " & plngBatchID, Cn, adOpenStatic, adLockReadOnly
    Set .ActiveConnection = Nothing
End With
Set GetBatchInfo = rs
Set rs = Nothing
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
Dim a_ReturnString As String

Set rs = Me.GetOrderHeader(plngOrderID)
    
If rs!iStatus = 0 Then
    Cn.Execute "qkoedOrder " & plngOrderID, , adExecuteNoRecords
    a_ReturnString = "Order status was ""NEW""." & vbCrLf & _
            "Entire Order was deleted."
Else
    Cn.Execute "qkoeCancelOrder " & plngOrderID, , adExecuteNoRecords
    a_ReturnString = "Order has been partially posted." & vbCrLf & _
            "Remaining Portion has been canceled."
End If
rs.Close
Set rs = Nothing

DeleteOrder = a_ReturnString
End Function

Public Function getInvoicebyNumber(in_InvoiceNumber As String) As String
    Dim rs As Recordset
    Set rs = New Recordset
    rs.Open "qkoegInvoiceByNumber '" & in_InvoiceNumber & "'", Cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        getInvoicebyNumber = ""
    Else
        getInvoicebyNumber = rs(0)
    End If
    rs.Close
    Set rs = Nothing
End Function
Public Function getInvoicebyOrder(in_OrderNumber As String) As String
Dim rs As Recordset
    Set rs = New Recordset
    rs.Open "qkoegInvoiceByOrderNumber '" & in_OrderNumber & "'", Cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        getInvoicebyOrder = ""
    Else
        getInvoicebyOrder = rs(0)
    End If
    rs.Close
    Set rs = Nothing
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
        Optional ByVal in_OrderId As Long = 0, _
        Optional ByVal in_ShipmentID As Long = 0)
Dim rs As Recordset

If Not (in_OrderId = 0) Then
    out_ReportSQL = "{vOrderHeader.iOrderID} = " & in_OrderId & " and " & _
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
Dim a_ReturnVal() As OrderableUnit
Dim a_Counter As Integer
Dim a_Unit As OrderableUnit

rs.CursorLocation = adUseClient
rs.Open "Select * from VendorSlabDetail where iVendorSlabTypeID = " & in_VendorSlabTypeID, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

a_Counter = 1
ReDim a_ReturnVal(1 To rs.RecordCount)
While Not rs.EOF
    Set a_Unit = New OrderableUnit
    a_Unit.PFC = rs!cPrice
    a_Unit.Length = rs!iLength * 12
    Set a_ReturnVal(a_Counter) = a_Unit
    a_Counter = a_Counter + 1
    rs.MoveNext
Wend

rs.Close

GetVendorOrderableUnits = a_ReturnVal
End Function

Public Sub InsertPODetailLines(in_Items As RawUnitItems, in_GroupByCode As String, in_VendorCode As String, in_Description As String, in_PO As Long)
    Dim a_LineItem As RawUnitItem
    Dim a_ChildItem As RawUnitItem
    Dim a_RS As Recordset
    Dim a_FieldNames As Variant
    Dim a_CurrentPODetailID As Long
    
    a_FieldNames = Array("iPurchaseOrderID", "vchGroupByCode", "vchVendorItemCode", "vchItemDescription", "vchItemSize", "iQuantity", "tiIsSlab")
    
    
    On Error GoTo errhandler:
    'Cn.BeginTrans
        Set a_RS = New Recordset
        a_RS.CursorLocation = adUseServer
        a_RS.Open "PurchaseOrderDetail", Cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
    For Each a_LineItem In in_Items
        a_RS.AddNew a_FieldNames, Array(in_PO, in_GroupByCode, in_VendorCode, in_Description, (a_LineItem.OrderableUnit.Length / 12) & "'", 1, True)
        a_CurrentPODetailID = a_RS!iPurchaseOrderLineID
        AddDetailLineMapping a_LineItem.SerialID, a_CurrentPODetailID
        If Not a_LineItem.Child Is Nothing Then
            Set a_ChildItem = a_LineItem.Child
            AddDetailLineMapping a_ChildItem.SerialID, a_CurrentPODetailID
            While Not a_ChildItem.Child Is Nothing
                Set a_ChildItem = a_ChildItem.Child
                AddDetailLineMapping a_ChildItem.SerialID, a_CurrentPODetailID
            Wend
        End If
    Next
    
    
    'Cn.CommitTrans
    
errhandler:
    'Cn.RollbackTrans
End Sub

Public Sub InsertPODetailLine(in_PO As Long, in_GroupByCode As String, in_VendorCode As String, in_Description As String, in_Length As String, in_Quantity As Long)
    Cn.Execute "Insert Into PurchaseOrderDetail (iPurchaseOrderID, vchGroupByCode, vchVendorItemCode, vchItemDescription, vchItemSize, iQuantity, tiIsSlab ) " & _
            "Values ( " & in_PO & ", '" & in_GroupByCode & "', '" & in_VendorCode & "', '" & _
            in_Description & "', '" & in_Length & "', " & in_Quantity & ", False )"
            
   
End Sub

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


Private Sub AddDetailLineMapping(in_SerialID, in_PurchaseOrderLineID)
    Cn.Execute "Insert Into PurchaseOrderDetailMapping (iSerialID, iPurchaseOrderLineID) Select " & in_SerialID & " As Expr1," & in_PurchaseOrderLineID & " as Expr2"
    Cn.Execute "Update OrderLine set tiOnPo = true where iSerialID = " & in_SerialID
End Sub
