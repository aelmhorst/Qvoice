Attribute VB_Name = "Constants"
Option Explicit

'// Database Fields

    '// Order Header
    Public Const ORDER_HEADER_ORDER_ID = "iOrderID"
    Public Const ORDER_HEADER_ORDER_RUSHFLAG = "tiRush"
    Public Const ORDER_HEADER_POSTATUSFLAG = "tiPOStatusID"

    '// Purchase Order
    Public Const PURCHASE_ORDER_ID = "iPurchaseOrderID"
    Public Const PURCHASE_ORDER_NUMBER = "vchPONumber"
    Public Const PURCHASE_ORDER_STATUS = "iPOStatus"
    Public Const PURCHASE_ORDER_DATE_REQUESTED = "dtRequested"
    Public Const PURCHASE_ORDER_DATE_ORDERED = "dtDateOrdered"
    Public Const PURCHASE_ORDER_DATE_RECEIVED = "dtReceived"
    Public Const PURCHASE_ORDER_USER1 = "vchUser1"
    Public Const PURCHASE_ORDER_USER2 = "vchUser2"
    
    '// Purchase Order Detail
    Public Const PO_DETAIL_ORDERED = "iQuantity"
    Public Const PO_DETAIL_RECEIVED = "iQuantityReceived"
    Public Const PO_DETAIL_GROUPBY = "vchGroupByCode"
    Public Const PO_DETAIL_VENDORCODE = "vchVendorItemCode"
    Public Const PO_DETAIL_VENDORDESC = "vchItemDescription"
    Public Const PO_DETAIL_LINEID = "iPurchaseOrderLineID"
    Public Const PO_DETAIL_SIZE = "flSize"
    Public Const PO_DETAIL_ISSLAB = "tiIsSlab"
    
    '// Order Rep
    Public Const O_REP_EXTENSION = "vchExtension"
    
    Public Const REPORT_NAME_PO_DETAIL = "purchaseorderdetail"
    Public Const REPORT_NAME_PO = "purchaseorder"
    Public Const REPORT_NAME_PO_OPEN = "openpurchaseorders"
    Public Const REPORT_NAME_PO_LABELS = "purchaseorderslablabels"
    Public Const REPORT_NAME_ORDER_SLAB_STATUS = "orderslabstatus"
    
    
