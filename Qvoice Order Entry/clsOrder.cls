VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Lines"
Attribute VB_Ext_KEY = "Member1" ,"Customer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

'Public properties and events for an Order (or Bid)
Public Event Error(ErrorMessage As String, ErrorLevel As Integer)
Public Event Message(EventMessage As String)
Public Event Saving(strMessage As String, dblPercentDone As Double)
Public Event CustomerChanged()

Public Enum eJobState
    ejsNew = 0
    ejsInProcess = 1
    ejsFinished = 2
End Enum

Public LastMessage              As String
Private mdteRequiredDate        As Date
Private mintCurrentLine         As Integer
Private mlngUserID              As Long
Private mintOriginalOrderType   As Integer
Private mstrOrdernumber         As String
Private mstrBidnumber           As String
Private mlngJobType             As Long
Private mState                  As eJobState
Private mcolProperties          As New Dictionary
Private mclsCustomer            As New Customer
Private mvarLines               As Lines



Public Property Get JobType() As Long
    JobType = mlngJobType
End Property
Public Property Let JobType(plngJobType As Long)
    mlngJobType = plngJobType
End Property

Public Property Get Editable() As Boolean
   Editable = (mState = ejsNew)
End Property
Public Property Get Customer() As Customer
    Set Customer = mclsCustomer
End Property

Public Property Get PriceListID() As Long
    PriceListID = mcolProperties("iPriceListID")
End Property

Public Property Let PriceListID(in_ID As Long)
    mcolProperties("iPriceListID") = in_ID
End Property

Public Property Get OrderID() As Long
    OrderID = mcolProperties(Constants.ORDER_HEADER_ORDER_ID)
End Property
Public Property Let OrderID(plngOrderID As Long)
    mcolProperties(Constants.ORDER_HEADER_ORDER_ID) = plngOrderID
End Property

Public Property Get Rush() As Boolean
    Rush = mcolProperties(Constants.ORDER_HEADER_ORDER_RUSHFLAG)
End Property

Public Property Let Rush(inRush As Boolean)
    mcolProperties(Constants.ORDER_HEADER_ORDER_RUSHFLAG) = inRush
End Property


Public Property Get PO() As String
    PO = mcolProperties(Constants.PURCHASE_ORDER_NUMBER)
End Property
Public Property Let PO(pstrPO As String)
    mcolProperties(Constants.PURCHASE_ORDER_NUMBER) = pstrPO
End Property
Public Property Get Trackingcode() As String
    Trackingcode = mcolProperties("vchTrackingCode")
End Property
Public Property Let Trackingcode(pstrTrackingcode As String)
    mcolProperties("vchTrackingCode") = pstrTrackingcode
End Property
Public Property Get Entrydate() As Date
    If IsEmpty(mcolProperties("dtEntryDate")) Then
        mcolProperties("dtEntryDate") = Now()
    End If
    Entrydate = mcolProperties("dtEntryDate")
End Property
Public Property Get ShipmentDate() As Date
    ShipmentDate = mcolProperties("dtShipDate")
End Property
Public Property Let Cartoned(pbolCartoned As Boolean)
   mcolProperties("tiCartoned") = pbolCartoned
End Property

Public Property Get Cartoned() As Boolean
    Cartoned = mcolProperties("tiCartoned")
End Property

Public Property Let Splined(pbolSplined As Boolean)
   mcolProperties("tiSplined") = pbolSplined
End Property

Public Property Get Splined() As Boolean
    Splined = mcolProperties("tiSplined")
End Property

Public Property Let Padded(pbolPadded As Boolean)
   mcolProperties("tiPadded") = pbolPadded
End Property

Public Property Get Padded() As Boolean
    Padded = mcolProperties("tiPadded")
End Property

Public Property Get Comment() As String
    Comment = mcolProperties("vchComment")
End Property
Public Property Get OrderNumber() As String
    OrderNumber = mcolProperties("vchOrderNumber")
End Property

Public Property Get Lines() As Lines
    If mvarLines Is Nothing Then
        Set mvarLines = New Lines
        mvarLines.Init Me
    End If
    Set Lines = mvarLines
End Property



Private Sub Class_Terminate()
    If Not mvarLines Is Nothing Then Set mvarLines = Nothing
End Sub

Public Property Get BidNumber() As String
BidNumber = mcolProperties("vchBidNumber")
End Property

Public Property Get ShipTo(in_LineNumber As Integer) As String
    ShipTo = mcolProperties("vchShipAddress" & CStr(in_LineNumber))
End Property

Public Property Let ShipTo(in_LineNumber As Integer, Value As String)
     mcolProperties("vchShipAddress" & CStr(in_LineNumber)) = Value
End Property

Public Function DeleteLine(LineNumber As Integer) As Boolean
Dim lbolOK As Boolean

If Not Editable Then
    lbolOK = False
    LastMessage = "Line cannot be removed!" & vbCrLf & _
        "This order has already been posted"
Else
    lbolOK = True
    Lines.Remove LineNumber
    If mintCurrentLine > Lines.Count Then mintCurrentLine = Lines.Count
    RaiseEvent Message("Line " & LineNumber & " removed")
End If

DeleteLine = lbolOK
End Function
Public Property Let RequiredDate(ReqDate As Date)
mcolProperties("dtRequestDate") = ReqDate
End Property
Public Property Get RequiredDate() As Date
    RequiredDate = mcolProperties("dtRequestDate")
End Property



'Public Property Get Errors() As Collection
''Dim lcolerrors As New Collection
''Dim lintCounter As Integer
''
''For lintCounter = 1 To mcolErrors.Count
''    lcolerrors.Add mcolErrors(lintCounter)
''Next
''Set Errors = lcolerrors
''Do Until mcolErrors.Count = 0
''    mcolErrors.Remove 1
''Loop
'End Property

Public Property Get Tax() As Currency
    Tax = mclsCustomer.Tax
End Property
Public Property Get Total() As Currency
    Total = CCur(CLng(List * mclsCustomer.Multiplier * 100) / 100)
End Property
Public Property Get List() As Currency
Dim lclsline As Line
Dim lcurLineTotals As Currency

For Each lclsline In Lines
    lcurLineTotals = lcurLineTotals + lclsline.List
Next
lcurLineTotals = lcurLineTotals + SpecialCharge.Price
List = lcurLineTotals

End Property
Public Property Let State(OrderState As eJobState)
mState = OrderState
End Property
Public Property Get State() As eJobState
State = mState
End Property

Public Property Get StateDescription() As String
    Select Case mState
        Case Is = ejsNew
            StateDescription = "New Order"
        Case Is = ejsInProcess
            StateDescription = "On Shipment"
        Case Is = ejsFinished
            StateDescription = "Closed"
      End Select
            
End Property
Public Property Get DescriptiveName() As String
Dim lstr As String
Dim lstrReturn As String

lstrReturn = Customer.AddressInfo.Name & " ( " & OrderNumber
lstr = GetProperty("vchJobName")
If Len(lstr) > 0 Then lstrReturn = lstrReturn & " \ " & lstr

If Len(PO) > 0 Then lstrReturn = lstrReturn & " \ " & PO

lstrReturn = lstrReturn & " )"
    DescriptiveName = lstrReturn
End Property
Public Sub SetProperty(pstrKey As String, pstrValue As String)
    With mcolProperties
        If Not .Exists(pstrKey) Then
            .Add pstrKey, pstrValue
        Else
            .Item(pstrKey) = pstrValue
        End If
    End With
End Sub
Public Function GetProperty(pstrKey As String) As Variant
    If mcolProperties.Exists(pstrKey) Then
        GetProperty = mcolProperties(pstrKey)
    Else
        GetProperty = ""
    End If
End Function

' Note: do not call this accessor if you are loading an order from the database
' Otherwise the pricelist will revert to the current one from the customer,
' which is not the desired behavior. The pricelist should only change if the customerid
' changes.
Public Property Let Customerid(in_CustomerID As Long)
    mclsCustomer.Init in_CustomerID
    Me.PriceListID = mclsCustomer.PriceListID
    RaiseEvent CustomerChanged
End Property

Public Property Let LineNumber(lintLine As Integer)
If lintLine <= Lines.Count Then
    mintCurrentLine = lintLine
Else
    RaiseEvent Error("Invalid Line Number", 1)
End If
End Property

Public Property Get LineNumber() As Integer
LineNumber = mintCurrentLine
End Property

Public Sub AddLine()
End Sub
Public Sub NewOrder(plngJobType As Long)
    mState = ejsNew
    mlngJobType = plngJobType
    With mcolProperties
        .Add "vchOrderNumber", ""
        .Add "tiCartoned", 0
        .Add "tiGroup", False
        .Add Constants.PURCHASE_ORDER_NUMBER, ""
        .Add "vchTrackingcode", ""
        .Add "dtPostedDate", Null
        .Add "iBatchID", 0
    End With
End Sub

Public Property Get SpecialCharge() As Charge
    Set SpecialCharge = QKGetSpecialCharge(mvarLines)
End Property

Public Sub Init(iOrderId As Long)
Dim rs          As Recordset
Dim fl          As Field

#If DBUG Then
   PerfLog "Loading Order", 0
#End If

Set rs = DataCenter.GetOrderHeader(iOrderId)

#If DBUG Then
   PerfLog "Order Header Retrieved", 1
#End If

For Each fl In rs.Fields
    If fl.Name = "iUserID" Then
        UserID = fl.Value
    Else
        mcolProperties.Add fl.Name, fl.Value
    End If
Next

Me.State = mcolProperties("iStatus")
mclsCustomer.Init mcolProperties("iCustomerID")

JobType = rs!iOrderType
mintOriginalOrderType = JobType

rs.Close
Set rs = Nothing

#If DBUG Then
   PerfLog "Loading Lines", 1
#End If
Set rs = DataCenter.GetLinesOnOrder(iOrderId)
Do Until rs.EOF
    #If DBUG Then
       PerfLog "Loading Line " & rs("iSerialID"), 2
    #End If
    Lines.Load rs("iSerialID"), rs
    rs.MoveNext
Loop
#If DBUG Then
   PerfLog "Completed Loading Lines", 1
#End If

rs.Close
Set rs = Nothing



#If DBUG Then
   PerfLog "Order Load Complete", 0
#End If

End Sub
Public Property Let UserID(plngUserID As Long)
    mlngUserID = plngUserID
End Property
Public Property Get UserID() As Long
    UserID = mlngUserID
End Property


Public Sub Save()
    Dim lcLine As qkorder.Line
    Dim lintCounter As Integer
    Dim lbolClone   As Boolean
    
    If JobType <> mintOriginalOrderType Then
        If CreateReference(mintOriginalOrderType) Then SetProperty "vchRefNumber", mcolProperties("vchOrderNumber")
        mcolProperties("vchOrderNumber") = ""
        If IsTemplate(mintOriginalOrderType) Then
            OrderID = 0
            lbolClone = True
        End If
        mintOriginalOrderType = JobType
    End If
    
    
    If Len(mcolProperties("vchOrderNumber")) = 0 Then
        mcolProperties("vchOrderNumber") = DataCenter.NextIncrement(JobType)
    End If
    
    DataCenter.SaveOrder Me
    
    For Each lcLine In Lines
        If lbolClone Then lcLine.SerialID = 0
        lintCounter = lintCounter + 1
        RaiseEvent Saving("Saving Order . . .", 100 * lintCounter \ Lines.Count)
        lcLine.Save OrderID
    Next
    RaiseEvent Saving(mstrOrdernumber, 100)
End Sub
