VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public UniqueID         As Long
Public Alert            As String
Private mPopUpAlert       As Boolean
Public PriceListID      As Long
Public AddressInfo      As New AddressInfo
'Public Code             As String
Public Tax              As Currency
Public Discount         As Currency
Public DeliveryArea     As String
Public DeliveryDay      As String
Public PrefCartoned     As Boolean
Public prefSplined      As Boolean
Public PrefPadded       As Boolean
Public PrefGroupedOrders As Boolean
Private mstrPricingHandler    As String

Public Property Get Multiplier() As Currency
    Multiplier = 1@ - Discount
End Property

Public Property Get PopUpAlert()
    PopUpAlert = mPopUpAlert
End Property

Public Property Get HasAlert()
    HasAlert = Len(Alert) > 0
End Property


Public Sub Init(ByVal in_CustomerID As Long)
Dim rs As Recordset

On Error GoTo errhandler:
    If in_CustomerID <> UniqueID And in_CustomerID > 0 Then
        Set rs = DataCenter.GetCustomer(in_CustomerID)
        If rs.EOF Then
            rs.Close
            Set rs = Nothing
            MsgBox "Customer is not properly set up in the system.", vbOKOnly Or vbCritical, App.Title
            UniqueID = 0
            'Code = ""
            Alert = ""
            mPopUpAlert = False
            With AddressInfo
                .Name = ""
                .Address1 = ""
                .Address2 = ""
                .City = ""
                .State = ""
                .Zip = ""
                .Phone = ""
            End With
            
            Tax = 0#
            Discount = 0
            DeliveryArea = ""
            DeliveryDay = ""
            PrefCartoned = False
            prefSplined = False
            PrefPadded = False
            PrefGroupedOrders = False
            PriceListID = 0
        Else
            'The customer Id has changed, load properties
            With rs
                UniqueID = in_CustomerID
                'Code = !vchCustomerCode
            
                AddressInfo.Name = !vchCustomerName
                AddressInfo.Address1 = !vchAddress1 & ""
                AddressInfo.Address2 = !vchAddress2 & ""
                AddressInfo.City = !vchCity & ""
                AddressInfo.State = !vchState & ""
                AddressInfo.Zip = !vchZip & ""
                AddressInfo.Phone = !vchPhoneNumber & ""
                Alert = !txtAlert & ""
                mPopUpAlert = HasAlert And !tiPopUpAlert
                Tax = !flTaxRate
                Discount = !flDiscount
                DeliveryArea = !vchDeliveryArea & ""
                DeliveryDay = !vchRouteDesc & ""
                PrefCartoned = !tiCartoned
                prefSplined = !tiSplined
                PrefPadded = !tiPadded
                PrefGroupedOrders = !tiGroupOrders
                PriceListID = !iPriceListID
                .Close
            End With
            Set rs = Nothing
        End If
   End If
Exit Sub
errhandler:
    HandleError "Error Loading Customer." & vbCrLf & _
               "Details: " & Err.Description, False
    Init 0&
End Sub
