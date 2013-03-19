VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POSpecialLineItemController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPOController
Implements iListener

Private m_BindingData As Recordset
Private m_POEditorView As frmPOEdit
Private m_PONumber As String
Private m_POID As Long
Private m_OrderID As Long
Private m_Is_Closing As Boolean

Private Sub iListener_Receive(pstrText As String)

End Sub


Private Function IPOController_CanCancel() As Boolean
    If m_BindingData.RecordCount = 0 Then
        IPOController_CanCancel = True
    ElseIf m_Is_Closing Or m_BindingData.EditMode = adEditNone Then
        IPOController_CanCancel = True
        Exit Function
    End If
    
    Dim ltresult As VbMsgBoxResult
    ltresult = MsgBox("Save Changes?", vbYesNo, App.ProductName)
    If ltresult = vbCancel Then
        IPOController_CanCancel = False
    ElseIf ltresult = vbYes Then
        IPOController_DoSaveAction
        SavePOToBulkOrder
        IPOController_CanCancel = True
    Else
        DataCenter.DeletePOLinesByPOID m_POID
        IPOController_CanCancel = True
    End If
    
End Function

Private Sub IPOController_DoCancelAction()
    Dispose
End Sub

Private Sub Dispose()
    m_Is_Closing = True
    Unload m_POEditorView
    Set m_POEditorView = Nothing
    Set m_BindingData = Nothing
    'POEditing.HandleFinishedPO Me, m_OrderID
End Sub

Private Sub IPOController_DoSaveAction()
    DataCenter.ReconnectAndUpdate m_BindingData
End Sub

Private Property Get IPOController_PrintingEnabled() As Boolean
    IPOController_PrintingEnabled = False
End Property

Private Property Get IPOController_PurchaseOrderID() As Long
    IPOController_PurchaseOrderID = m_POID
End Property

Private Sub IPOController_UndoChanges()
    m_BindingData.CancelBatch
End Sub

Public Sub POSpecialLineItemController(in_OrderID As Long)
    m_OrderID = in_OrderID
    m_POID = 0 - in_OrderID
    Set m_BindingData = DataCenter.getSpecialVendorItemsOnOrder(in_OrderID)
    If Not m_BindingData.EOF Then
        If MsgBox("Create Purchase Order for Special Vendor Code Items?", vbYesNo, App.ProductName) = vbYes Then
                Dim a_RS As Recordset
                Set a_RS = DataCenter.getChargesForSpecialVendorItemsOnOrder(in_OrderID)
                Do While Not a_RS.EOF
                    DataCenter.InsertPODetailLine m_POID, a_RS!vchLaminateCode, a_RS!vchVendorCode, a_RS!vchVendorDesc, "", a_RS!Quantity, 1
                    a_RS.MoveNext
                Loop
                a_RS.Close
                Set a_RS = Nothing
            With m_BindingData
                Do Until .EOF
                    DataCenter.InsertPODetailLine m_POID, !vchLaminateCode, "", !vchSlabDesc, CStr(!flSlabLength), !iOrdered, 1, !iSerialID
                    .MoveNext
                Loop
                .Close
            End With

        End If
        Set m_BindingData = DataCenter.GetPOBindingRecordSet(m_POID)
        Set m_POEditorView = New frmPOEdit
        m_POEditorView.Show
        m_POEditorView.Init Me, True
    Else
        POEditing.HandleFinishedPO Me, m_OrderID
    End If
    
End Sub

Private Property Get IPOController_WindowCaption() As String
    IPOController_WindowCaption = "Edit Purchase Order " & m_PONumber
End Property

Private Property Get IPOController_DataSource() As ADODB.Recordset
    Set IPOController_DataSource = m_BindingData
End Property




Private Sub IPOController_DoFinalAction()
    DataCenter.ReconnectAndUpdate m_BindingData
    SavePOToBulkOrder
    Dispose
End Sub

Private Sub SavePOToBulkOrder()
    DataCenter.MovePODetailLines 0 - m_OrderID, -1
    MsgBox "Line Items have been added to next bulk order.", vbOKOnly, App.ProductName
End Sub

Private Property Get IPOController_FinishButtonCaption() As String
    IPOController_FinishButtonCaption = "&Finish"
End Property

