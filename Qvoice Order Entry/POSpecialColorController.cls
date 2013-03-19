VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPOCreationController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPOController
Implements iListener

Private mDeletedItems As New Collection
Private mBindingData As Recordset
Private mOrderItemsList As dbPOSlabList
Private mPOEditorView As frmPOEdit
Attribute mPOEditorView.VB_VarHelpID = -1
Private mPONumber As String
Private mPOID As Long
Private mOrderID As Long
Private mUser1 As String
Private mUser2 As String
Private mIs_Closing As Boolean
Private Const STOCK_PO_NUM As String = "Slab Order"


Private Sub iListener_Receive(pstrText As String)

End Sub

Private Function IPOController_CanCancel() As Boolean
    If mBindingData.RecordCount = 0 Then
        IPOController_CanCancel = True
    ElseIf mIs_Closing Or mBindingData.EditMode = adEditNone Then
        IPOController_CanCancel = True
        Exit Function
    End If
    If Complete Then
        If MsgBox("Save Changes?", vbYesNo, App.Title) = vbYes Then
            IPOController_DoSaveAction
        End If
        IPOController_CanCancel = True
    Else
        Dim aResults As VbMsgBoxResult
        aResults = MsgBox("Not all PO's have been created or saved. Are you sure you want to exit?", vbYesNo, App.Title)
        IPOController_CanCancel = (aResults = vbNo)
    End If
End Function




Private Sub IPOController_DeleteCurrentRow()
    mDeletedItems.Add mBindingData!iPurchaseOrderLineId.Value
End Sub

Private Sub IPOController_DoCancelAction()
    Dispose
End Sub

Private Sub Dispose()
    mIs_Closing = True
    Unload mPOEditorView
    Set mPOEditorView = Nothing
    Set mBindingData = Nothing
    Set mOrderItemsList = Nothing
    'POEditing.HandleFinishedPO Me, mOrderID
End Sub

Private Sub IPOController_DoSaveAction()
    POEditing.SavePOOrderLineItems mBindingData, mDeletedItems
End Sub

Private Property Get IPOController_EnableCreatePurchaseOrder() As Boolean
     IPOController_EnableCreatePurchaseOrder = (mPONumber = STOCK_PO_NUM)
End Property

Private Property Get IPOController_PrintingEnabled() As Boolean
    IPOController_PrintingEnabled = True
End Property

Private Property Get IPOController_PurchaseOrderID() As Long
    IPOController_PurchaseOrderID = mPOID
End Property

Private Sub IPOController_UndoChanges()
    While mDeletedItems.Count > 0
        mDeletedItems.Remove 1
    Wend
    mBindingData.CancelBatch
End Sub

Public Sub OrderPOCreationController(in_OrderID As Long, inUser1 As String, inUser2 As String)
    mUser1 = inUser1
    mUser2 = inUser2
    mOrderID = in_OrderID
    
    Set mOrderItemsList = DataCenter.GetPOItemsOnOrder(in_OrderID)
    If Not mOrderItemsList.EOF Then
        If MsgBox("Create Purchase Orders?", vbYesNo, App.Title) = vbYes Then
            Set mPOEditorView = New frmPOEdit
            mPOEditorView.Show
            CreateNextPO
        End If
    Else
        Set mOrderItemsList = Nothing
    End If
End Sub

Private Property Get IPOController_WindowCaption() As String
    IPOController_WindowCaption = "Edit Purchase Order " & mPONumber
End Property

Private Property Get IPOController_DataSource() As ADODB.Recordset
    Set IPOController_DataSource = mBindingData
End Property

Private Property Get Complete()
    Complete = mOrderItemsList.EOF
End Property


Private Sub PrintPO(inUpdateStatus As Boolean)
    POEditing.PrintPurchaseOrder mPOID, False, inUpdateStatus
End Sub

Private Sub IPOController_DoFinalAction()
    IPOController_DoSaveAction

    '// Ask About Printing special color POs
    Dim aResults As VbMsgBoxResult
    If mBindingData.RecordCount > 0 And mPONumber <> STOCK_PO_NUM Then
        aResults = MsgBox("This PO has been saved. Would you like to print?", vbYesNoCancel, App.Title)
        If aResults = vbCancel Then
            Exit Sub
        ElseIf aResults = vbYes Then
            PrintPO True
        End If
    End If
    
    '// Go on to the next PO
    If Not Complete Then
        CreateNextPO
    Else
        Dispose
    End If
End Sub


Private Property Get IPOController_FinishButtonCaption() As String
    IPOController_FinishButtonCaption = "&Finish"
End Property


Private Sub CreateNextPO()
    If mOrderItemsList.BrandCode = STOCK_BRAND Then
        mPONumber = STOCK_PO_NUM
        mPOID = 0 - mOrderID
    Else
        DataCenter.CreatePO Now, 24, mUser1, mUser2, mPONumber, mPOID
    End If
    
    '// Order the caps for this color
    POEditing.OrderCapsForPO mOrderID, mOrderItemsList.BrandCode, mPOID
    POEditing.CombineLineItemsToPO mPOID, mOrderItemsList, POAlgorithmAction.CreatePOLineItemsSingleBrand, Me
    Set mBindingData = DataCenter.GetPOBindingRecordSet(mPOID)
    mPOEditorView.Init Me, True
End Sub

Private Sub IPOController_CreatePurchaseOrder()

    DataCenter.CreatePO Now, 24, mUser1, mUser2, mPONumber, mPOID
    '//Move the PO Items to the new PO
    
    DataCenter.MovePODetailLines 0 - mOrderID, mPOID
    
    Set mBindingData = DataCenter.GetPOBindingRecordSet(mPOID)
    mPOEditorView.Init Me, True

End Sub
