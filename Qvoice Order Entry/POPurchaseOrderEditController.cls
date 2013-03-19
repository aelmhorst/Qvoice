VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "POPurchaseOrderEditController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPOController
Private mBindingData As Recordset
Private mDeletedItems As New Collection
Private mPOEditorView As frmPOEdit
Private mPONumber As String
Private mPOID As Long
Private mIs_Closing As Boolean


Public Sub POPurchaseOrderEditController(in_PONumber As String, in_POID As Long)
    mPONumber = in_PONumber
    mPOID = in_POID
    Set mBindingData = DataCenter.GetPOBindingRecordSet(in_POID)
    If mBindingData.EOF Then
        DataCenter.InsertPODetailLine in_POID, "", "", "", 0, 1, 1, 0@, 0
        Set mBindingData = DataCenter.GetPOBindingRecordSet(in_POID)
    End If
    Set mPOEditorView = New frmPOEdit
    'mPOEditorView.Show
    mPOEditorView.Init Me, False
End Sub

Private Function IPOController_CanCancel() As Boolean
    If mBindingData.RecordCount = 0 Then
        IPOController_CanCancel = True
        Exit Function
    ElseIf mIs_Closing Or mBindingData.EditMode = adEditNone Then
        IPOController_CanCancel = True
        Exit Function
    End If
    
    If MsgBox("Save Changes?", vbYesNo, App.Title) = vbYes Then
        IPOController_DoSaveAction
    End If
    IPOController_CanCancel = True
    
End Function


Private Sub IPOController_CreatePurchaseOrder()

End Sub

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
End Sub

Private Sub IPOController_DoSaveAction()
    POEditing.SavePOOrderLineItems mBindingData, mDeletedItems
End Sub

Private Property Get IPOController_EnableCreatePurchaseOrder() As Boolean
    IPOController_EnableCreatePurchaseOrder = False
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

Private Property Get IPOController_WindowCaption() As String
    IPOController_WindowCaption = "Edit Purchase Order " & mPONumber
End Property

Private Property Get IPOController_DataSource() As ADODB.Recordset
    Set IPOController_DataSource = mBindingData
End Property


Private Sub PrintPO()
    POEditing.PrintPurchaseOrder mPOID, False, True
End Sub

Private Sub IPOController_DoFinalAction()
    Dim aResults As VbMsgBoxResult
    
    IPOController_DoSaveAction
    
    If mPOID > 0 And mBindingData.RecordCount > 0 Then
        aResults = MsgBox("This PO has been saved. Would you like to print?", vbYesNoCancel, App.Title)
        If aResults = vbCancel Then
            Exit Sub
        ElseIf aResults = vbYes Then
            PrintPO
        End If
    End If
    Dispose
End Sub


Private Property Get IPOController_FinishButtonCaption() As String
    IPOController_FinishButtonCaption = "&Save Changes"
End Property



