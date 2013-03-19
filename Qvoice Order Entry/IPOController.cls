VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPOController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get DataSource() As Recordset
End Property
Public Property Get WindowCaption() As String
End Property
Public Property Get FinishButtonCaption() As String
End Property
Public Property Get PurchaseOrderID() As Long
End Property
Public Property Get EnableCreatePurchaseOrder() As Boolean
End Property
Public Sub DoFinalAction()
End Sub
Public Sub DoSaveAction()
End Sub
Public Function CanCancel() As Boolean
End Function
Public Sub DoCancelAction()
End Sub
Public Sub UndoChanges()
End Sub
Public Property Get PrintingEnabled() As Boolean
End Property
Public Sub DeleteCurrentRow()
End Sub
Public Sub CreatePurchaseOrder()
End Sub

