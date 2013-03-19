VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QkOrderScanListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mFlicSession As FlicSession
Attribute mFlicSession.VB_VarHelpID = -1


Public Sub Init(inComm As MSComm, inFlicSettings As FlicSettings)
    Set mFlicSession = New FlicSession
    With mFlicSession
        .FlicSession_cctr inComm, inFlicSettings
        .Connect
        .Signal flicFlashAndBeep
    End With
End Sub


Private Sub Class_Terminate()
    If Not mFlicSession Is Nothing Then
        Set mFlicSession = Nothing
    End If
End Sub



Private Sub mFlicSession_OnScansReceived(inScanItems As ScanItems, outIsHandled As Boolean)
    Dim aFrmProgress As frmItemProcess
    Dim aShowProgress As Boolean
    Dim aScanData As String
    Dim aScan As Scan
    Dim aCurrentItem As Integer
    Dim aOk As Boolean
    Dim aPOLineID As String
    
    '// TODO: in the future as other types of scans are allowed besides PO Line ITems,
    '// We will need to have this code just do a branch to the appropriate class
    '// or classes, depending on whether we allow multiples types of scans at once
    
'try:
    On Error GoTo catch:
    If Screen.ActiveForm Is Globals.MainDocumentWindow Then
        aShowProgress = True
    ElseIf Screen.ActiveForm.MDIChild Then
        aShowProgress = True
    End If
    If aShowProgress Then
        Set aFrmProgress = New frmItemProcess
        aFrmProgress.Value = 0
        aFrmProgress.Max = inScanItems.Count * 1.5
        aFrmProgress.Show OwnerForm:=Globals.MainDocumentWindow
        aFrmProgress.ProgressCaption = "Processing Scan Items"
    End If
    For Each aScan In inScanItems
        aOk = False
        If Len(aScan.BarCode) > 1 Then
            If Left$(aScan.BarCode, 1) = "P" Then
                aPOLineID = Mid$(aScan.BarCode, 2)
                If IsNumeric(aPOLineID) Then
                    POEditing.ReceivePODetailItem (CLng(aPOLineID))
                    aOk = True
                End If
            End If
            If Not aOk Then
                MsgBox "Invalid Scan was discarded: " & aScan.BarCode, vbOKOnly, App.Title
            End If
        End If
        aCurrentItem = aCurrentItem + 1
        If aShowProgress Then aFrmProgress.Value = aCurrentItem
    Next
    
    '// Update the order header status for these items
    If aShowProgress Then aFrmProgress.Caption = "Updating Order Status"
    POEditing.UpdatePOStatusOnAllOpenOrders
    GoTo finally:
    
catch:
    HandleError "Error Saving Scan Data: " & aPOLineID & Err.Description, False
    Resume Next
finally:
    outIsHandled = True
    If aShowProgress Then Unload aFrmProgress
End Sub
