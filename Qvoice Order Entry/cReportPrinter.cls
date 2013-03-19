VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReportPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mstrCurrentReport   As String
Private mApplication     As CRPEAuto.Application
Private mReport          As CRPEAuto.Report

Public Sub PrintDeliveryReceipt(inOrderId As Long, inOrderDescription As String, inPrintToPreviewWindow As Boolean, Optional inOwnerWindowHandle As Long)
    Dim aReportSQL As String
    aReportSQL = "{vOrderHeader.iOrderID}=" & inOrderId & _
        " AND {OrderLine.ToShip} > 0"
    If inPrintToPreviewWindow Then
        ReportPrinter.PrintGenericDocument "delivery_receipt", _
        aReportSQL, False, inPrintToPreviewWindow, _
        inOrderDescription, inOwnerWindowHandle
    Else
        ReportPrinter.PrintGenericDocument "delivery_receipt", _
        aReportSQL
    End If
End Sub


Public Sub PrintOrderConfirmation(inOrderId As Long, inOrderDescription As String, inPrintToPreviewWindow As Boolean, Optional inOwnerWindowHandle As Long)
    Dim aReportSQL As String
    aReportSQL = "{vOrderHeader.iOrderID}=" & inOrderId
    If inPrintToPreviewWindow Then
        ReportPrinter.PrintGenericDocument "quote", _
        aReportSQL, False, inPrintToPreviewWindow, _
        inOrderDescription, inOwnerWindowHandle
    Else
        ReportPrinter.PrintGenericDocument "quote", _
        aReportSQL
    End If
End Sub


Public Sub PrintGenericDocument(pstrDocKey As String, pstrSelectionFormula As String, Optional pbolPromptUser As Boolean = False, Optional pbolPreviewWindow As Boolean = False, _
                Optional ByVal pstrWindowTitle As String = "Preview Window", Optional OwnerWindow As Long = 0, Optional ShowGroupTree As Boolean = False, Optional inParameters As Variant)
    Dim lOptions As CRPEAuto.PrintWindowOptions
    On Error GoTo errhandler:
    
    Screen.MousePointer = MousePointerConstants.vbHourglass
    If mstrCurrentReport <> pstrDocKey Then
        OpenReport pstrDocKey
    End If
    With mReport
        If .HasSavedData Then .DiscardSavedData
        '// Set the selection formula
        If Len(pstrSelectionFormula) > 0 Then .RecordSelectionFormula = pstrSelectionFormula
        
        '// Set the report parameters, if specified
        .ParameterPromptingEnabled = False
        If Not IsMissing(inParameters) Then
            Dim aCounter As Integer
            Dim aField As ParameterFieldDefinition
            For aCounter = 1 To UBound(inParameters)
                Set aField = .ParameterFields.Item(aCounter)
                '//aField.ClearCurrentValueAndRange '// CRAXDRT version
                aField.SetCurrentValue inParameters(aCounter)
            Next
        End If
        
        
        '// Set the preview options
        If pbolPreviewWindow Then
            Set lOptions = .PrintWindowOptions
            With lOptions
                .HasExportButton = True
                .HasPrintButton = True
                .HasPrintSetupButton = True
                .HasRefreshButton = True
                .HasSearchButton = True
                .HasZoomControl = True
                .HasGroupTree = ShowGroupTree
                .CanDrillDown = True
            End With
            pstrWindowTitle = "Print Preview - " & pstrWindowTitle
            .Options.ZoomMode = crPageWidth
            If OwnerWindow > 0 Then
                .Preview pstrWindowTitle, , , , , , OwnerWindow
            Else
                .Preview pstrWindowTitle
            End If
        Else
            '// Print out
            .PrintOut pbolPromptUser
        End If
    End With
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    Exit Sub
errhandler:
        Screen.MousePointer = MousePointerConstants.vbDefault
        If Err.Number = 20545 Then
        'Request was cancelled by user
        Else
            HandleError "Error Printing Report:" & vbCrLf & _
               "Error Number " & Err.Number & vbCrLf & _
               Err.Description & " : " & Err.Source, False
        End If

End Sub

Private Function OpenReport(pstrReportKey As String) As Boolean
Dim lPrinter            As Printer
Dim lstrReportPath      As String
Dim lstrReportPrinter   As String

'First, open the report engine
If pstrReportKey <> mstrCurrentReport Then
    lstrReportPath = Settings.GetReportPath(pstrReportKey)
    lstrReportPrinter = Settings.GetPrinter(pstrReportKey)
    If InStr(lstrReportPath, ":") = 0 Then lstrReportPath = App.Path & lstrReportPath
    
    If mApplication Is Nothing Then
        Set mApplication = New CRPEAuto.Application
    End If
    If Not mReport Is Nothing Then Set mReport = Nothing
    Set mReport = mApplication.OpenReport(lstrReportPath)
        
    If Len(lstrReportPrinter) > 0 And mReport.PrinterInfo.PrinterName <> lstrReportPrinter Then
        For Each lPrinter In Printers
            If lPrinter.DeviceName = lstrReportPrinter Then
                mReport.SelectPrinter lPrinter.DriverName, lPrinter.DeviceName, lPrinter.Port
                Exit For
            End If
        Next
    End If
    mstrCurrentReport = pstrReportKey
End If
End Function

Public Sub PrintInvoiceBatch(plngBatchID As Long, pbolSingleCopyOnly As Boolean, pListener As iListener)
Dim rs              As Recordset
Dim mlngCustomerID  As Long
Dim lintCopies      As Long

If mstrCurrentReport <> "invoice" Then OpenReport "invoice"

If pbolSingleCopyOnly Then
    With mReport
         If .HasSavedData Then .DiscardSavedData
        .RecordSelectionFormula = "{InvoiceHeader.iInvoiceBatchID} = " & plngBatchID
        '.RecordSortFields.Add crAscendingOrder, "{Customer.iCustomerID}"
        '.RecordSortFields.Add crAscendingOrder, "{InvoiceHeader.vchInvoiceNumber}"
        pListener.Receive "Beginning Print Process . . ."
        .PrintOut False
        pListener.Receive "Invoice Batch Printed . . ."
    End With
Else
    'mReport.RecordSortFields.Add crAscendingOrder, "{InvoiceHeader.vchInvoiceNumber}"
    Set rs = DataCenter.GetBatchInfo(plngBatchID)
    Do Until rs.EOF
        If mlngCustomerID <> rs!iCustomerID Then
                pListener.Receive "Printing Invoices for " & rs!vchCustomerName & " . . ."
                mlngCustomerID = rs!iCustomerID
                If IsNull(rs!iNumInvCopies) Then
                    lintCopies = 1
                Else
                    lintCopies = CInt(rs!iNumInvCopies)
                End If
                If lintCopies < 1 Then lintCopies = 1
            With mReport
                 If .HasSavedData Then .DiscardSavedData
                .RecordSelectionFormula = "{InvoiceHeader.iInvoiceBatchID} = " & plngBatchID & " AND {Customer.iCustomerID} = " & mlngCustomerID
                .PrintOut False, lintCopies
            End With
        End If
        rs.MoveNext
    Loop
    pListener.Receive "Invoice Batch Printed . . ."
End If
End Sub
