VERSION 5.00
Begin VB.Form frmPrintOrderDocuments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Documents for "
   ClientHeight    =   5310
   ClientLeft      =   2235
   ClientTop       =   6300
   ClientWidth     =   7365
   Icon            =   "frmPrintLabels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPrintDeliveryReceipts 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton optDelRecOutput 
         Caption         =   "Preview Window"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optDelRecOutput 
         Caption         =   "Print to Printer"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CheckBox chkPrintDeliveryReceipts 
         Caption         =   "Print Delivery Receipts For This Order"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame frmPrintOrderConfirmation 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   6975
      Begin VB.OptionButton optOrderConfirmationOutput 
         Caption         =   "Fax"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optOrderConfirmationOutput 
         Caption         =   "Email"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optOrderConfirmationOutput 
         Caption         =   "Preview Window"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optOrderConfirmationOutput 
         Caption         =   "Print to Printer"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CheckBox chkPrintOrderConfirmation 
         Caption         =   "Print Order Confirmation"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame frmLabels 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   6975
      Begin VB.Frame frmlabelprintoptions 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   6375
         Begin VB.OptionButton optLabelOutput 
            Caption         =   "Preview Window"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   18
            Top             =   0
            Width           =   2295
         End
         Begin VB.OptionButton optLabelOutput 
            Caption         =   "Print to Printer"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   0
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkPrintLabels 
         Caption         =   "Print Labels"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
      Begin VB.CheckBox chkNotPrinted 
         Caption         =   "Only print new, unprinted line items"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.TextBox txtShipmentNumber 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton OptSpecificShipment 
         Caption         =   "A specific shipment: "
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optCurrentOrder 
         Caption         =   "This Order"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   4335
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintOrderDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOrderID As Long
Private mOrderDescription As String

Public Sub EnableOrderConfirmationPrint()
    chkPrintOrderConfirmation.Value = vbChecked
End Sub

Public Sub EnableDeliveryReceiptPrint()
    chkPrintDeliveryReceipts.Value = vbChecked
End Sub

Public Sub EnableLabelPrint()
    chkPrintLabels.Value = vbChecked
End Sub

Public Sub SetOrderInfo(in_OrderInfo As String, in_OrderID As Integer)
    mOrderDescription = in_OrderInfo
    mOrderID = in_OrderID
    
    If in_OrderID = 0 Then
        Me.Caption = "Print Shipment Labels"
    
        '// Disable the order label printing options
        optCurrentOrder.Enabled = False
        OptSpecificShipment.Value = True
        
        '// Disable the Delivery Receipt and Order Confirmation
        frmPrintOrderConfirmation.Enabled = False
        frmPrintDeliveryReceipts.Enabled = False
    Else
        Me.Caption = "Print Documents for " & in_OrderInfo
    End If
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub




'Private Sub Form_Load()
'    OptSpecificShipment.Value = True
'    OptSpecificShipment_Click
'End Sub

Private Sub OKButton_Click()
    HandlePrintDeliveryReceipt
    HandlePrintOrderConfirmation
    If HandlePrintLabels Then
        Unload Me
    End If
    
End Sub

Private Sub HandlePrintDeliveryReceipt()
    If chkPrintDeliveryReceipts.Value = vbChecked Then
       ReportPrinter.PrintDeliveryReceipt mOrderID, "Delivery Receipt for " & mOrderDescription, optDelRecOutput(1).Value, Globals.MainDocumentWindow.hwnd
       chkPrintDeliveryReceipts.Value = vbUnchecked
    End If
End Sub

Private Sub HandlePrintOrderConfirmation()
    If chkPrintOrderConfirmation = vbChecked Then
        ReportPrinter.PrintOrderConfirmation mOrderID, "Order Confirmation for " & mOrderDescription, optOrderConfirmationOutput(1).Value, Globals.MainDocumentWindow.hwnd
        chkPrintOrderConfirmation.Value = vbUnchecked
    End If
End Sub

Private Function HandlePrintLabels() As Boolean
    If chkPrintLabels.Value Then
        If optCurrentOrder.Value Then
            HandlePrintLabels = PrintOrderLabels(optLabelOutput(1).Value)
        Else
            HandlePrintLabels = PrintShipmentLabels(optLabelOutput(1).Value)
        End If
    Else
        HandlePrintLabels = True
    End If
        
End Function

Private Sub optCurrentOrder_Click()
    EnableShipment False
End Sub

Private Sub optDelRecOutput_Click(Index As Integer)
    chkPrintDeliveryReceipts.Value = vbChecked
End Sub

Private Sub optLabelOutput_Click(Index As Integer)
    chkPrintLabels.Value = vbChecked
End Sub

Private Sub optOrderConfirmationOutput_Click(Index As Integer)
    chkPrintOrderConfirmation.Value = vbChecked
End Sub

'Private Sub optSpecificDeliver_Click()
'    EnableDelivery True
'    EnableShipment False
'End Sub

Private Sub OptSpecificShipment_Click()
    EnableShipment True
    'EnableDelivery False
End Sub

Private Sub EnableShipment(in_Flag As Boolean)
    txtShipmentNumber.Enabled = in_Flag
    txtShipmentNumber.BackColor = IIf(in_Flag, vbWhite, RGB(127, 127, 127))
End Sub

'Private Sub EnableDelivery(in_Flag As Boolean)
'    'cmbDelivery.Enabled = in_Flag
'    'cmbDelivery.BackColor = IIf(in_Flag, vbWhite, RGB(127, 127, 127))
'End Sub


Private Function PrintOrderLabels(inPreview As Boolean) As Boolean
    Dim lstrReportSQL As String
    Dim lstrDBSQL As String
    Dim llngRecords As Long
    
     DataCenter.GetLabelReportInfo llngRecords, lstrReportSQL, lstrDBSQL, (chkNotPrinted.Value = vbChecked), in_OrderID:=mOrderID
     
    If llngRecords > 0 Then
        ReportPrinter.PrintGenericDocument "labels", lstrReportSQL, False, inPreview, "Labels For" & mOrderDescription, Globals.MainDocumentWindow.hwnd
        DataCenter.UpdateLabels lstrDBSQL
        PrintOrderLabels = True
    Else
        MsgBox "No labels were found to print for this order. " & vbCrLf & _
            IIf((chkNotPrinted.Value = vbChecked), _
                "Try unchecking the ""Only print new, unprinted line items"" check box. ", _
                "Check the order details."), _
                vbOKOnly Or vbInformation, App.Title
        PrintOrderLabels = False
    End If
    
End Function

Private Function PrintShipmentLabels(inPreview As Boolean) As Boolean
    Dim lstrReportSQL As String
    Dim lstrDBSQL As String
    Dim llngBatchid As Long
    Dim llngRecords As Long
    
    If Not IsNumeric(txtShipmentNumber.Text) Then
        MsgBox "Invalid Shipment Number", vbExclamation, App.Title
        PrintShipmentLabels = False
        Exit Function
    End If
    llngBatchid = CLng(txtShipmentNumber.Text)

    DataCenter.GetLabelReportInfo llngRecords, lstrReportSQL, lstrDBSQL, (chkNotPrinted.Value = vbChecked), in_ShipmentID:=llngBatchid
      
    If llngRecords > 0 Then
        ReportPrinter.PrintGenericDocument "labels", lstrReportSQL, False, inPreview, "Labels for " & mOrderDescription, Globals.MainDocumentWindow.hwnd
        DataCenter.UpdateLabels lstrDBSQL
        PrintShipmentLabels = True
    Else
        MsgBox "No labels were found to print for this shipment. " & vbCrLf & _
            IIf((chkNotPrinted.Value = vbChecked), _
                "Try unchecking the ""Only print new, unprinted line items"" check box. ", _
                "Check the shipment number."), _
                vbOKOnly Or vbInformation, App.Title
        PrintShipmentLabels = False
    End If
    
End Function


