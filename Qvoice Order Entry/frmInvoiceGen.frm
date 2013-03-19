VERSION 5.00
Begin VB.Form frmInvoiceGen 
   Caption         =   " Invoice Generation"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "frmInvoiceGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmStep1 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Next >>"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Step 1: Invoice Generation . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You are about to generate invoices for posted shipments. Press NEXT >> to continue."
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   0
         Left            =   120
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame frmStep1 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin VB.OptionButton optReport 
         Caption         =   "Print a single copy only of each Invoice"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CheckBox chkSummary 
         Caption         =   "Print Invoice Summary Report"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Print invoices as per customer preferences"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Generation is complete. Click the PRINT button to print the generated Invoices."
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   1
         Left            =   120
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Step 2: Invoice Printing . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.Frame frmStep1 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdBack 
         Caption         =   "<< &Back"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish Batch"
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInvoiceGen.frx":0442
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         FillColor       =   &H80000005&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   2
         Left            =   120
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "You are about to generate invoices for posted shipments. Press NEXT to continue. "
         Height          =   615
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Step 3: Mark Batch as complete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   5400
   End
End
Attribute VB_Name = "frmInvoiceGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements iListener

Private Type Batch
    Started As Boolean
    ID As Long
    Count As Long
End Type

Private CurrentBatch As Batch


Private Sub cmdBack_Click()
    frmStep1(1).ZOrder
End Sub

Private Sub cmdFinish_Click()
    Screen.MousePointer = MousePointerConstants.vbHourglass
    cmdFinish.Enabled = False
    cmdBack.Enabled = False
    Status "Finishing Invoice Batch . . ."
    DataCenter.FinishInvoiceBatch CurrentBatch.ID
    
    Status ""
    Me.Visible = False
    cmdBack.Enabled = True
    cmdFinish.Enabled = True
    CurrentBatch.Started = False
    frmStep1(0).ZOrder
    Screen.MousePointer = MousePointerConstants.vbDefault
    MsgBox "Invoice Generation Complete.", vbOKOnly Or vbInformation, App.Title
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Screen.MousePointer = MousePointerConstants.vbHourglass
    cmdGenerate.Enabled = False
    Status "Generating invoices . . ."
    DataCenter.GenerateInvoiceBatch CurrentBatch.ID, CurrentBatch.Count
    cmdGenerate.Enabled = True
    Screen.MousePointer = MousePointerConstants.vbDefault
    If CurrentBatch.Count = 0 Then
        CurrentBatch.Started = False
        MsgBox "No invoices to process", vbOKOnly Or vbInformation, App.Title
        Unload Me
        Exit Sub
    Else
        CurrentBatch.Started = True
        Caption = Caption & " - batch " & CurrentBatch.ID
    End If
    Status "Invoice generation complete. " & CurrentBatch.Count & " invoices generated."
    frmStep1(1).ZOrder
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = MousePointerConstants.vbHourglass
    cmdPrint.Enabled = False
    
    
    'Print the Invoice summary report
    If chkSummary.Value = vbChecked Then
        Status "Printing Invoice Summary . . ."
        ReportPrinter.PrintGenericDocument "invoice_summary", _
            "{InvoiceHeader.iInvoiceBatchID} = " & CurrentBatch.ID, _
            False, False
        Status "Invoice Summary Complete . . ."
    End If
    
    'Print the invoice batch
    Status "Printing Invoice Batch . . ."
    ReportPrinter.PrintInvoiceBatch CurrentBatch.ID, optReport(1).Value, Me
    Status "Invoice Printing Complete."
    
    frmStep1(2).ZOrder
    cmdPrint.Enabled = True
    Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub Form_Load()
    frmStep1(0).ZOrder
End Sub


Private Sub Status(txt As String)
    lblStatus = txt
    lblStatus.Refresh
    Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ltresult As VbMsgBoxResult
    If CurrentBatch.Started Then
        ltresult = MsgBox("You have begun the invoice generation process." & vbCrLf & vbCrLf & _
            "If you close this window, these invoices will have to be finished manually." & vbCrLf & vbCrLf & _
            "Do you want to continue with the invoice generation process?", vbYesNo, App.Title & " - Batch " & CurrentBatch.ID)
        If ltresult = vbNo Then
            ltresult = MsgBox("Terminate Invoice Generation?", vbYesNo, App.Title & " - Batch " & CurrentBatch.ID)
            If ltresult = vbYes Then
                MsgBox "Please notify your system administrator that batch number " & _
                    CurrentBatch.ID & " will have to be manually finished.", vbOKOnly Or vbExclamation, App.Title
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmStep1(0).ZOrder
End Sub

Private Sub iListener_Receive(pstrText As String)
    Status pstrText
End Sub
