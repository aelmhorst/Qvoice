VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLookupInvoices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Look for an Invoice"
   ClientHeight    =   6195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tbLookupType 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8493
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Number"
      TabPicture(0)   =   "frmLookupInvoices.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "optSearchType(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSearch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optSearchType(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Customer"
      TabPicture(1)   =   "frmLookupInvoices.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFilter"
      Tab(1).Control(1)=   "cmbTimeFrame"
      Tab(1).Control(2)=   "ucCust"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Date"
      TabPicture(2)   =   "frmLookupInvoices.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ucDateInvoiced"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtFilter 
         Height          =   375
         Left            =   -74520
         TabIndex        =   13
         Top             =   4200
         Width           =   4935
      End
      Begin qkorder.ucDate ucDateInvoiced 
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "&OrderNumber"
         Height          =   375
         Index           =   1
         Left            =   -74400
         TabIndex        =   10
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   -74460
         TabIndex        =   8
         Top             =   3240
         Width           =   4335
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "In&voice Number"
         Height          =   375
         Index           =   0
         Left            =   -74400
         TabIndex        =   6
         Top             =   2280
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.ComboBox cmbTimeFrame 
         Height          =   315
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3360
         Width           =   5055
      End
      Begin qkorder.ucCustomer ucCust 
         Height          =   2775
         Left            =   -74520
         TabIndex        =   3
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
      End
      Begin VB.Label Label5 
         Caption         =   "You may enter part of a PO number or job name here, if you know it."
         Height          =   375
         Left            =   -74520
         TabIndex        =   14
         Top             =   3840
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Enter the desired date in the field below and click ""Show Invoices"" to display all Invoices created on that date."
         Height          =   615
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Enter the Invoice or Order number below and click ""Show Invoices to find the specified invoice."
         Height          =   615
         Left            =   -74520
         TabIndex        =   9
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Search by . . . ."
         Height          =   375
         Left            =   -74520
         TabIndex        =   7
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Select a customer and time frame to view Invoices."
         Height          =   255
         Left            =   -74640
         TabIndex        =   4
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Show &Invoices"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "frmLookupInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub Form_Initialize()
    Me.tbLookupType.Tab = 1
End Sub

Private Sub Form_Load()
Dim lvar
    With cmbTimeFrame
        For Each lvar In Array(30, 60, 90, 120, 180, 360)
            .AddItem "From the last " & lvar & " days"
            .ItemData(.NewIndex) = CLng(lvar)
        Next
        .ListIndex = 0
    End With
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub


Private Sub OKButton_Click()

OKButton.SetFocus
DoEvents
If tbLookupType.Tab = 1 Then
    lookupByCustomer
ElseIf tbLookupType.Tab = 0 Then
    lookupByNumber
Else
    lookupByDate
End If
End Sub
Private Sub lookupByDate()
Dim lstrSQL As String
Dim lstrDate As String

lstrDate = Format$(ucDateInvoiced.dtDate, "MM/DD/YYYY")
lstrSQL = "{InvoiceHeader.dtShipDate} >= #" & _
        lstrDate & " 00:00:0# And {InvoiceHeader.dtShipDate} <= #" & _
        lstrDate & " 23:59:0#"

ReportPrinter.PrintGenericDocument "invoice", lstrSQL, False, True, _
        "Invoices for " & Format$(ucDateInvoiced.dtDate, "MM/DD/YYYY"), Globals.MainDocumentWindow.hwnd, True
Me.Hide

End Sub

Private Sub lookupByCustomer()
    Dim ltresult As VbMsgBoxResult
    Dim llngTimeFrame As Long
    Dim lstrSQL As String
    Dim lstrPrompt As String
    Dim lResults() As String
    If ucCust.Customer.UniqueID = 0 Then
        MsgBox "Please Select a customer", vbOKOnly, App.Title
        Exit Sub
    End If
    
    llngTimeFrame = cmbTimeFrame.ItemData(cmbTimeFrame.ListIndex)
    
    
    lstrPrompt = "View all Invoices for " & _
                ucCust.Customer.AddressInfo.Name & _
                " for the last " & llngTimeFrame & " days?"
       
    If Len(txtFilter.Text) > 0 Then
        lResults = DataCenter.getInvoiceListByFilters(ucCust.Customer.UniqueID, DateAdd("d", -llngTimeFrame, Now()), txtFilter.Text)
        If lResults(0) = "-1" Then
            MsgBox "No Invoices found matching criteria", vbOKOnly, App.Title
            Exit Sub
        End If
        lstrSQL = "{InvoiceHeader.iInvoiceID} = " & Join(lResults, " OR {InvoiceHeader.iInvoiceID} = ")
        lstrPrompt = lstrPrompt & vbCrLf & "(Filter the results by searching " & _
            "for [ " & txtFilter.Text & " ]" & vbCrLf & _
            "( " & (UBound(lResults) + 1) & " items found. )"
    Else
        lstrSQL = "{Customer.iCustomerID} = " & ucCust.Customer.UniqueID & _
                   " AND {InvoiceHeader.dtShipDate} > #" & _
                   Format$(DateAdd("d", -llngTimeFrame, Now()), "MM/DD/YYYY") & "#"
    End If
    
    ltresult = MsgBox(lstrPrompt, vbYesNo, App.Title)
        
    If ltresult = vbYes Then
        ReportPrinter.PrintGenericDocument "invoice", lstrSQL, False, True, _
                "Invoices for " & ucCust.Customer.AddressInfo.Name, Globals.MainDocumentWindow.hwnd, True
        txtFilter.Text = ""
        Me.Hide
    End If

End Sub

Private Sub lookupByNumber()

    Dim lSearchStr As String
    Dim lInvoiceNumber As String
    
    If optSearchType(0).Value Then
        lSearchStr = CleanDocumentName(txtSearch.Text, "INV")
        lInvoiceNumber = DataCenter.getInvoicebyNumber(lSearchStr)
    Else
        lSearchStr = CleanDocumentName(txtSearch.Text, "ORD")
        lInvoiceNumber = DataCenter.getInvoicebyOrder(lSearchStr)
    End If
    
    txtSearch.Text = lSearchStr
    
    If Len(lInvoiceNumber) = 0 Then
        MsgBox "No invoice found matching '" & lSearchStr & "'"
    Else
        ReportPrinter.PrintGenericDocument "invoice", _
            "{InvoiceHeader.vchInvoiceNumber} = '" & lInvoiceNumber & "'", _
            False, True, "Invoice " & lInvoiceNumber, Globals.MainDocumentWindow.hwnd, True
        Me.Hide
    End If
End Sub




