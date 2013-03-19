VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPOCreationDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create or Search for PO"
   ClientHeight    =   5820
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   9135
   Icon            =   "POCreationDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab TabPOs 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   9551
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Create a Purchase Order"
      TabPicture(0)   =   "POCreationDialog.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblOrderRepPrompt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OERep"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ucPODate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frameCreateNewPurchaseOrder"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Print Purchase Orders"
      TabPicture(1)   =   "POCreationDialog.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPOPrintPrompt"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ucBeginDate"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ucEndDate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdPrint"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optReportType(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "optReportType(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "optReportType(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "optReportType(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Edit Purchase Orders"
      TabPicture(2)   =   "POCreationDialog.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCheckinPOs"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtPurchaseOrder"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSearch"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblEditPrompt"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdCheckinPOs 
         Caption         =   "Check In Purchase Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70320
         TabIndex        =   27
         Top             =   3720
         Width           =   2655
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Print Purchase Order Labels"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   26
         Top             =   3840
         Width           =   4695
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Print Open Purchase Order Items Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   23
         Top             =   3360
         Width           =   4695
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Print Cutting Detail Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   22
         Top             =   2880
         Width           =   3375
      End
      Begin VB.OptionButton optReportType 
         Caption         =   "Print Standard Purchase Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   2400
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.TextBox txtPurchaseOrder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   14
         Top             =   1560
         Width           =   5775
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Edit Purchase Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70320
         TabIndex        =   13
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   12
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Frame Frame4 
         Caption         =   "Create a Slab Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -70320
         TabIndex        =   8
         Top             =   1980
         Width           =   4335
         Begin qkorder.ucDate ucMaxPODate 
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   1200
            Width           =   2895
            _ExtentX        =   4260
            _ExtentY        =   873
         End
         Begin VB.CommandButton cmdAnalyze 
            Caption         =   "Analyze"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   9
            Top             =   2160
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Enter a date and click ""Analyze"" to see the raw footage totals for all open orders up to that date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame frameCreateNewPurchaseOrder 
         Caption         =   "Create a blank Purchase Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         TabIndex        =   4
         Top             =   1980
         Width           =   4335
         Begin VB.TextBox txtVendorID 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   360
            TabIndex        =   6
            Text            =   "24"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CommandButton CmdCreateBlankPO 
            Caption         =   "Create a blank PO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label Label2 
            Caption         =   "For Vendor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   840
            Width           =   2055
         End
      End
      Begin qkorder.ucDate ucPODate 
         Height          =   375
         Left            =   -72360
         TabIndex        =   2
         Top             =   840
         Width           =   3495
         _ExtentX        =   4048
         _ExtentY        =   661
      End
      Begin qkorder.ucDate ucEndDate 
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin qkorder.ucDate ucBeginDate 
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin qkorder.ucOERep OERep 
         Height          =   375
         Left            =   -72360
         TabIndex        =   24
         Top             =   1440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
      End
      Begin VB.Label lblOrderRepPrompt 
         Caption         =   "Order Rep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblPOPrintPrompt 
         Caption         =   "Select the Begin Date, End Date and Type of Report. Click ""Preview"" to view the report."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Top             =   1080
         Width           =   7695
      End
      Begin VB.Label Label5 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Begin Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblEditPrompt 
         Caption         =   "Enter the Purchase Order Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73440
         TabIndex        =   15
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Purchase Order Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Label mStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   9135
   End
End
Attribute VB_Name = "frmPOCreationDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements iListener
Private mMsgBoxResult As New SharedDataObject

Public Property Let DialogResult(inResult As VbMsgBoxResult)
    mMsgBoxResult = inResult
End Property


Private Sub cmdAnalyze_Click()
Dim aRequest As PendingPORequest
Dim aSlabOrderForm As frmPOPendingOrders
Dim aTotalFootage As Integer

On Error GoTo errhandler

    mMsgBoxResult.DialogResultData = vbCancel
    Screen.MousePointer = MousePointerConstants.vbHourglass
    Set aRequest = DataCenter.GetPendingPOFootage(Me.ucMaxPODate.dtDate, Me.ucPODate.dtDate)
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    If aRequest.HasData Then
        Set aSlabOrderForm = New frmPOPendingOrders
        aSlabOrderForm.Init aRequest, mMsgBoxResult, OERep.SelectedRepName, OERep.SelectedRepExtension
        aSlabOrderForm.Show vbModal, Me
        If mMsgBoxResult.DialogResultData = vbOK Then
            Unload Me
        End If
        Set aSlabOrderForm = Nothing
    Else
        MsgBox "No Footage Found", vbOKOnly, App.Title
    End If
    
    Exit Sub
errhandler:
    Screen.MousePointer = MousePointerConstants.vbDefault
    Dim aError As String
    aError = Err.Description & "-" & Err.Source & "-" & Err.Number
    LogIt aError
    MsgBox aError, vbOKOnly, App.Title
    Unload Me
End Sub




Private Sub cmdCheckinPOs_Click()
'    Dim aFrm As Form
'    For Each aFrm In Forms
'        If aFrm Is frmPostPO Then
'            aFrm.ZOrder
'            Exit Sub
'        End If
'    Next
'
    Dim afrmPostPO As New frmPostPO
    afrmPostPO.Show vbModal, Me
End Sub

Private Sub CmdCreateBlankPO_Click()
    POEditing.GenerateBlankPO CLng(Me.txtVendorID.Text), Me.ucPODate.dtDate, OERep.SelectedRepName, OERep.SelectedRepName
End Sub



Private Sub cmdPostPO_Click()
    Dim lPONum As Long
    Me.txtPurchaseOrder.Text = CleanDocumentName(txtPurchaseOrder.Text, "PO")
    lPONum = DataCenter.getPOID(txtPurchaseOrder.Text)
    If lPONum > 0 Then
        Unload Me
        Dim aPOCreationDialog As New frmPostPO
        aPOCreationDialog.Show vbModal
        '//aPOCreationDialog.Init lPONum
    Else
        MsgBox "No Purchase Order Found", vbOKOnly Or vbInformation, App.Title
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim lPONum As Long
    Me.txtPurchaseOrder.Text = CleanDocumentName(txtPurchaseOrder.Text, "PO")
    lPONum = DataCenter.getPOID(txtPurchaseOrder.Text)
    If lPONum > 0 Then
        Dim aController As New POPurchaseOrderEditController
        aController.POPurchaseOrderEditController txtPurchaseOrder.Text, lPONum
        Unload Me
    Else
        MsgBox "No Purchase Order Found", vbOKOnly Or vbInformation, App.Title
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim aReportName As String
    Dim lstrSQL As String
    Dim lstrWindowCaption As String
    

    
    If Me.optReportType(0).Value Then
        lstrSQL = GetReportSQLForDateRange(Me.ucBeginDate.dtDate, Me.ucEndDate.dtDate)
        aReportName = Constants.REPORT_NAME_PO_DETAIL
        lstrWindowCaption = "View PurchaseOrders between " & Format$(ucBeginDate.dtDate, "MM/DD/YYYY") & _
        " and " & Format$(ucEndDate.dtDate, "MM/DD/YYYY")
    ElseIf Me.optReportType(1).Value Then
        lstrSQL = GetReportSQLForDateRange(Me.ucBeginDate.dtDate, Me.ucEndDate.dtDate)
        aReportName = Constants.REPORT_NAME_PO
        lstrWindowCaption = "View Cutting Report for PurchaseOrders between " & Format$(ucBeginDate.dtDate, "MM/DD/YYYY") & _
        " and " & Format$(ucEndDate.dtDate, "MM/DD/YYYY")
    ElseIf Me.optReportType(2).Value Then
        aReportName = Constants.REPORT_NAME_PO_OPEN
        lstrSQL = GetReportSQLForOpenPOs(Me.ucBeginDate.dtDate, Me.ucEndDate.dtDate)
        lstrWindowCaption = "View Open PurchaseOrder Items between " & Format$(ucBeginDate.dtDate, "MM/DD/YYYY") & _
        " and " & Format$(ucEndDate.dtDate, "MM/DD/YYYY")
    Else
        lstrSQL = GetReportSQLForDateRange(Me.ucBeginDate.dtDate, Me.ucEndDate.dtDate) & " AND {PurchaseOrderDetail.tiIsSlab} = TRUE"
        aReportName = Constants.REPORT_NAME_PO_LABELS
        lstrWindowCaption = "View Slab Labels for PurchaseOrders between " & Format$(ucBeginDate.dtDate, "MM/DD/YYYY") & _
        " and " & Format$(ucEndDate.dtDate, "MM/DD/YYYY")
    End If
    
    '// Unload the options window first
    Unload Me
    
    '// Print the report
    ReportPrinter.PrintGenericDocument aReportName, lstrSQL, False, True, _
    lstrWindowCaption, Globals.MainDocumentWindow.hwnd, True

    
End Sub

Private Function GetReportSQLForDateRange(inBeginDate As Date, inEndDate As Date) As String
    Dim lstrBeginDate As String
    Dim lstrEndDate As String
    
    lstrBeginDate = Format$(ucBeginDate.dtDate, "MM/DD/YYYY")
    lstrEndDate = Format$(ucEndDate.dtDate, "MM/DD/YYYY")
    
    GetReportSQLForDateRange = "{PurchaseOrder.dtDateOrdered} >= #" & _
       lstrBeginDate & " 00:00:00# And {PurchaseOrder.dtDateOrdered}  <= #" & _
       lstrEndDate & " 23:59:59#"
End Function

Private Function GetReportSQLForOpenPOs(inBeginDate As Date, inEndDate As Date) As String
    GetReportSQLForOpenPOs = GetReportSQLForDateRange(inBeginDate, inEndDate) & _
        " AND {PurchaseOrder.iPOStatus} < 3 AND " & _
        "{PurchaseOrderDetail.iQuantityReceived} < {PurchaseOrderDetail.iQuantity}"
End Function



Private Sub Form_Load()
    TabPOs.Tab = 0
End Sub

Private Sub iListener_Receive(pstrText As String)
    mStatus.Caption = pstrText
End Sub

Private Sub OKButton_Click()
    If Not IsNumeric(txtVendorID.Text) Then
        MsgBox "Please enter valid vendor ID", vbOKOnly Or vbInformation, App.Title
        Exit Sub
    End If
    If MsgBox("Generate New PO for Vendor " & txtVendorID.Text & "?", vbYesNo, App.Title) = vbNo Then Exit Sub
    POEditing.GenerateBlankPO CLng(txtVendorID.Text), Me.ucPODate.dtDate, OERep.SelectedRepName, OERep.SelectedRepName
    Unload Me
End Sub





Private Sub txtPurchaseOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
    End If
End Sub
