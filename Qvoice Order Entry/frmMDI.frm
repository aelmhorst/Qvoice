VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Qvoice Order Entry"
   ClientHeight    =   6195
   ClientLeft      =   1740
   ClientTop       =   945
   ClientWidth     =   8190
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSCommLib.MSComm mComPort 
      Left            =   7440
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuHyphen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Begin VB.Menu mnuOrderType 
            Caption         =   ""
            Index           =   1
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu mnudash9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuPrintDoc 
            Caption         =   "&Delivery Receipt"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuPrintQuote 
            Caption         =   "&Quote / Confirmation"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuLabels 
            Caption         =   "&Labels"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnudash10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Pre&view"
         Begin VB.Menu mnuPrintOrderStatus 
            Caption         =   "Order Status"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnudash8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "Recent Orders"
         Enabled         =   0   'False
         Begin VB.Menu mnuRecentOrders 
            Caption         =   "None"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuShipments 
         Caption         =   "&Shipments"
         Begin VB.Menu mnuGenerateBatch 
            Caption         =   "&Generate Shipment"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuReprintShipmentBatch 
            Caption         =   "&Reprint Shipment"
         End
         Begin VB.Menu mnuDeleteShipmentBatch 
            Caption         =   "Delete Shipment"
         End
      End
      Begin VB.Menu mnuPostMain 
         Caption         =   "Pos&ting"
         Begin VB.Menu mnuPost 
            Caption         =   ""
            Index           =   1
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuPostShipment 
            Caption         =   "Post Shipment"
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuPurchaseOrders 
         Caption         =   "&Purchase Orders"
         Begin VB.Menu mnuCreatePO 
            Caption         =   "&Edit And Search for POs"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuEditNext 
            Caption         =   "Edit Next &Bulk Order"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuOrdersNoPos 
            Caption         =   "&Find Orders W/out POs"
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu mnuReports 
         Caption         =   "&Reports"
         Begin VB.Menu mnuLaminateLookup 
            Caption         =   "&Laminate Lookup"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuChargeCodes 
            Caption         =   "&Charge Code Lookup"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuHyphen5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInvoiceLookup 
            Caption         =   "&Invoice Lookup"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuInvoiceSummary 
            Caption         =   "InvoiceSummary"
         End
      End
      Begin VB.Menu mnuhyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Begin VB.Menu mnuDeleteOrder 
            Caption         =   "Delete Order"
         End
      End
      Begin VB.Menu mnuHyphen11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintInvoices 
         Caption         =   "&Invoice Generation"
      End
      Begin VB.Menu mnuMisc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuAngleCalculator 
         Caption         =   "&Angle Calculator"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Op&tions"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuTile 
         Caption         =   "&Tile"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Qvoice"
      End
      Begin VB.Menu mnuQvoiceHelp 
         Caption         =   "&Qvoice Help"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'   frmMDI
'   Created By:     Andy Elmhorst
'   Date:           08/22/1998
'   Purpose:        The parent MDI form for the QK Order Entry Program
'***********************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const SW_SHOWNORMAL As Long = 1

Option Explicit

Private Sub MDIForm_Load()
    Dim lintState As Integer
    'Set up the caption and our Window Location
    With Me
        .Caption = App.Title
        .Left = GetSetting("qkorder", "Settings_" & .Name, "MainLeft", 1000)
        .Top = GetSetting("qkorder", "Settings_" & .Name, "MainTop", 1000)
        .Width = GetSetting("qkorder", "Settings_" & .Name, "MainWidth", 6500)
        .Height = GetSetting("qkorder", "Settings_" & .Name, "MainHeight", 6500)
    End With
End Sub


Public Sub PostInitialize()
    Dim lintMax     As Integer
    Dim lintCounter As Integer
    Dim lintPost    As Integer
    
    lintMax = UBound(OrderTypes)
    
    lintPost = 1
    
    mnuOrderType(1).Caption = OrderTypes(1).Caption
    mnuOrderType(1).Tag = OrderTypes(1).ID
    
    If OrderTypes(1).CanPost Then
        mnuPost(1).Caption = OrderTypes(1).Caption
        mnuPost(1).Tag = OrderTypes(1).ID
        lintPost = lintPost + 1
    End If
    
    For lintCounter = 2 To lintMax
        Load mnuOrderType(lintCounter)
        mnuOrderType(lintCounter).Caption = OrderTypes(lintCounter).Caption
        mnuOrderType(lintCounter).Tag = OrderTypes(lintCounter).ID
        If OrderTypes(lintCounter).CanPost Then
            If lintPost > 1 Then Load mnuPost(lintPost)
            With mnuPost(lintPost)
                .Caption = OrderTypes(lintCounter).Caption
                .Tag = OrderTypes(lintCounter).ID
            End With
            lintPost = lintPost + 1
        End If
    Next
    mnuAbout.Caption = "About " & App.Title
    
    ReloadSettings

End Sub

Private Sub ReloadSettings()
    Dim aFlicSettings As FlicSettings
    Set aFlicSettings = MainModule.GetFlicSettings(1)
    If Not aFlicSettings Is Nothing Then
        If aFlicSettings.Enabled Then
            On Error GoTo errhandler
            Set Globals.BarCodeListener = New QkOrderScanListener
            Globals.BarCodeListener.Init Me.mComPort, aFlicSettings
        Else
            Set Globals.BarCodeListener = Nothing
        End If
    End If
    Exit Sub
errhandler:
    MsgBox "Scanner could not be enabled. " & Err.Description
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Me.WindowState <> vbMinimized Then
    With Me
        SaveSetting "qkorder", "Settings_" & .Name, "LastState", CStr(Me.WindowState)
        SaveSetting "qkorder", "Settings_" & .Name, "MainLeft", Me.Left
        SaveSetting "qkorder", "Settings_" & .Name, "MainTop", Me.Top
        SaveSetting "qkorder", "Settings_" & .Name, "MainWidth", Me.Width
        SaveSetting "qkorder", "Settings_" & .Name, "MainHeight", Me.Height
    End With
End If


'Tell the data access object to close all connections
'DataCenter.ReleaseAll

End Sub

Private Sub mnuAbout_Click()
  frmSplash.EnableUnload = True
    frmSplash.Show vbModal, Me
  

End Sub

Private Sub mnuAngleCalculator_Click()
    frmAngleCalc.Show
End Sub

Private Sub mnuCalculator_Click()
Dim ReturnValue As Double
    ReturnValue = Shell("CALC.EXE", 1)  ' Run Calculator.
    AppActivate ReturnValue     ' Activate the Calculator.
    SendKeys "%vs", True
End Sub

Private Sub mnuCascade_Click()
    Arrange 0
End Sub

Private Sub mnuChargeCodes_Click()
ReportPrinter.PrintGenericDocument "chargecodes", "", True, True, "Charge Codes Lookup", Me.hwnd, True
End Sub

Private Sub mnuCopy_Click()
    Dim lcontrol As Control
    
    If Screen.ActiveControl Is Nothing Then Exit Sub
    
    Set lcontrol = Screen.ActiveControl
    If TypeOf lcontrol Is TextBox Then
        Clipboard.SetText lcontrol.Text
    ElseIf TypeOf lcontrol Is ComboBox Then
        Clipboard.SetText lcontrol.Text
    ElseIf TypeOf lcontrol Is MSFlexGrid Then
        Clipboard.SetText lcontrol.Clip
    End If
End Sub



Private Sub mnuCreatePO_Click()
    POEditing.CreateGeneralPO Me
End Sub

Private Sub mnuDeleteOrder_Click()
Dim aDeletionStatus As String
'Ask the data access object to show orders that can be opened
DataCenter.ShowOpen 1, ForEditing

If DataCenter.SelectedJobID > 0 Then
    If MsgBox("Delete Order " & DataCenter.SelectedJobInfo & "?", vbYesNo, App.Title) = vbYes Then
        aDeletionStatus = DataCenter.DeleteOrder(DataCenter.SelectedJobID)
        MsgBox DataCenter.SelectedJobInfo & " has been deleted with the following results." & vbCrLf & vbCrLf & _
        aDeletionStatus, vbOKOnly Or vbInformation, App.Title
    End If
    DataCenter.SelectedJobID = 0
End If
End Sub

Private Sub mnuDeleteShipmentBatch_Click()
Dim ltresult As VbMsgBoxResult
Dim llngBatchid As Long

llngBatchid = GetShipmentBatch
If llngBatchid > 0 Then
    ltresult = MsgBox("Are you sure you want to delete Shipment " & _
            llngBatchid & "?", vbYesNo, App.Title)
    If ltresult = vbYes Then
        DataCenter.DeleteShipment llngBatchid
        MsgBox "Shipment " & llngBatchid & " was deleted.", vbOKOnly Or vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuEditNext_Click()
    POEditing.EditPendingBulkPO
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuGenerateBatch_Click()
    frmGenerateBatch.Show vbModal, Me
End Sub

Private Sub mnuInvoiceLookup_Click()
    frmLookupInvoices.Show vbModal, Me
End Sub

Private Sub mnuInvoiceSummary_Click()
Dim lstrSelection As String
Dim ldtBeginDate As Date
Dim ldtEndDate As Date

If Not GetDate("Please Enter Begin Date.", ldtBeginDate) Then Exit Sub
If Not GetDate("Please Enter End Date.", ldtEndDate) Then Exit Sub
'{InvoiceHeader.dtShipDate} in Date (2000, 08, 12) to Date (2000, 08, 13)"
lstrSelection = "{InvoiceHeader.dtShipDate} in Date (" & _
               Year(ldtBeginDate) & ", " & Month(ldtBeginDate) & ", " & _
               Day(ldtBeginDate) & ") to Date (" & _
               Year(ldtEndDate) & ", " & Month(ldtEndDate) & ", " & _
               Day(ldtEndDate) & ")"

ReportPrinter.PrintGenericDocument "invoice_lookup", lstrSelection, _
                         False, True, "Invoices from " & _
                         ldtBeginDate & " to " & ldtEndDate, Me.hwnd, True

End Sub
Private Function GetDate(pstrCaption As String, pdtDate As Date) As Boolean
Dim lstrString As String
Dim lbolExit As VbTriState

lbolExit = vbFalse

Do Until lbolExit <> vbFalse
    lstrString = InputBox(pstrCaption, App.Title)
    If Len(lstrString) = 0 Then
        lbolExit = vbTrue
    Else
        If IsDate(lstrString) Then
            pdtDate = CDate(lstrString)
            lbolExit = vbUseDefault
        End If
    End If
Loop

GetDate = (lbolExit = vbUseDefault)

End Function



Private Sub mnuLabels_Click()
    If FrmMainIsActive Then
       ActiveForm.PrintLabels
    Else
        Dim aFrm As New frmPrintOrderDocuments
        aFrm.SetOrderInfo "", 0
        aFrm.EnableLabelPrint
        aFrm.Show vbModal, Me
    End If
End Sub

Private Function FrmMainIsActive() As Boolean
    FrmMainIsActive = False
    If Not ActiveForm Is Nothing Then
        If TypeOf ActiveForm Is frmMain Then FrmMainIsActive = True
    End If
End Function

Private Sub mnuLaminateLookup_Click()
    frmColorLookup.Show
End Sub

'***********************************************************************
'   Private Sub mnuNew_Click()
'   Created By      Andy Elmhorst
'   Purpose         Opens a new order window
'***********************************************************************
Private Sub mnuNew_Click(Index As Integer)
Dim lfrmmain As New frmMain

    'Show the form
    lfrmmain.Show
    
    'Activate the NewJob procedure in the form
    lfrmmain.NewJob CLng(Index)
    
    Set lfrmmain = Nothing
End Sub





Private Sub mnuOrdersNoPos_Click()
    
    '// Ask the data access object to show orders that have no PO
    DataCenter.ShowOpen CInt(1), ForCreatingPOs
        
    If DataCenter.SelectedJobID > 0 Then
        PrivateOpenOrder DataCenter.SelectedJobID
        DataCenter.SelectedJobID = 0
    End If
End Sub

Private Sub mnuOrderType_Click(Index As Integer)
    'Ask the data access object to show orders that can be opened
    DataCenter.ShowOpen CInt(mnuOrderType(Index).Tag), ForEditing
    
    If DataCenter.SelectedJobID > 0 Then
        PrivateOpenOrder DataCenter.SelectedJobID
    End If
End Sub

Public Sub OpenOrder(in_OrderID As Long)
    '// Determine if the order is already open
    Dim X As Integer
    Dim aFrmMain As frmMain
    For X = (Forms.Count - 1) To 0 Step -1
        If TypeOf Forms(X) Is frmMain Then
            Set aFrmMain = Forms(X)
            If aFrmMain.OrderID = in_OrderID Then
                aFrmMain.ZOrder
                Exit Sub
            End If
        End If
    Next X

    '// We Didn't Find it, open a new window
    Set aFrmMain = New frmMain
    aFrmMain.OpenOrder in_OrderID
    
End Sub


Private Sub PrivateOpenOrder(in_OrderID As Long)
    Dim lfrmmain As New frmMain
    lfrmmain.OpenOrder in_OrderID
    DataCenter.SelectedJobID = 0
End Sub

Private Sub mnuPaste_Click()
    Dim lcontrol As Control
    
    If Screen.ActiveControl Is Nothing Then Exit Sub
    
    Set lcontrol = Screen.ActiveControl
    
    If TypeOf lcontrol Is TextBox Then
        lcontrol.Text = Clipboard.GetText
    ElseIf TypeOf lcontrol Is ComboBox Then
        lcontrol.Text = Clipboard.GetText
    End If
End Sub

Private Sub mnuPost_Click(Index As Integer)
    Dim lOrder As Order
    
    'Ask the data access object to show orders that can be opened
    DataCenter.ShowOpen CInt(mnuPost(Index).Tag), ForPosting
    
    'Show the form
    If DataCenter.SelectedJobID > 0 Then
        Set lOrder = New Order
        lOrder.Init DataCenter.SelectedJobID
        frmPost.Init lOrder
        frmPost.Show vbModal, Me
        DataCenter.SelectedJobID = 0
    End If
End Sub



Private Sub mnuPostShipment_Click()
    frmPostShipment.Show vbModal, Me
End Sub



Private Sub mnuPrintDoc_Click()
    If FrmMainIsActive Then
        ActiveForm.PrintDeliveryReceipt
    End If
End Sub

Private Sub mnuPrintInvoices_Click()
    frmInvoiceGen.Show , Me
End Sub



Private Sub mnuPrintOrderStatus_Click()
    If FrmMainIsActive Then
        ActiveForm.PreviewOrderStatus
    End If
End Sub

Private Sub mnuPrintQuote_Click()
    If FrmMainIsActive Then
        ActiveForm.PrintOrderDocument
    End If
End Sub

Private Sub mnuQvoiceHelp_Click()
    Dim lstrHelp As String
    Dim hwndDesk As Long
    
    hwndDesk = GetDesktopWindow
    
    lstrHelp = App.Path & "\Help\index.htm"
    Call ShellExecute(hwndDesk, "Open", lstrHelp, "", 0&, SW_SHOWNORMAL)
End Sub

Private Sub mnuReprintShipmentBatch_Click()
Dim llngBatch
llngBatch = GetShipmentBatch

If llngBatch > 0 Then
    ReportPrinter.PrintGenericDocument "delivery_receipt", _
            "{vOrderHeader.iBatchID}=" & llngBatch, _
            False, True, "Preview Shipment Batch " & llngBatch, _
            Me.hwnd
End If
End Sub

Public Function GetShipmentBatch() As Long
    Dim ltrResult As VbMsgBoxResult
    Dim llngBatch As Long
    Dim lstrResult As String
    llngBatch = -1
    
    Do Until llngBatch > -1
        lstrResult = InputBox("Please Enter Batch ID", App.Title)
        If Len(lstrResult) = 0 Then
            llngBatch = 0
        ElseIf IsNumeric(lstrResult) Then
            llngBatch = CLng(lstrResult)
        End If
    Loop
    GetShipmentBatch = llngBatch
End Function

Private Sub mnuSave_Click()
    If FrmMainIsActive Then
        Dim frm As frmMain
        Set frm = ActiveForm
        frm.Save
    End If
End Sub

Private Sub mnuSettings_Click()
    frmOptions.Show vbModal, Me
    ReloadSettings
End Sub

Private Sub mnuTile_Click()
    Arrange 1
End Sub
