VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOEdit 
   Caption         =   "Po Editor"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10905
   Begin VB.Frame frmBottom 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   10815
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Finish"
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
         Left            =   8520
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Undo My Changes"
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
         Left            =   8520
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.ListBox lstDetailItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   10575
      End
      Begin VB.CheckBox chkDetails 
         Caption         =   "Show Details For Each Line"
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3015
      End
      Begin VB.Frame frmPrintArea 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   3600
         TabIndex        =   2
         Top             =   960
         Width           =   4695
         Begin VB.CommandButton cmdPrint 
            Caption         =   "P&review"
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
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Print Purchase Order"
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
            Index           =   0
            Left            =   1920
            TabIndex        =   6
            Top             =   0
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
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
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Print PO Cutting  Report"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Slab Labels"
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
            Index           =   2
            Left            =   1920
            TabIndex        =   3
            Top             =   720
            Width           =   2775
         End
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPOEdit.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Slab Order"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "iQuantity"
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "flSize"
         Caption         =   "Size"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "vchVendorItemCode"
         Caption         =   "Item Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "vchItemDescription"
         Caption         =   "Item Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "vchGroupByCode"
         Caption         =   "Color Group"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   1
         BeginProperty Column00 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc mData 
      Height          =   330
      Left            =   3360
      Top             =   2640
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "QKORDER"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrintMain 
         Caption         =   "&Print"
         Begin VB.Menu mnuPrint 
            Caption         =   "Purchase Order"
            Index           =   0
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuPrint 
            Caption         =   "PO Cutting Report"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPreVMain 
         Caption         =   "Pre&view"
         Begin VB.Menu mnuPreView 
            Caption         =   "Purchase Order"
            Index           =   0
         End
         Begin VB.Menu mnuPreView 
            Caption         =   "PO Cutting Report"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuFont 
         Caption         =   "F&ont Size"
         Begin VB.Menu mnuFontSize 
            Caption         =   "8 Point"
            Index           =   0
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "10 Point"
            Index           =   1
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "12 Point"
            Index           =   2
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "14 Point"
            Index           =   3
         End
         Begin VB.Menu mnuFontSize 
            Caption         =   "16 Point"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuMapPOLine 
         Caption         =   "Map PO Line to Order"
      End
      Begin VB.Menu mnuCreateFromStock 
         Caption         =   "Create Purchase Order Now"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPOEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum POEditMode
    IndividualOrder
    GeneralPO
End Enum
Private mController As IPOController
    
Public Sub Init(in_Controller As IPOController, in_DetailViewDefault As Boolean)
    Set mController = in_Controller
    SetupControls
    Me.frmPrintArea.Visible = mController.PrintingEnabled
    Me.chkDetails.Value = IIf(in_DetailViewDefault, vbChecked, vbUnchecked)
    Me.mnuCreateFromStock.Enabled = mController.EnableCreatePurchaseOrder
    chkDetails_Click
End Sub

Public Sub ClearDataBind()
    Set mData.Recordset = Nothing
End Sub

Public Sub DoDatabind()
    Set mData.Recordset = mController.DataSource
End Sub

Private Sub SetupControls()
    cmdNext.Caption = mController.FinishButtonCaption
    Caption = mController.WindowCaption
    SetPrintingEnabled (mController.PurchaseOrderID > 0)
    DoDatabind
End Sub


Private Sub SetPrintingEnabled(in_Enabled As Boolean)
    cmdPrint(0).Enabled = in_Enabled
    cmdPrint(1).Enabled = in_Enabled
End Sub

Private Sub chkDetails_Click()
    Me.lstDetailItems.Visible = chkDetails.Value = vbChecked
End Sub

Private Sub cmdNext_Click()
    mController.DoFinalAction
End Sub

Private Sub cmdCancel_Click()
    mController.UndoChanges
End Sub

Private Sub cmdPrint_Click(Index As Integer)
    If mController.DataSource.EditMode <> adEditNone Then
        Dim lResult As VbMsgBoxResult
        lResult = MsgBox("Save Changes?", vbYesNoCancel, App.Title)
        If lResult = vbCancel Then Exit Sub
        If lResult = vbYes Then mController.DoSaveAction
    End If
    Dim aShowPreview As Boolean
    aShowPreview = (Index = 1)
    
    If Me.optReport(0).Value Then
        POEditing.PrintPurchaseOrder mController.PurchaseOrderID, aShowPreview, True
    ElseIf Me.optReport(1).Value Then
        POEditing.PrintPurchaseOrderDetail mController.PurchaseOrderID, aShowPreview
    Else
        POEditing.PrintPurchaseOrderSlabLabels mController.PurchaseOrderID, aShowPreview
    End If
End Sub




Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    mController.DeleteCurrentRow
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    mData.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " ASC"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = Not mController.CanCancel
   
End Sub



Private Sub Form_Resize()
    Me.DataGrid1.Move 0, 0, ScaleWidth, ScaleHeight - frmBottom.Height
    frmBottom.Move 0, ScaleHeight - frmBottom.Height, ScaleWidth, frmBottom.Height
End Sub




Private Sub lstDetailItems_DblClick()
    If lstDetailItems.ListIndex > -1 Then
        Globals.MainDocumentWindow.OpenOrder lstDetailItems.ItemData(lstDetailItems.ListIndex)
    End If
End Sub



Private Sub mData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim aDescription As String
    If Not lstDetailItems.Visible Then Exit Sub
    If (adReason = adRsnMove) And pRecordset.AbsolutePosition > 0 Then
        lstDetailItems.Clear
        Dim aID As Long
        aID = pRecordset!iPurchaseOrderLineId
        If aID > 0 Then
            Dim rs As Recordset
            Set rs = DataCenter.GetOeLinesForPOLine(pRecordset!iPurchaseOrderLineId)
            Do Until rs.EOF
                aDescription = rs!vchInfo & ": Line " & rs!iLineNumber & " (" & rs!iOrdered & " ) - " & rs!txtLineDesc
'                If (Not IsNull(rs!dcLengthUsed)) And rs!dcLengthUsed > 0 Then
'                    aDescription = aDescription & " " & rs!dcLengthUsed & " inches needed."
'                End If
                lstDetailItems.AddItem aDescription
                lstDetailItems.ItemData(lstDetailItems.NewIndex) = rs!iOrderId
                rs.MoveNext
            Loop
            rs.Close
        End If
    End If
End Sub

Private Sub mData_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adReason = adRsnAddNew Then
        pRecordset.Fields.Item(Constants.PURCHASE_ORDER_ID).Value = mController.PurchaseOrderID
    End If
End Sub

Private Sub mData_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    'To do
End Sub

Private Sub mnuCreateFromStock_Click()
    mController.CreatePurchaseOrder
End Sub

Private Sub mnuFontSize_Click(Index As Integer)
    Dim aFontSizes As Variant
    aFontSizes = Array(8, 10, 12, 14, 18)
    Me.DataGrid1.Font.Size = CInt(aFontSizes(Index))
End Sub



Private Sub mnuMapPOLine_Click()
    Dim aOrderID As Long, aPurchaseOrderLineID As Long, aSerialID As Long
    Dim aFound As Boolean
    Dim aInputString As String
    
    aPurchaseOrderLineID = mData.Recordset!iPurchaseOrderLineId
    
    If aPurchaseOrderLineID = 0 Then
        MsgBox "This PO Must be saved and then re-opened before mapping this line to an order.", vbExclamation, App.Title
        Exit Sub
    End If
    
    DataCenter.ShowOpen 1, ForEditing
    aOrderID = DataCenter.SelectedJobID
    DataCenter.SelectedJobID = 0
    
    If aOrderID = 0 Then Exit Sub
    
    While Not aFound
        aInputString = GetLineNumberFromUser
        If Len(aInputString) = 0 Then Exit Sub
        If IsNumeric(aInputString) Then
            aSerialID = DataCenter.GetSerialIDForLineNumber(aOrderID, CInt(aInputString))
            If aSerialID > 0 Then
                DataCenter.AddDetailLineMapping aSerialID, aPurchaseOrderLineID, 1
                MsgBox "Line Detail Mapping Added", vbOKOnly, App.Title
            Else
                MsgBox "Line not found. Please Try Again.", vbOKOnly, App.Title
            End If
            
            
            Exit Sub
        End If
    Wend
End Sub

Private Function GetLineNumberFromUser() As String
    GetLineNumberFromUser = InputBox("Please Enter the line number from this order to associate with this PO Item", App.Title)
End Function
