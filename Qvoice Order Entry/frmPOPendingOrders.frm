VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOPendingOrders 
   Caption         =   "Slab Order Creation"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10800
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgPoReport 
      Bindings        =   "frmPOPendingOrders.frx":0000
      Height          =   4815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Slab Order Report"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "vchOrderNumber"
         Caption         =   "Order Number"
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
         DataField       =   "TotalFootage"
         Caption         =   "Feet"
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
         DataField       =   "vchCustomerName"
         Caption         =   "Customer"
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
         DataField       =   "dtRequestDate"
         Caption         =   "Request Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "vchPONumber"
         Caption         =   "Po Number"
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
      BeginProperty Column05 
         DataField       =   "vchJobName"
         Caption         =   "JobName"
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
      BeginProperty Column06 
         DataField       =   "vchGroupByCode"
         Caption         =   "Laminate"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdGoToOrder 
      Caption         =   "&Go To Order"
      Height          =   495
      Left            =   3080
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteRow 
      Caption         =   "&Remove Order"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreateOrder 
      Caption         =   "Create Slab Order"
      Height          =   495
      Left            =   6040
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc dbPOReport 
      Height          =   495
      Left            =   1440
      Top             =   1680
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
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
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   5520
      Width           =   10695
   End
End
Attribute VB_Name = "frmPOPendingOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDataSource As PendingPORequest
Private mDialogResultObj As SharedDataObject
Private mUser1 As String
Private mUser2 As String

Private Sub cmdCancel_Click()
    mDialogResultObj.DialogResultData = vbCancel
End Sub

Private Sub cmdCancelChanges_Click()
    dbPOReport.Recordset.CancelBatch
    Me.dgPoReport.ReBind
End Sub

Private Sub cmdCreateOrder_Click()
On Error GoTo errhandler:

    Screen.MousePointer = MousePointerConstants.vbHourglass
    
    mDialogResultObj.DialogResultData = vbOK
            
    Dim aOrderIDs As Variant
    aOrderIDs = dbPOReport.Recordset.GetRows(Start:=adBookmarkFirst, Fields:=Array(Constants.ORDER_HEADER_ORDER_ID))
    
    POEditing.GeneratePOFromStockOrderItems aOrderIDs, mDataSource.PODate, mUser1, mUser2
    
    Screen.MousePointer = MousePointerConstants.vbDefault
    Unload Me
    
Exit Sub
errhandler:
    Screen.MousePointer = MousePointerConstants.vbDefault
    MsgBox Err.Description, vbCritical, App.Title
End Sub


Private Sub cmdDeleteRow_Click()
    If mDataSource.Data.RecordCount > 0 Then
        mDataSource.TotalFootage = mDataSource.TotalFootage - dbPOReport.Recordset!TotalFootage
        dbPOReport.Recordset.Delete
        If dbPOReport.Recordset.RecordCount = 0 Then
            Unload Me
        Else
            ShowDetails
        End If
    End If
End Sub

Private Sub cmdGoToOrder_Click()
    If mDataSource.Data.RecordCount > 0 Then
        Globals.MainDocumentWindow.OpenOrder (dbPOReport.Recordset!iOrderId)
    End If
End Sub

Public Sub Init(inDataSource As PendingPORequest, _
    inDialogResult As SharedDataObject, _
    inUser1 As String, _
    inUser2 As String)
    Set mDialogResultObj = inDialogResult
    Set mDataSource = inDataSource
    
    mUser1 = inUser1
    mUser2 = inUser2
    inDataSource.Data.MoveFirst
    
    dbPOReport.Enabled = False
    Set dbPOReport.Recordset = inDataSource.Data
    dbPOReport.Enabled = True
    ShowDetails
    
    Me.Refresh
End Sub


Private Sub ShowDetails()
    lblStatus.Caption = mDataSource.TotalFootage & " Feet for Orders up to " & FormatDateTime(mDataSource.MaxDate, vbLongDate)
End Sub

Private Sub dgPoReport_HeadClick(ByVal ColIndex As Integer)
    dbPOReport.Recordset.Sort = dgPoReport.Columns(ColIndex).DataField & " ASC"
End Sub


Private Sub Form_Resize()
    dgPoReport.Move 0# * ScaleWidth, 0# * ScaleHeight, 1# * ScaleWidth, 0.81 * ScaleHeight
    cmdGoToOrder.Move 0.28 * ScaleWidth, 0.83 * ScaleHeight, 0.18 * ScaleWidth, 0.08 * ScaleHeight
    cmdDeleteRow.Move 0.01 * ScaleWidth, 0.83 * ScaleHeight, 0.18 * ScaleWidth, 0.08 * ScaleHeight
    cmdCancel.Move 0.8 * ScaleWidth, 0.83 * ScaleHeight, 0.18 * ScaleWidth, 0.08 * ScaleHeight
    cmdCreateOrder.Move 0.54 * ScaleWidth, 0.83 * ScaleHeight, 0.18 * ScaleWidth, 0.08 * ScaleHeight
    'dbPOReport.Move 0.16 * ScaleWidth, 0.28 * ScaleHeight, 0.47 * ScaleWidth, 0.08 * ScaleHeight
    lblStatus.Move 0.01 * ScaleWidth, 0.93 * ScaleHeight, 0.98 * ScaleWidth, 0.06 * ScaleHeight
End Sub


