VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPostPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order posting"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check In All"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo Changes"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton cmdCheckinAllCaps 
      Caption         =   "Check In All Caps"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   60
      Width           =   1815
   End
   Begin VB.CheckBox chkOnlyShowNotReceivedItems 
      Caption         =   "Only show line items not yet received"
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.VScrollBar vscrollReceived 
      Height          =   495
      Left            =   3360
      Max             =   0
      Min             =   -1
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView lvItems 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   13996
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Purchase Order"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Request Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Last Received Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin qkorder.MyDataGrid mDataGrid 
      Height          =   7935
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   13996
   End
   Begin VB.Label lblCaption 
      Caption         =   "Open Purchase Orders"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPostPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mBoundColumns(6) As String
Private mBoundColumnHeaders(6) As String
Private mPONumber As Long
Private mCurrentPORecordset As Recordset
Private mbolCurrentHasChanges As Boolean
Private mbolAnyPOChanges    As Boolean
Private mbolSavedChanges As Boolean
Private mbolPOFinal As Boolean

Public Sub LoadPO(inPONumber As Long)
    mPONumber = inPONumber
    Set mCurrentPORecordset = DataCenter.GetPOBindingRecordSet(inPONumber)
    mbolPOFinal = True
    BindGrid
End Sub

Private Sub BindGrid()
    If Not mCurrentPORecordset Is Nothing Then
        If Me.chkOnlyShowNotReceivedItems.Value = vbChecked Then
            mCurrentPORecordset.Filter = "iRemaining > 0 "
        Else
            mCurrentPORecordset.Filter = adFilterNone
        End If
        Me.mDataGrid.DatabindSpecificColumns mCurrentPORecordset, mBoundColumns, mBoundColumnHeaders
    End If
End Sub

Private Sub chkOnlyShowNotReceivedItems_Click()
    BindGrid
End Sub

Private Sub cmdCheckAll_Click()
    CheckinAll False
End Sub

Private Sub cmdCheckinAllCaps_Click()
    CheckinAll True
End Sub

Private Sub CheckinAll(inCapsOnly As Boolean)
    Dim aReceivedField As Field
    Dim aOrderedField As Field
    With mCurrentPORecordset
        .MoveFirst
        While Not .EOF
            Set aReceivedField = .Fields(Constants.PO_DETAIL_RECEIVED)
            Set aOrderedField = .Fields(Constants.PO_DETAIL_ORDERED)
            If (Not inCapsOnly) Or (.Fields(Constants.PO_DETAIL_ISSLAB).Value = False) Then
                If aReceivedField.Value <> aOrderedField.Value Then
                    aReceivedField.Value = aOrderedField.Value
                    .Fields("iRemaining") = 0
                    mbolCurrentHasChanges = True
                    mbolAnyPOChanges = True
                End If
            End If
            .MoveNext
        Wend
    End With
    BindGrid
End Sub

Private Sub cmdUndo_Click()
    mCurrentPORecordset.CancelBatch
    BindGrid
End Sub

Private Sub Form_Load()
    '// Load the Purchase Orders that are not yet posted
    Dim rs As Recordset
    Set rs = DataCenter.GetPOsByStatus(POStatusOnPO)
    BindPOList rs
    
    Set rs = DataCenter.GetPOsByStatus(POStatusPartial)
    BindPOList rs
    

    mBoundColumns(0) = Constants.PO_DETAIL_RECEIVED
    mBoundColumns(1) = Constants.PO_DETAIL_ORDERED
    mBoundColumns(2) = Constants.PO_DETAIL_GROUPBY
    mBoundColumns(3) = Constants.PO_DETAIL_SIZE
    mBoundColumns(4) = Constants.PO_DETAIL_VENDORCODE
    mBoundColumns(5) = Constants.PO_DETAIL_VENDORDESC
    mBoundColumns(6) = Constants.PO_DETAIL_LINEID
    
    
    mBoundColumnHeaders(0) = "Received"
    mBoundColumnHeaders(1) = "Ordered"
    mBoundColumnHeaders(2) = "Color"
    mBoundColumnHeaders(3) = "Size"
    mBoundColumnHeaders(4) = "Code"
    mBoundColumnHeaders(5) = "Description"
    mBoundColumnHeaders(6) = "ID"
    
    If lvItems.ListItems.Count = 0 Then
        MsgBox "No Purchase Orders Found", vbOKOnly, App.Title
    Else
        lvItems.SelectedItem = lvItems.ListItems(1)
        BindToPo
    End If
    
End Sub


Private Sub BindPOList(inRS As Recordset)
    Dim aListItem As ListItem
    
    While Not inRS.EOF
        Set aListItem = lvItems.ListItems.Add(, , inRS.Fields.Item(Constants.PURCHASE_ORDER_NUMBER).Value)
        aListItem.Tag = inRS.Fields.Item(Constants.PURCHASE_ORDER_ID).Value
        aListItem.SubItems(1) = Format$(inRS.Fields.Item(Constants.PURCHASE_ORDER_DATE_REQUESTED).Value, "YYYY/MM/DD")
        aListItem.SubItems(2) = Format$(inRS.Fields.Item(Constants.PURCHASE_ORDER_DATE_ORDERED).Value, "YYYY/MM/DD")
        If IsNull(inRS.Fields.Item(Constants.PURCHASE_ORDER_DATE_RECEIVED)) Then
            aListItem.SubItems(3) = "N/A"
        Else
            aListItem.SubItems(3) = Format$(inRS.Fields.Item(Constants.PURCHASE_ORDER_DATE_REQUESTED).Value, "YYYY/MM/DD")
        End If
        '// If this PO has already been edited, mark it
        If inRS.Fields.Item(Constants.PURCHASE_ORDER_STATUS).Value = POStatusEnum.POStatusPartial Then
            aListItem.Text = aListItem.Text & "*"
        End If
        inRS.MoveNext
    Wend
End Sub

Private Sub BindToPo()
    Dim aItem As ListItem
    
    If CheckForSave Then
        If lvItems.ListItems.Count = 0 Then
            MsgBox "All POs have been checked in", vbOKOnly, App.Title
            Unload Me
        Else
            Set aItem = lvItems.SelectedItem
            If Not aItem Is Nothing Then
                mbolCurrentHasChanges = False
                vscrollReceived.Visible = False
                LoadPO CLng(aItem.Tag)
                Me.Caption = "Checkin Purchase Order " & aItem.Text & " - " & aItem.SubItems(1)
            End If
        End If
    End If
End Sub

Private Function CheckForSave() As Boolean
    CheckForSave = True
    If mbolCurrentHasChanges Or mbolPOFinal Then
        Dim ltresult As VbMsgBoxResult
        ltresult = MsgBox("Save Changes?", vbYesNoCancel, App.Title)
        Select Case ltresult
            Case vbCancel
                CheckForSave = False
                Exit Function
            Case vbYes
                SaveAndRefreshScreen
                mbolCurrentHasChanges = False
        End Select
    End If
End Function


Private Sub SaveAndRefreshScreen()

    Dim aPOState As POStatusEnum
    aPOState = SaveChangesToRecordset

    '// Update the UI list of POs to reflect the change
    Dim aItem As ListItem
    For Each aItem In lvItems.ListItems
        If aItem.Tag = mPONumber Then
            If aPOState = POStatusPartial Then
                If Not Right$(aItem.Text, 1) = "*" Then
                    aItem.Text = aItem.Text & "*"
                End If
            ElseIf aPOState = POStatusReceived Then
                lvItems.ListItems.Remove aItem.Index
            End If
            Exit For
        End If
    Next

End Sub

Private Sub DoPostCloseSaveChanges()
    If mbolAnyPOChanges Then
        POEditing.UpdatePOStatusOnAllOpenOrders
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CheckForSave Then
        Cancel = True
    Else
        Me.Hide
        DoPostCloseSaveChanges
    End If
End Sub



Private Sub lvItems_Click()
    If Not lvItems.SelectedItem Is Nothing Then
        BindToPo
    End If
End Sub

Private Sub mDataGrid_FormatRow(inGrid As MSFlexGridLib.MSFlexGrid, inRow As Integer, inRS As ADODB.Recordset)
    If inRS.Fields.Item(Constants.PO_DETAIL_ORDERED) - inRS.Fields.Item(Constants.PO_DETAIL_RECEIVED) = 0 Then
        Dim lCounter As Integer
        For lCounter = 0 To inGrid.Cols - 1
            inGrid.Col = lCounter
            inGrid.CellForeColor = SystemColorConstants.vbGrayText
        Next
    Else
        mbolPOFinal = False
    End If
End Sub

Private Sub mDataGrid_ItemSelected()
   With mDataGrid.Grid
       SetQuantityReceived .Row, CInt(.TextMatrix(.Row, 1))
   End With
End Sub

Private Sub mDataGrid_KeyPress(KeyAscii As Integer)
    If mDataGrid.Grid.Row > 0 Then
        Select Case KeyAscii
        Case vbKey6, vbKeySpace, vbKeyAdd
            IncrementReceived 1
        Case vbKey4, vbKeySubtract
            IncrementReceived -1
        End Select
        
    End If
End Sub

Private Sub IncrementReceived(inByHowMuch As Integer)
    Dim aOrdered As Integer
    Dim aReceived As Integer
    
    With mDataGrid.Grid
        aOrdered = CInt(.TextMatrix(.Row, 1))
        aReceived = CInt(.TextMatrix(.Row, 0))
        
        aReceived = aReceived + inByHowMuch
        If aReceived > -1 And aReceived <= aOrdered Then
            SetQuantityReceived .Row, aReceived
        End If
    End With
    
End Sub

Private Sub mDataGrid_RowChange(inPreviousRow As Integer)
   '// Move the updown control to slightly left of the grid at the same position as the row
   vscrollReceived.Visible = True
    With mDataGrid.Grid
        vscrollReceived.Move mDataGrid.Left - vscrollReceived.Width - 20, mDataGrid.Top + .RowPos(mDataGrid.Grid.Row) - 100, vscrollReceived.Width, vscrollReceived.Height
        vscrollReceived.Min = 0 - CInt(.TextMatrix(.Row, 1))
        Dim aStr As String
        aStr = .TextMatrix(.Row, 0)
        If Len(aStr) = 0 Then
            vscrollReceived.Value = 0
            .TextMatrix(.Row, 0) = "0"
        Else
            vscrollReceived.Value = 0 - CInt(aStr)
        End If
        
    End With
End Sub

Private Sub vscrollReceived_Change()
    With mDataGrid.Grid
        SetQuantityReceived .Row, 0 - vscrollReceived.Value
    End With
End Sub

Private Sub SetQuantityReceived(inRow As Integer, inValue As Integer)
    mDataGrid.Grid.TextMatrix(inRow, 0) = CStr(inValue)
    mbolCurrentHasChanges = True
    mbolAnyPOChanges = True
End Sub


Private Function SaveChangesToRecordset() As POStatusEnum
    Dim aCounter As Long
    Dim aPOState As POStatusEnum
    Dim aReceived As Long
    Dim aOrdered As Long
    
    On Error GoTo errhandler
    
   Me.chkOnlyShowNotReceivedItems.Value = False
    
    aPOState = POStatusReceived
    With mDataGrid.Grid
        .Visible = False
        For aCounter = 1 To .Rows - 1
            .Row = aCounter
            mCurrentPORecordset.AbsolutePosition = aCounter
            '// Determine the state of the row
            aReceived = CInt(.TextMatrix(.Row, 0))
            aOrdered = CInt(.TextMatrix(.Row, 1))
            If (aReceived <> aOrdered) Then
                aPOState = POStatusPartial
            End If
            mCurrentPORecordset.Fields.Item(Constants.PO_DETAIL_RECEIVED) = aReceived
        Next
        .Visible = True
    End With
    DataCenter.ReconnectAndUpdate mCurrentPORecordset
    POEditing.UpdatePOStatus mPONumber, aPOState
    
       SaveChangesToRecordset = aPOState
errhandler:
    Screen.MousePointer = vbDefault
End Function
