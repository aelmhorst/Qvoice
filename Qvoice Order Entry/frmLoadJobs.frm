VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoadJobs 
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10785
   Icon            =   "frmLoadJobs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10785
   Begin VB.Frame frmBottom 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   10815
      Begin VB.CommandButton cmdOpen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Open"
         Default         =   -1  'True
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmdcancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   8040
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "&All"
         Height          =   255
         Index           =   1
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "A-&F"
         Height          =   255
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "&G-J"
         Height          =   255
         Index           =   3
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "&K-R"
         Height          =   255
         Index           =   4
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "S-&V"
         Height          =   255
         Index           =   5
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "W-&Z"
         Height          =   255
         Index           =   6
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAlphaRange 
         Appearance      =   0  'Flat
         Caption         =   "&WH"
         Height          =   255
         Index           =   7
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPrompt 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin qkorder.MyDataGrid mDataGrid 
      Height          =   4815
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   9615
      _extentx        =   16960
      _extenty        =   8493
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadJobs.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadJobs.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadJobs.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLoadJobs.frx":113E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOrderType 
      Appearance      =   0  'Flat
      Caption         =   "O&rders"
      DownPicture     =   "frmLoadJobs.frx":1592
      Height          =   735
      Index           =   1
      Left            =   120
      MaskColor       =   &H8000000F&
      Picture         =   "frmLoadJobs.frx":19D4
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Uninitialized"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Label lblButtonHolder 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoadJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Private rs                  As Recordset
Private mOrderOpenReason   As OpenOrderReason
Private mbolInitialized     As Boolean
Private mbolDescendingSort   As Boolean
Private mSortCol            As Byte
Private mstrSearchString    As String
Private mstrOrdercaption    As String
Private mstrAlphacaption    As String
Private mintActiveButton    As Integer
Private mintOrderType       As Integer
Private mintFilter          As Integer
Private Const UNNASSIGNED_FILTER = -1

Private Sub RefreshJobList()
    rs.Filter = mstrSearchString
    Dim aHiddenColumns(0 To 2) As Integer
    aHiddenColumns(0) = 0
    aHiddenColumns(1) = 9
    aHiddenColumns(2) = 10
    '// Bind the grid
    mDataGrid.DataBindWithHiddenColumns rs, aHiddenColumns
           
    '// Reset the Screen Caption
    Caption = "Open " & mstrOrdercaption & " " & mstrAlphacaption & " - " & rs.RecordCount & " Records Found"
    Screen.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub cmdAlphaRange_Click(Index As Integer)
    SetFilterToIndex Index
    RefreshJobList
End Sub

Private Sub SetFilterToIndex(inIndex As Integer)
    If inIndex <> mintFilter Then
        If mintFilter > 0 Then
            With cmdAlphaRange(mintFilter)
                .FontBold = False
                .BackColor = vbButtonFace
            End With
        End If
        mintFilter = inIndex
        With cmdAlphaRange(mintFilter)
                .FontBold = True
                .BackColor = &HC0FFFF
        End With
        Select Case mintFilter
            Case Is = 1
                mstrSearchString = ""
                mstrAlphacaption = "All"
            Case Is = 2
                mstrSearchString = "Customer < 'G'"
                mstrAlphacaption = "( A - F )"
            Case Is = 3
                mstrSearchString = "Customer > 'G' AND Customer < 'K'"
                mstrAlphacaption = "( G - J )"
            Case Is = 4
                mstrSearchString = "Customer > 'K' AND Customer < 'S'"
                mstrAlphacaption = "( K - R )"
            Case Is = 5
                mstrSearchString = "Customer > 'S' AND Customer < 'W'"
                mstrAlphacaption = "( S - V )"
            Case Is = 6
                mstrSearchString = "Customer > 'W' AND Customer <> 'Wausau Homes'"
                mstrAlphacaption = "( W - Z )"
            Case Is = 7
                mstrSearchString = "Customer like 'Wausau Homes*'"
                mstrAlphacaption = "Wausau Homes"
        End Select
        
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    SelectOrder
End Sub

Private Sub SelectOrder()
    Dim aGrid As MSFlexGrid
    Set aGrid = mDataGrid.Grid
    Dim aRowValue As Long
    
    If aGrid.Row > 0 Then
        aRowValue = CLng(aGrid.RowData(aGrid.Row))
        If aRowValue > 0 Then
            DataCenter.SelectedJobID = aRowValue
            DataCenter.SelectedJobInfo = aGrid.TextMatrix(aGrid.Row, 1) & _
            " - " & aGrid.TextMatrix(aGrid.Row, 2) & _
            " - " & aGrid.TextMatrix(aGrid.Row, 3)
            
            rs.Close
            Set rs = Nothing
            SetActiveButton 0
            mintOrderType = 0
            Me.Hide
        End If
    End If
End Sub

Private Sub cmdOrderType_Click(Index As Integer)
    SetActiveButton Index
    GetOpenOrders OrderTypes(Index)
End Sub
Private Sub SetActiveButton(Index As Integer)
If mintActiveButton > 0 Then
    With cmdOrderType(mintActiveButton)
        .FontBold = False
        .Picture = Me.ImageList1.ListImages(3).Picture
    End With
End If
mintActiveButton = Index
If mintActiveButton > 0 Then
    With cmdOrderType(mintActiveButton)
        .MaskColor = vbWhite
        .FontBold = True
        .Picture = ImageList1.ListImages(4).Picture
    End With
Else
    mintOrderType = 0
End If
End Sub



Private Sub Form_Initialize()
    mintFilter = UNNASSIGNED_FILTER
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If mintOrderType > 0 Then
       SetActiveButton 0
       mintOrderType = 0
    End If
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    Dim aLeftSize As Integer, aBottomSize As Integer
    aLeftSize = Me.lblButtonHolder.Width
    aBottomSize = frmBottom.Height
    
    If Me.Width > (aLeftSize + 100) And Me.Height > (aBottomSize + 100) Then
        '// Move the three objects
        Me.lblButtonHolder.Move 0, 0, aLeftSize, Me.Height - aBottomSize
        Me.mDataGrid.Move aLeftSize, 0, Me.Width - aLeftSize - 95, Me.Height - aBottomSize
        Me.frmBottom.Move aLeftSize, mDataGrid.Height, Me.Width, aBottomSize
        ResizeCommandButtons
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "OpenOrderFilter", mintFilter
End Sub



Public Sub Init(pintOrderType As Integer, in_OpenReason As OpenOrderReason)
    Dim aOrderType As OrderType
    
    If mintFilter = UNNASSIGNED_FILTER Then
        Dim aSavedSetting As Integer
        'Default the filter to last saved filter
        aSavedSetting = GetSetting(App.Title, "Settings", "OpenOrderFilter", 2)
        SetFilterToIndex aSavedSetting
    End If
    
    mOrderOpenReason = in_OpenReason
    aOrderType = LoadCommandButtons(pintOrderType)
    GetOpenOrders aOrderType
End Sub
Private Sub GetOpenOrders(inOrderType As OrderType)
    Dim lintCounter As Integer
    Dim aMinDate As Date
    
    
    If Not mintOrderType = inOrderType.ID Then
        Screen.MousePointer = MousePointerConstants.vbHourglass
        
        mstrOrdercaption = inOrderType.Name
        
        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
        End If
        
        If mOrderOpenReason = ForCreatingPOs Then
            mstrOrdercaption = "Orders Without Pos"
            Set rs = DataCenter.GetOrdersWithoutPos(inOrderType.ID)
        Else
            mstrOrdercaption = inOrderType.Name
            If inOrderType.CreateReference Then
                aMinDate = DateAdd("yyyy", -1, Now)
            Else
                aMinDate = #1/1/1900#
            End If
            
            Set rs = DataCenter.GetOrders(inOrderType.ID, aMinDate)
        End If
        
        
        
      
        mintOrderType = inOrderType.ID
    End If
    '// Refresh the view
    RefreshJobList
    Screen.MousePointer = MousePointerConstants.vbDefault
End Sub



    Private Function LoadCommandButtons(pintOrderType As Integer) As OrderType
    Dim lintMax     As Integer
    Dim lintCounter As Integer
    Dim lintHeight  As Integer
    
    lintMax = UBound(OrderTypes)
    

    
    'Check to see if all of the buttons are loaded. If not, load
    If cmdOrderType(1).Tag = "Uninitialized" Then
        With cmdOrderType(1)
            .Caption = OrderTypes(1).Caption
            .Tag = Format$(OrderTypes(1).ID, "#")
       End With
        For lintCounter = 2 To lintMax
            Load cmdOrderType(lintCounter)
            With cmdOrderType(lintCounter)
                .Caption = OrderTypes(lintCounter).Caption
                .Tag = Format$(OrderTypes(lintCounter).ID, "#")
               .Visible = True
                .BackColor = vbButtonFace
                .FontBold = False
            End With
        Next
    End If
    
    '// If we are only opening for posting, disable any buttons
    '// that represent non-posting order types
    '// For PO creation and posting, disable buttons that are not able to be finished
    For lintCounter = 1 To lintMax
        If OrderTypes(lintCounter).ID = pintOrderType Then
            SetActiveButton lintCounter
            LoadCommandButtons = OrderTypes(lintCounter)
        End If
        cmdOrderType(lintCounter).Enabled = ((mOrderOpenReason = ForEditing) Or (OrderTypes(lintCounter).CanPost))
    Next
End Function

Private Sub ResizeCommandButtons()
    Dim aMax As Integer
    Dim aSpaceForEachButton As Integer
    Dim aCounter As Integer
    
    aMax = UBound(OrderTypes)
    aSpaceForEachButton = lblButtonHolder.Height / aMax
    
    For aCounter = 1 To aMax
        With cmdOrderType(aCounter)
            .Top = aSpaceForEachButton * (aCounter - 1) + 50
            .Height = aSpaceForEachButton - 100
        End With
    Next
 
   
End Sub



Private Sub mDataGrid_FormatRow(inGrid As MSFlexGridLib.MSFlexGrid, inRow As Integer, inRS As ADODB.Recordset)
    
    Dim lCurrentColor As Long
    lCurrentColor = Globals.GetOrderStatusColor(inRS!tiPOStatusID, inRS.Fields(Constants.ORDER_HEADER_ORDER_RUSHFLAG).Value, inRS!ReqDate)
    inGrid.Row = inRow
    
    Dim lCounter As Integer
    For lCounter = 1 To (inRS.Fields.Count - 1)
        inGrid.Col = lCounter
        inGrid.CellForeColor = lCurrentColor
        inGrid.CellAlignment = flexAlignLeftCenter
    Next
    
    inGrid.RowData(inRow) = inRS!iOrderId
    

End Sub

Private Sub mDataGrid_ItemSelected()
    SelectOrder
End Sub
