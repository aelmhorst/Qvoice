VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl MyDataGrid 
   BackColor       =   &H80000001&
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   FillColor       =   &H80000005&
   ScaleHeight     =   3180
   ScaleWidth      =   3855
   Begin MSFlexGridLib.MSFlexGrid mGrid 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "MyDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbolDescendingSort As Boolean
Private mSortCol As Integer
Private mCurrentRow As Integer
Private mboolIsDataBinding As Boolean

Public Event Click()
Public Event KeyPress(KeyAscii As Integer)
Public Event ItemSelected()
Public Event RowChange(inPreviousRow As Integer)
Public Event FormatRow(inGrid As MSFlexGrid, inRow As Integer, inRS As Recordset)

Public Property Get Grid() As MSFlexGrid
    Set Grid = mGrid
End Property

Private Sub mGrid_Click()
    RaiseEvent Click
End Sub

Private Sub mGrid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    If mGrid.TextMatrix(Row1, mGrid.Col) > mGrid.TextMatrix(Row2, mGrid.Col) Then
        Cmp = 1
    Else
        Cmp = -1
    End If
End Sub

Private Sub mGrid_DblClick()
    If mGrid.Row > 0 Then
        SelectItem
    End If
End Sub


Private Sub mGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        If mGrid.Row > 0 Then
            SelectItem
        End If
    Else
        RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub SelectItem()
   RaiseEvent ItemSelected
End Sub

Private Sub mGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If y < mGrid.RowHeight(0) Then
    '// Determine which column was clicked on
    Dim aRunningTotal As Integer
    Dim aCol As Integer
    For aCol = 0 To mGrid.Cols - 1
        If mGrid.ColIsVisible(aCol) Then
            aRunningTotal = aRunningTotal + mGrid.ColWidth(aCol)
            If aRunningTotal > X Then
                SortGrid aCol
                Exit For
            End If
        End If
    Next
    End If
End Sub

Private Sub SortGrid(inCol As Integer)
    '// Determine if we need to sort descending or ascending
    If inCol = mSortCol Then
        mbolDescendingSort = Not mbolDescendingSort
    Else
        mbolDescendingSort = False
        mSortCol = inCol
    End If

    '// Set the grid sort
    mGrid.Col = inCol
    If mbolDescendingSort Then
        mGrid.Sort = flexSortGenericDescending
    Else
        mGrid.Sort = flexSortGenericAscending
    End If
End Sub


Private Sub mGrid_RowColChange()
    If Not mboolIsDataBinding And mGrid.Visible = True And mCurrentRow <> mGrid.Row And mGrid.Row > 0 Then
        RaiseEvent RowChange(mCurrentRow)
        mCurrentRow = mGrid.Row
    End If
End Sub

Private Sub UserControl_Initialize()
    '//ResizeGrid

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_Resize()
    ResizeGrid
End Sub

Private Sub ResizeGrid()
    mGrid.Width = Width
    mGrid.Height = Height
End Sub

Public Sub DataBindWithHiddenColumns(inRS As Recordset, inHiddenColumns() As Integer)
    Dim aCounter As Integer
    BeginDataBind
    
    Databind inRS
    
    For aCounter = 0 To UBound(inHiddenColumns)
        mGrid.ColWidth(inHiddenColumns(aCounter)) = 0
    Next
    
    EndDataBind
End Sub

Public Sub DatabindSpecificColumns(inRS As Recordset, inColumnNames() As String, inColumnHeaders() As String)
    Dim aColumnName As Variant
    Dim aOrdinal As Integer
    Dim aCounter As Integer
    Dim aCurrentCol As Integer
    
    BeginDataBind
    
    On Error GoTo errhandler
    Screen.MousePointer = MousePointerConstants.vbHourglass
    With mGrid
        .Visible = False
        .Sort = flexSortNone
        .Clear
        .Cols = UBound(inColumnNames) + 1
        .Rows = inRS.RecordCount + 1
        If .Rows < 2 Then .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .Row = 0
        For Each aColumnName In inColumnNames
            For aCounter = 0 To (inRS.Fields.Count - 1)
                If inRS.Fields(aCounter).Name = aColumnName Then
                    .Col = aCurrentCol
                    .Text = inColumnHeaders(aCurrentCol)
                    .ColAlignment(aCurrentCol) = flexAlignLeftCenter
                    .ColData(aCurrentCol) = aCounter
                    Exit For
                End If
            Next
            aCurrentCol = aCurrentCol + 1
        Next
        If Not inRS.EOF Then
            inRS.MoveFirst
            BindRows inRS
        End If
        
        .Visible = True
    End With
errhandler:
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    EndDataBind
End Sub


Public Sub Databind(inRS As Recordset)
    Dim aCounter As Integer
    
    BeginDataBind
    
    On Error GoTo errhandler
    Screen.MousePointer = MousePointerConstants.vbHourglass
    With mGrid
        .Visible = False
        .Sort = flexSortNone
        .Clear
        .Cols = inRS.Fields.Count
        .Rows = inRS.RecordCount + 1
        If .Rows < 2 Then .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .Row = 0
        For aCounter = 0 To (inRS.Fields.Count - 1)
            .Col = aCounter
            .ColData(aCounter) = aCounter
            .Text = inRS(aCounter).Name
            .ColAlignment(aCounter) = flexAlignLeftCenter
        Next
        If Not inRS.EOF Then
            inRS.MoveFirst
            BindRows inRS
        End If
        
        .Visible = True
    End With
errhandler:
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    EndDataBind
End Sub

Private Sub BindRows(inRS As Recordset)
    Dim lCounter    As Integer
    Dim lFieldCount As Integer
    Dim lFieldWidth As Integer
    Dim aText       As String
    Dim aField      As Field
    
   
    lFieldCount = mGrid.Cols - 1
    
    With mGrid
    While Not inRS.EOF
        .Col = 0
        .ColSel = lFieldCount
        .Row = inRS.AbsolutePosition
        .RowSel = .Row
        '//.Clip = inRS.GetString(adClipString, 1, vbTab, vbCrLf, "")
        For lCounter = 0 To lFieldCount
            Set aField = inRS.Fields(.ColData(lCounter))
            Select Case aField.Type
                Case DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDate, DataTypeEnum.adDBDate
                    aText = Format$(aField.Value, "YYYY.MM.DD")
                Case Else
                    aText = CStr(aField.Value & "")
            End Select
            .Col = lCounter
            .Text = aText
            If .ColWidth(lCounter) > 0 Then
                lFieldWidth = TextWidth(aText) * 1.1
                If lFieldWidth > .ColWidth(lCounter) Then
                    .ColWidth(lCounter) = lFieldWidth
                End If
            End If
        Next
        RaiseEvent FormatRow(mGrid, inRS.AbsolutePosition, inRS)
        inRS.MoveNext
    Wend
    End With

End Sub

Private Sub BeginDataBind()
    mboolIsDataBinding = True
End Sub

Private Sub EndDataBind()
    mboolIsDataBinding = False
End Sub
