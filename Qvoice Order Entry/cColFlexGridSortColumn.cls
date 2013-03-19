VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColFlexGridSortColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents m_oGrid As MSFlexGrid
Attribute m_oGrid.VB_VarHelpID = -1
Private m_colItems As Collection
Private m_sKey As String
Private m_IsOnFixedPart As Boolean
Private m_IsZeroRow As Boolean
Private m_IsZeroCol As Boolean
   
Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_colItems = Nothing
    Set m_oGrid = Nothing
    On Error GoTo 0
End Sub

Public Property Get Count() As Long
    Count = m_colItems.Count
End Property

Public Sub Add(ByVal p_lSortCol As Long _
        , ByVal p_bIsSortable As Boolean _
        , Optional ByVal p_lDefaultSort As SortSettings = flexSortNone _
        , Optional ByVal p_vAlternateSort As Variant = Null _
        , Optional ByVal p_vHiddenSortColumn As Variant = Null _
        , Optional ByVal p_vBefore As Variant _
        , Optional ByVal p_vAfter As Variant _
        )
        
    ' Use Sort Column as Key
       
    On Error GoTo Oops
    
    Dim oItem As cFlexGridSortColumn: Set oItem = New cFlexGridSortColumn
    With oItem
        .SortColumn = p_lSortCol
        .IsSortable = p_bIsSortable
        .DefaultSortOrder = p_lDefaultSort
        If IsNull(p_vAlternateSort) Then
            .AlternateSortOrder = p_lDefaultSort
        Else
            .AlternateSortOrder = CLng(p_vAlternateSort)
        End If
        .HiddenSortColumn = p_vHiddenSortColumn
    End With
    
    Dim IsPosition As Boolean: IsPosition = True
    If IsEmpty(p_vBefore) And IsEmpty(p_vAfter) Then
        IsPosition = False
    End If
    
    If IsPosition Then
        If IsEmpty(p_vBefore) Then
            m_colItems.Add oItem, CStr(p_lSortCol), p_vBefore
        Else
            m_colItems.Add oItem, CStr(p_lSortCol), , p_vAfter
        End If
    Else
        m_colItems.Add oItem, CStr(p_lSortCol)
    End If
       
    Exit Sub

Oops:
    Err.Raise Err.Number, _
            Err.Source & ">" & TypeName(Me) & "::Add()", _
            Err.Description & IIf(Right(Err.Description, 1) = ".", "", ".") & "  " _
            & "In " & TypeName(Me) & "::Add()" _
            & IIf(Erl = 0, "", ", at or after Line " & CStr(Erl) & ".")

End Sub


Public Function HasItem(p_vKey As Variant) As Boolean
    ' Return TRUE if item with specified key exists in collection
        
    Dim o As Variant
    On Error Resume Next
    Set o = Item(p_vKey)
    If Err.Number = 13 Then '' type mismatch
        Err.Clear
        o = Item(p_vKey)
    End If
    On Error GoTo 0
        
    
    Dim bReturn As Boolean
    Dim b As Boolean
    On Error Resume Next
    b = (o Is Nothing)
    If Err.Number = 0 Then
        ' Item is a valid object
        bReturn = True
    Else
        ' not an object, so check if there was an assigned value
         bReturn = IIf(IsEmpty(o), False, True)
    End If
    On Error GoTo 0

    HasItem = bReturn
End Function

Public Function Item(p_vKey As Variant) As Variant
    '    If you want to handle error for missing item entirely in calling
    ' object, first call .HasItem, then call this only if item is present.
    ' Handles both native data types and objects.  Don't know if assigning
    ' a variable or an object variable (which requires a SET), so trap for type
    ' mismatch error.
    
    On Error GoTo CheckForTypeMismatch
    
    ' Handle objects
    If VarType(p_vKey) = vbString Then
        Set Item = m_colItems.Item(CStr(p_vKey))
    Else
        Set Item = m_colItems(p_vKey)
    End If
    
    Exit Function
    
CheckForTypeMismatch:
    If Err.Number = 13 Then
        On Error GoTo Oops
        'Handle native
        If VarType(p_vKey) = vbString Then
            Item = m_colItems.Item(CStr(p_vKey))
        Else
            Item = m_colItems(p_vKey)
    End If
    
    Exit Function

    End If

Oops:
    Err.Raise Err.Number, _
            Err.Source & ">" & TypeName(Me) & "::Item()", _
            Err.Description & IIf(Right(Err.Description, 1) = ".", "", ".") & "  " _
            & "In " & TypeName(Me) & "::Item()" _
            & IIf(Erl = 0, "", ", at or after Line " & CStr(Erl) & ".")
            
End Function

Public Sub Remove(p_vKey As Variant)
    ' No error raised if item is not there.  If you want to know if item is there before
    ' attempting to delete it, call .HasItem first.
    
    On Error Resume Next
    m_colItems.Remove p_vKey
    On Error GoTo 0

End Sub

Public Property Get Items() As Collection
    Set Items = m_colItems
End Property

Public Sub Clear()
    Set m_colItems = New Collection
End Sub

Public Sub Replace(ByVal p_vItem As Variant _
            , Optional ByVal p_vKey As Variant)
    ' If item with Key exists, Remove it then add new item
    Me.Remove p_vKey
    Me.Add p_vItem, p_vKey
End Sub


Public Sub MouseDown(ByVal p_x As Single _
            , ByVal p_y As Single)

    With m_oGrid
        m_IsZeroRow = (.RowHeight(0) >= p_y)
        
        m_IsZeroCol = (p_x < .ColWidth(0))
    End With
    
End Sub

Public Sub Sort()
    
    With m_oGrid
        If m_IsZeroRow And Not m_IsZeroCol Then
            If Me.HasItem(m_oGrid.Col) Then
                Dim oSortCol As cFlexGridSortColumn
                Set oSortCol = Me.Item(m_oGrid.Col)
                
                If oSortCol.IsSortable Then
                    Dim lOldCol As Long: lOldCol = .Col
                    
                    .Col = oSortCol.SortColumn
                    .Sort = oSortCol.SortOrder
                    .Col = lOldCol
                Else
                    MsgBox "Column is not sortable", vbInformation
                End If
                
                Set oSortCol = Nothing
            End If  '' If Me.HasItem(m_oGrid.Col) Then
        End If '' If m_IsZeroCol And Not m_IsZeroRow Then
    End With

End Sub

Public Property Set Grid(RHS As MSFlexGrid)
    Set m_oGrid = RHS
End Property

Private Sub m_oGrid_Click()
    Me.Sort
End Sub

Private Sub m_oGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MouseDown x, y
End Sub
