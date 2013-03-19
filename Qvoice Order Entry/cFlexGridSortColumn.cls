VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlexGridSortColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bIsSortable As Boolean
Private m_vHiddenCol As Variant
Private m_lCol As Long
Private m_lDefaultSort As SortSettings
Private m_lAltSort As SortSettings
Private m_lCurrentSort As SortSettings

Public Property Get IsSortable() As Boolean
    IsSortable = m_bIsSortable
End Property

Public Property Let IsSortable(ByVal RHS As Boolean)
    m_bIsSortable = RHS
End Property

Public Property Let HiddenSortColumn(ByVal RHS As Variant)
    m_vHiddenCol = RHS
End Property

Public Property Get SortColumn() As Long
    If IsNull(m_vHiddenCol) Then
        SortColumn = m_lCol
    Else
        SortColumn = CLng(m_vHiddenCol)
    End If
End Property

Public Property Let SortColumn(ByVal RHS As Long)
    m_lCol = RHS
End Property

Public Property Let DefaultSortOrder(ByVal RHS As SortSettings)
    m_lDefaultSort = RHS
End Property

Public Property Let AlternateSortOrder(ByVal RHS As SortSettings)
    m_lAltSort = RHS
End Property

Public Property Get SortOrder() As SortSettings
    
    Dim lReturn As SortSettings

    If Not m_bIsSortable Then
        lReturn = flexSortNone
    Else
        If m_lCurrentSort = flexSortNone Then
            lReturn = m_lDefaultSort
            m_lCurrentSort = m_lDefaultSort
        Else
            If m_lCurrentSort = m_lDefaultSort Then
                lReturn = m_lAltSort
                m_lCurrentSort = m_lAltSort
            Else
                lReturn = m_lDefaultSort
                m_lCurrentSort = m_lDefaultSort
            End If
        End If ''If m_lCurrentSort = flexSortNone Then
    End If  '' If Not m_bIsSortable Then
    
    SortOrder = lReturn

End Property