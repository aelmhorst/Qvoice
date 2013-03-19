VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbPOSlabList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Enum RecordType
    SingleOrder
    POList
End Enum

Private mRecordType As RecordType
Private mRecordset As Recordset

Public Sub MoveNext()
    mRecordset.MoveNext
End Sub

Public Property Get EOF() As Boolean
    EOF = mRecordset.EOF
End Property

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = mRecordset.AbsolutePosition
End Property

Public Property Get RecordCount() As Long
    RecordCount = mRecordset.RecordCount
End Property

Public Sub POSlabList(inRS As Recordset, inRecordType As RecordType)
    Set mRecordset = inRS
    mRecordType = inRecordType
End Sub

Public Property Get LaminateCode() As String
    LaminateCode = mRecordset!vchLaminateCode
End Property

Public Property Get VendorCode() As String
    VendorCode = mRecordset!vchVendorCode & ""
End Property

Public Property Get LengthInInches() As Currency
    If mRecordType = SingleOrder Then
        LengthInInches = mRecordset!flOrderableLength
    Else
        LengthInInches = POEditing.GetOrderableLength(mRecordset!flSize, mRecordset!dcLengthUsed)
    End If
End Property

Public Property Get SerialID() As Long
    SerialID = mRecordset!iSerialID
End Property

Public Property Get Ordered() As Integer
    Ordered = mRecordset!iOrdered
End Property

Public Property Get BrandCode() As String
    BrandCode = mRecordset!vchBrandCode
End Property

Public Property Get VendorDescription() As String
    VendorDescription = mRecordset!vchVendorDesc
End Property

Public Property Get VendorSlabType() As Integer
    VendorSlabType = mRecordset!iVendorSlabTypeID
End Property






