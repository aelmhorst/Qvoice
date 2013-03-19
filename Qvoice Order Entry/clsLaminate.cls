VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Laminate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Sub Init(in_LaminateID As Long)
If m_lngLaminateID <> pLaminateID Then
    Changed = True
    Set rs = DataCenter.GetLaminate(pLaminateID, m_Order.Customer.ID, m_lngSlabID)
        m_lngLaminateID = pLaminateID
        m_strLamCode = rs!vchLaminateCode
        LaminateDesc = rs!vchLaminateDesc
        m_curLaminateUpch = rs!mSlabUpCharge
        m_curLamJobCharge = rs!flJobUpcharge
        m_curLamTopCharge = rs!flTopUpcharge
        m_strLamBrand = rs!vchBrandDescription
    rs.Close
    Set rs = Nothing
End If
End Sub
