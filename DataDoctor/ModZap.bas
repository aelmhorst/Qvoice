Attribute VB_Name = "ModZap"
Public Sub ZapDb(inFileName As String)
  Dim db As Database
    Set db = OpenDatabase(inFileName)
    With db
        .Execute "Delete from InvoiceDetail"
        .Execute "Delete from InvoiceHeader"
        .Execute "Delete from PurchaseOrderDetail"
        .Execute "Delete from PurchaseOrderDetailMapping"
        .Execute "Delete from PurchaseOrder"
        .Execute "Delete from OrderLineCharge"
        .Execute "Delete from OrderLine"
        .Execute "Delete from OrderHeader"
        .Close
    End With
    
    Dim aO As Object
    Set aO = CreateObject("Access.Application")
    DBEngine.CompactDatabase inFileName, "compacted.mdb"
    Set aO = Nothing
End Sub
