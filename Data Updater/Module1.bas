Attribute VB_Name = "ModData"
Option Explicit
Option Compare Text


Public Sub UpdateDatabase(DatabasePath As String)
    Dim dbsQvoice As Database
    Set dbsQvoice = OpenDatabase(DatabasePath)
    UpdateTo2Point7 dbsQvoice
End Sub

Private Function IsVersion(inVersion As String, inDatabase As Database) As Boolean
    Dim aRs As Recordset
    Set aRs = inDatabase.OpenRecordset("Select vchSetting from AppSettings where vchSettingKey = ""database_version""")
    IsVersion = aRs("vchSetting").Value = inVersion
    aRs.Close
    Set aRs = Nothing
End Function

Private Function UpdateVersion(inNewVersion As String, inDatabase As Database)
    Dim aSQL As String
    aSQL = "Update AppSettings set vchSetting = """ & inNewVersion & """ where vchSettingKey = ""database_version"""
    inDatabase.Execute aSQL
End Function


Private Function UpdateTo2Point7(inDatabase As Database)
    If IsVersion("2.6.0.636", inDatabase) Then
        Update_2dot6_To_2dot7 inDatabase
        Update_2dot7_To_3dot0 inDatabase
        '// TODO: Rebuild Indexes and Relations
    ElseIf IsVersion("2.7.0.650", inDatabase) Then
        Update_2dot7_To_3dot0 inDatabase
        '// TODO: Rebuild Indexes and Relations
    Else
        MsgBox "Database is not at the correct version to do this update!", vbOKOnly, App.Title
    End If
End Function

Private Sub Update_2dot6_To_2dot7(inDatabase As Database)
    '// Change OrderLine.FlDiscountPercent to flOrderableLength
    Dim tdf As TableDef
    Dim aField As Field
    Set tdf = inDatabase.TableDefs("OrderLine")
    tdf.Fields.Delete ("flDiscountPercent")
    tdf.Fields.Append tdf.CreateField("flOrderableLength", DataTypeEnum.dbCurrency)
    'tdf.Fields.Refresh
    'inDatabase.TableDefs.Refresh
    '// Update flOrderableLength to flSlabLength
    inDatabase.Execute "Update OrderLine set flOrderableLength = flSlabLength"
    
    '// Update vOrderLine
    Dim aQueryDef As QueryDef
    Set aQueryDef = inDatabase.QueryDefs("vOrderLine")
    aQueryDef.SQL = "SELECT OrderLine.*, OrderLine.[iOrdered]-[iShipped] AS ToShip FROM OrderLine;"

    '// Add flOrderableLength to qkpovPOItems, remove flSlabLength
    Set aQueryDef = inDatabase.QueryDefs("qkpovPOItems")
    aQueryDef.SQL = "SELECT Laminate.vchBrandCode, Laminate.vchLaminateCode, Slab.vchVendorCode, Slab.vchVendorDesc, Slab.iSlabID, OrderLine.iSerialID, OrderLine.txtLineDesc, OrderLine.iLineNumber, OrderLine.flOrderableLength, OrderLine.iOrdered, Slab.iVendorSlabTypeID, VendorSlabType.iVendorID, LaminateUpcharge.tiSpecialOrder, Laminate.iPricecode, OrderLine.iOrderID, OrderLine.tiOnPo " & _
                    "FROM LaminateUpcharge INNER JOIN ((Laminate INNER JOIN (Slab INNER JOIN OrderLine ON Slab.iSlabID = OrderLine.iSlabid) ON Laminate.iLaminateID = OrderLine.iLaminateID) INNER JOIN VendorSlabType ON Slab.iVendorSlabTypeID = VendorSlabType.iVendorSlabTypeID) ON LaminateUpcharge.iPriceCode = Laminate.iPricecode;"
                    
    '// Add flOrderableLength to qkpovOrderItemsOnOrder remove flSlabLength
    Set aQueryDef = inDatabase.QueryDefs("qkpovOrderItemsOnOrder")
    aQueryDef.SQL = "PARAMETERS iOrderID Long; " & vbCrLf & _
    "SELECT qkpovPOItems.vchBrandCode, qkpovPOItems.vchLaminateCode, qkpovPOItems.vchVendorCode, qkpovPOItems.vchVendorDesc, qkpovPOItems.flOrderableLength, qkpovPOItems.iSlabID, qkpovPOItems.iSerialID, qkpovPOItems.txtLineDesc, qkpovPOItems.iLineNumber, qkpovPOItems.iOrdered, qkpovPOItems.iVendorSlabTypeID, qkpovPOItems.iVendorID, qkpovPOItems.tiSpecialOrder, qkpovPOItems.iPricecode, qkpovPOItems.iOrderID, qkpovPOItems.tiOnPo " & _
    "From qkpovPOItems " & _
    "WHERE (((qkpovPOItems.iOrderID)=[iOrderID]) AND ((qkpovPOItems.tiOnPo)=False));"

    '// Add flOrderableLength to qkpogPoItemsOnOrder remove flSlabLength
    Set aQueryDef = inDatabase.QueryDefs("qkpogPOItemsOnOrder")
    aQueryDef.SQL = "PARAMETERS iOrderID Long; " & vbCrLf & _
    "SELECT IIf([tiSpecialOrder],[qkpovOrderItemsOnOrder.vchBrandCode],""Z"") AS vchBrandCode, qkpovOrderItemsOnOrder.vchLaminateCode, qkpovOrderItemsOnOrder.vchVendorCode, qkpovOrderItemsOnOrder.vchVendorDesc, qkpovOrderItemsOnOrder.iSlabID, qkpovOrderItemsOnOrder.iSerialID, qkpovOrderItemsOnOrder.txtLineDesc, qkpovOrderItemsOnOrder.iLineNumber, qkpovOrderItemsOnOrder.flOrderableLength, qkpovOrderItemsOnOrder.iOrdered, qkpovOrderItemsOnOrder.iVendorSlabTypeID, qkpovOrderItemsOnOrder.iVendorID, qkpovOrderItemsOnOrder.tiSpecialOrder " & _
    "From qkpovOrderItemsOnOrder " & _
    "ORDER BY IIf([tiSpecialOrder],[qkpovOrderItemsOnOrder.vchBrandCode],""Z""), qkpovOrderItemsOnOrder.vchLaminateCode, qkpovOrderItemsOnOrder.vchVendorCode;"


    UpdateVersion "2.7.0.650", inDatabase
    
    MsgBox "Update to version 2.7 complete", vbInformation, App.Title
End Sub

Private Sub Update_2dot7_To_3dot0(inDatabase As Database)
    '// Add a PO Status Field to Order Header
    Dim tdf As TableDef
    Dim aField As Field
    
    '** POSTATUS
    Set tdf = inDatabase.CreateTableDef("POStatus")
    Set aField = tdf.CreateField("tiPOStatusID", DataTypeEnum.dbByte)
    tdf.Fields.Append aField
    tdf.Fields.Append tdf.CreateField("StatusText", DataTypeEnum.dbText, 20)
    inDatabase.TableDefs.Append tdf
    
    inDatabase.Execute "Insert Into POStatus Values( 0, ""New"" )"
    inDatabase.Execute "Insert Into POStatus Values( 1, ""OnPO"" )"
    inDatabase.Execute "Insert Into POStatus Values( 2, ""Partial"" )"
    inDatabase.Execute "Insert Into POStatus Values( 3, ""Received"" )"
    
    '**PURCHASEORDERDETAIL
    '// Update the PO Detail table, add the iQuantityReceived field
    Set tdf = inDatabase.TableDefs("PurchaseOrderDetail")
    Set aField = tdf.CreateField("iQuantityReceived", DataTypeEnum.dbLong)
    aField.DefaultValue = 0
    tdf.Fields.Append aField
    
    inDatabase.Execute "Update PurchaseOrderDetail Set iQuantityReceived = 0"
    
    inDatabase.Execute "Update PurchaseOrderDetail Set iQuantityReceived = iQuantity Where iPurchaseOrderID > 5"
    
    
    '** ORDERHEADER
    Set tdf = inDatabase.TableDefs("OrderHeader")
    Set aField = tdf.CreateField("tiPOStatusID", DataTypeEnum.dbByte)
    aField.DefaultValue = 0
    tdf.Fields.Append aField
    
    Set aField = tdf.CreateField("tiRush", DataTypeEnum.dbBoolean)
    aField.DefaultValue = False
    tdf.Fields.Append aField
    
    inDatabase.Execute "Update OrderHeader set tiPOStatusID = 0, tiRush = false"
        
    '// TODO: Update the order status to On PO where appropriate and Received where order is closed
    
    '// ORDERREP
    Set tdf = inDatabase.TableDefs("OrderRep")
    Set aField = tdf.CreateField("vchExtension", DataTypeEnum.dbText, 50)
    aField.DefaultValue = 0
    tdf.Fields.Append aField
    
    
        
    '**PURCHASEORDER
    Set tdf = inDatabase.TableDefs("PurchaseOrder")
    tdf.Fields.Delete "tiFinalized"
    Set aField = tdf.CreateField("iPOStatus", DataTypeEnum.dbByte)
    tdf.Fields.Append aField
    Set aField = tdf.CreateField("dtReceived", DataTypeEnum.dbDate)
    aField.OrdinalPosition = 4
    tdf.Fields.Append aField
    
    Set aField = tdf.CreateField("vchUser1", DataTypeEnum.dbText, 30)
    aField.AllowZeroLength = True
    tdf.Fields.Append aField
    
    Set aField = tdf.CreateField("vchUser2", DataTypeEnum.dbText, 30)
    aField.AllowZeroLength = True
    tdf.Fields.Append aField
    
    inDatabase.Execute "Update PurchaseOrder set iPOStatus = 3, dtReceived = dtRequested"
    
    '// Update the QueryDefs
    QueryDefSQLCreator.UpdateDatabaseQueryDefs inDatabase
    '// Update the Indexes
    
            
    '// Update the stored version in the database
    UpdateVersion "3.0.0.660", inDatabase
    MsgBox "Update to version 3.0 complete", vbInformation, App.Title
End Sub



