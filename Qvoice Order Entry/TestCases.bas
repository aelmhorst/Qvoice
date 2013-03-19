Attribute VB_Name = "TestCases"
Option Explicit


Public Sub RunTestCases()


    CreatePurchaseOrdersForTestCases
End Sub


Private Sub RunIteration1()
    Dim aOrder As New Order
End Sub

Private Sub RunIteration2()
    Dim aOrder As New Order

End Sub

Private Sub RunIteration3()
    Dim aOrder As New Order

End Sub


Private Sub CreateOrderConfirmation(inOrder As Order)

End Sub

Private Sub CreatePurchaseOrdersForTestCases()
    Dim aOrderIDs As Variant
    aOrderIDs = Array(Array(mOrderID1, mOrderID2, mOrderID3))
    POEditing.GeneratePOFromStockOrderItems aOrderIDs, mDataSource.PODate
    
End Sub
