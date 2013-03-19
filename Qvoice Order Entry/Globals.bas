Attribute VB_Name = "Globals"
Option Explicit

Public Const vbToolTip = &H80000018
Public Const vbToolTipText = &H80000017

Public Const vbPurple = &H800080
Public Const vbDarkGreen = 32768


'Public Types
Public OrderTypes() As OrderType

'Private Objects
Private mMainWindow      As frmMDI
Private mData            As New cDataAccess
Private mReportPrinter   As New cReportPrinter
Private mFlicSession     As QkOrderScanListener

Public Property Get MainDocumentWindow() As frmMDI
    Set MainDocumentWindow = mMainWindow
End Property

Public Property Set MainDocumentWindow(inWindow As frmMDI)
    Set mMainWindow = inWindow
End Property

Public Property Get DataCenter() As cDataAccess
    Set DataCenter = mData
End Property

Public Property Get ReportPrinter() As cReportPrinter
    Set ReportPrinter = mReportPrinter
End Property

Public Property Set BarCodeListener(inScanner As QkOrderScanListener)
    Set mFlicSession = inScanner
End Property

Public Property Get BarCodeListener() As QkOrderScanListener
    Set BarCodeListener = mFlicSession
End Property

Public Function GetOrderStatusColor(inPoStatusID As Long, inRushFlag As Boolean, inRequestDate As Date)
    GetOrderStatusColor = vbBlack
     If inRushFlag Then
        GetOrderStatusColor = vbRed
    Else
        Select Case inPoStatusID
            Case POStatusNew
                GetOrderStatusColor = vbBlack
            Case POStatusOnPO
                '// Only show the color as red if we are within two days of request date or after
                '// Or the order is marked as Rush
                If inRequestDate < DateAdd("d", 2, Now) Then
                    GetOrderStatusColor = vbRed
                End If
            Case POStatusPartial
                GetOrderStatusColor = vbPurple
            Case POStatusReceived
                GetOrderStatusColor = vbDarkGreen
        End Select
    End If
End Function
