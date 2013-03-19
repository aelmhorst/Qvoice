VERSION 5.00
Begin VB.Form frmGenerateBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shipment Generation"
   ClientHeight    =   5250
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5745
   Icon            =   "frmGenerateBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin qkorder.ucDate ucDtDelivered 
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
   End
   Begin qkorder.ucCustomer ucCust 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Generate Shipment"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Delivery Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   $"frmGenerateBatch.frx":000C
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmGenerateBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCustomer As New Customer

Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim ltresult As VbMsgBoxResult

ltresult = MsgBox("Generate Shipment?", vbYesNo, App.Title)

If ltresult = vbYes Then
    DoShipmentGeneration
End If
    
End Sub
Private Sub DoShipmentGeneration()
Dim llngRecordsAffected As Long
Dim llngBatchid As Long
Dim ltresult As VbMsgBoxResult

Dim ldtRequested As Date

ldtRequested = ucDtDelivered.dtDate

DataCenter.CreateShipment ldtRequested, llngRecordsAffected, llngBatchid, ucCust.Customer.UniqueID

If llngRecordsAffected = 0 Then
    MsgBox "No Orders met criteria.", vbOKOnly Or vbInformation, App.Title
Else
    ltresult = MsgBox("Shipment " & llngBatchid & " was generated. " & vbCrLf & _
       "Would you like to print the delivery receipts for this batch?", vbYesNo, App.Title)
    If ltresult = vbYes Then
        ReportPrinter.PrintGenericDocument "delivery_receipt", "{vOrderHeader.iBatchID}=" & llngBatchid
    End If
End If

End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With ucCust
    .SetLabel "(Optionaly) Select Customer"
    .NoCustomerSelectedLabel = "< All Customers >"
End With

End Sub


