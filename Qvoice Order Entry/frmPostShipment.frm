VERSION 5.00
Begin VB.Form frmPostShipment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post A Shipment"
   ClientHeight    =   3150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin qkorder.ucDate ucDtDeliveryDate 
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin VB.TextBox txtShipmentID 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Post Shipment"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   $"frmPostShipment.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipment Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmPostShipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements iListener



Private Sub CancelButton_Click()
Unload Me
End Sub


Private Sub iListener_Receive(pstrText As String)
    lblStatus = pstrText
    Me.Refresh
    lblStatus.Refresh
End Sub

Private Sub OKButton_Click()
Dim llngShipmentID As Long
Dim llngAffectedOrders As Long
Dim ldtDeliveryDate As Date

If IsNumeric(txtShipmentID.Text) Then
    llngShipmentID = CLng(txtShipmentID.Text)
Else
    MsgBox "Please enter a numeric shipment number.", vbOKOnly, App.Title
    Exit Sub
End If
If ucDtDeliveryDate.isValid Then
    ldtDeliveryDate = ucDtDeliveryDate.dtDate
Else
    MsgBox "Please enter a valid delivery date.", vbOKOnly, App.Title
    Exit Sub
End If

DataCenter.PostShipment llngShipmentID, ldtDeliveryDate, _
           llngAffectedOrders, Me
           
If llngAffectedOrders = 0 Then
    MsgBox "No orders matching shipment.", vbOKOnly, App.Title
Else
    MsgBox "Posting of shipment " & llngShipmentID & _
        " completed with " & llngAffectedOrders & _
        " orders.", vbOKOnly, App.Title
End If
End Sub
