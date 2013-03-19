VERSION 5.00
Begin VB.Form frmComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComments 
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Edit the text below to add a comment to this order. This comment will appear on delivery receipts and invoices for this order."
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mOrder As Order

Public Sub Init(pOrder As Order)
    Set mOrder = pOrder
    With mOrder
        Caption = "Edit Order Comments for - " & .DescriptiveName
        txtComments.Text = .GetProperty("txtComment")
    End With
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mOrder Is Nothing Then Set mOrder = Nothing
End Sub

Private Sub OKButton_Click()
If Not mOrder Is Nothing Then
    mOrder.SetProperty "txtComment", txtComments.Text
    Set mOrder = Nothing
    Unload Me
End If
End Sub
