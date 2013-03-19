VERSION 5.00
Begin VB.UserControl ucCustomer 
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   ScaleHeight     =   2595
   ScaleWidth      =   5115
   Begin VB.Frame frmCust 
      Caption         =   "Customer Lookup"
      Height          =   2535
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4695
         Begin VB.Label lblData 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label lblData 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   4455
         End
         Begin VB.Label lblData 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label lblData 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   3
            Top             =   1200
            Width           =   4455
         End
      End
      Begin VB.ComboBox cmbCustomer 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "cmbCustomer"
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "ucCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCustomer As New Customer
Private mbolInitialized As Boolean
Private mbolAllowAllCustomers As Boolean

Public Property Let NoCustomerSelectedLabel(pstrLabel As String)
    cmbCustomer.List(0) = pstrLabel
End Property

Public Property Get Customer() As Customer
    Set Customer = mCustomer
End Property

Public Sub SetLabel(pstrLabel As String)
    frmCust.Caption = pstrLabel
End Sub

Private Sub cmbCustomer_KeyPress(KeyAscii As Integer)
        AutoMatch cmbCustomer, KeyAscii
End Sub

Private Sub cmbCustomer_LostFocus()
Dim lintCounter As Integer
    
    If cmbCustomer.ListIndex > 0 Then
        If cmbCustomer.ItemData(cmbCustomer.ListIndex) <> mCustomer.UniqueID Then
            mCustomer.Init cmbCustomer.ItemData(cmbCustomer.ListIndex)
            For lintCounter = 2 To 4
                lblData(lintCounter - 2) = mCustomer.AddressInfo.GetAddressLine(lintCounter)
            Next
        End If
    Else
        mCustomer.Init 0
        For lintCounter = 0 To 3
            lblData(lintCounter) = ""
        Next
    End If
End Sub

Private Sub UserControl_Show()
Dim rs As Recordset
    If Ambient.UserMode Then
        If Not mbolInitialized Then
            Set rs = DataCenter.GetCustomerList
            With cmbCustomer
                .AddItem "Please Select a Customer"
                Do Until rs.EOF
                    .AddItem rs!vchCustomerName: .ItemData(.NewIndex) = rs!iCustomerID
                    rs.MoveNext
                Loop
            End With
            rs.Close
            Set rs = Nothing
            cmbCustomer.ListIndex = 0
            
            mbolInitialized = True
        End If
    Else
        cmbCustomer.AddItem "Example Customer"
        cmbCustomer.ListIndex = 0
    End If
End Sub
