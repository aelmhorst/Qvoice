VERSION 5.00
Begin VB.Form frmSelectShipLocation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Ship Location"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmSelectShipLocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLocations 
      Height          =   2010
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3975
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
   Begin VB.Label lblShip 
      Caption         =   "Select a Ship To Location"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmSelectShipLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mAddressInfo As AddressInfo
Private mCustomer As Customer
Public Sub Init(in_Customer As Customer)
    Set mAddressInfo = Nothing
    Set mCustomer = in_Customer
    With lstLocations
        .Clear
        .AddItem "[default]"
        .ItemData(.ListCount - 1) = 0
    End With
    Dim aRS As Recordset
    Set aRS = DataCenter.GetCustomerShipLocations(in_Customer.UniqueID)
    With lstLocations
        While Not aRS.EOF
            .AddItem aRS!vchLocationName
            .ItemData(.ListCount - 1) = aRS!iAddressID
            aRS.MoveNext
        Wend
    End With
    aRS.Close
    
End Sub

Public Property Get SelectedAddressInfo() As AddressInfo
    
    Set SelectedAddressInfo = mAddressInfo
End Property
Public Property Get AddressSelected() As Boolean
    AddressSelected = Not mAddressInfo Is Nothing
End Property

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub lstLocations_DblClick()
    ReturnShipLocation
End Sub

Private Sub ReturnShipLocation()
    
    If lstLocations.ListIndex = 0 Then
        Set mAddressInfo = mCustomer.AddressInfo
    Else
        Set mAddressInfo = DataCenter.GetAddress(lstLocations.ItemData(lstLocations.ListIndex))
    End If

    Set mCustomer = Nothing
     Me.Hide
End Sub

Private Sub OKButton_Click()
    If Me.lstLocations.ListIndex > -1 Then
        ReturnShipLocation
    Else
        Beep
    End If
End Sub
