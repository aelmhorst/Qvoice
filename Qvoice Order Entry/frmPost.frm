VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post An Order"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "frmPost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9315
   Begin VB.CheckBox chkChargeTax 
      Caption         =   "Charge Tax"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin qkorder.ucDate ucDateDelivered 
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
      _extentx        =   2566
      _extenty        =   450
   End
   Begin VB.ComboBox cmbSignature 
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txtDeliveryCharge 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "Delivery"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optType 
      Caption         =   "Pick Up"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid lstlines 
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   5
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdChangeQuantity 
      Caption         =   "C&hange"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   4750
      Width           =   855
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post &Order"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox txtQuantityDelivered 
      Height          =   285
      Left            =   6960
      TabIndex        =   0
      Top             =   4770
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Signature"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Delivery Charge"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Delivery Date"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label LblShipped 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity Shipped"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblOrdered 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity Ordered"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity Shipped Now"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
End
Attribute VB_Name = "frmPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mOrder As Order

Public Sub Init(Order As Order)
Dim lLine As Line
Dim lstrClip As String
Dim lintCounter As Integer
Dim lrs As Recordset
Dim lintx As Integer

lstlines.Redraw = False

Set mOrder = Order

With mOrder
    Me.Caption = "Post Order " & .OrderNumber & " (" & .Customer.AddressInfo.Name & ")"
    If .Customer.Tax = 0@ Then
       Me.chkChargeTax.Enabled = False
       Me.chkChargeTax.Value = vbUnchecked
    End If
End With

lstlines.Clear
lstrClip = "Line" & vbTab & "Ordered" & vbTab & "Shipped" & vbTab & _
    "To Ship" & vbTab & "Line Description"

With lstlines
    .Redraw = False
    .ColWidth(0) = 600
    .ColWidth(1) = 600
    .ColWidth(2) = 600
    .ColWidth(3) = 600
    .ColWidth(4) = 6400
    .Row = 0
    .Col = 0
    .ColSel = 4
    .Clip = lstrClip
    .Rows = 1
End With

For Each lLine In mOrder.Lines
    lintCounter = lintCounter + 1
    lLine.ShippedNow = lLine.Unshipped
    lstrClip = CStr(lintCounter) & vbTab & _
        lLine.Ordered & vbTab & _
        lLine.Shipped & vbTab & _
        lLine.ShippedNow & vbTab & lLine.LineDescription
    With lstlines
        .AddItem lstrClip
        .Row = lintCounter
        .Col = 4
        .RowHeight(lintCounter) = _
        (((TextWidth(.Text) / _
        .ColWidth(4)) Mod 225) + 1) * 225
        For lintx = 0 To 4
            .Col = lintx
            .CellAlignment = flexAlignLeftTop
        Next
    End With
Next

Set lrs = DataCenter.GetRecentSignatures(mOrder.Customer.UniqueID)
With lrs
    Do Until .EOF
        cmbSignature.AddItem !vchSignature
        .MoveNext
    Loop
End With
lstlines.Redraw = True
End Sub



Private Sub cmbSignature_KeyPress(KeyAscii As Integer)
    AutoMatch cmbSignature, KeyAscii
End Sub

Private Sub Cmdcancel_Click()
    Set mOrder = Nothing
    Unload Me
End Sub

Private Sub cmdChangeQuantity_Click()
    Dim lLine As Line
    Dim lintQuantShipped As Integer
    
    lintQuantShipped = CInt(txtQuantityDelivered)
    
    Set lLine = mOrder.Lines(lstlines.Row)
    If lintQuantShipped >= 0 And lintQuantShipped <= lLine.Unshipped Then
        lLine.ShippedNow = lintQuantShipped
        lstlines.TextMatrix(lstlines.Row, 3) = CStr(lLine.ShippedNow)
    Else
        MsgBox "Unable to change Quantity Shipped", vbExclamation, _
                "Qk Order Entry"
    End If
End Sub



Private Sub cmdPost_Click()
    Dim ldtDelivered    As Date
    Dim lcurDelCharge   As Currency
    Dim ltresult        As VbMsgBoxResult
    Dim lbolOK          As Boolean
    Dim lstrValidation  As String
    Dim lUseLocalTax    As Boolean
    
    lbolOK = True
    
    'Validate the delivery date
    If ucDateDelivered.isValid Then
        ldtDelivered = ucDateDelivered.dtDate
        If DateAdd("d", -30, Now()) > ldtDelivered Or DateAdd("d", 10, Now()) < ldtDelivered Then
            lbolOK = False
            lstrValidation = lstrValidation & "The Delivery date must be within 30 days prior or 10 days after today's date." & vbCrLf
        End If
    Else
        lbolOK = False
        lstrValidation = lstrValidation & "You must specify a valid delivery date." & vbCrLf
    End If
    
    lcurDelCharge = 0@
    If Me.optType(1).Value Then
        '// Delivery specified
        If IsNumeric(txtDeliveryCharge) Then
           lcurDelCharge = CCur(txtDeliveryCharge)
        Else
            lcurDelCharge = 0@
        End If
    Else
        '// Local Pickup specified
        lUseLocalTax = True
    End If
    
    txtDeliveryCharge = Format$(lcurDelCharge, "0.00")
    
    If Not lbolOK Then
        MsgBox lstrValidation, vbExclamation, "Unable to Post Order"
        Exit Sub
    Else
        ltresult = MsgBox("Post order " & mOrder.OrderNumber & vbCrLf & _
                "Delivery Date: " & ldtDelivered & vbCrLf & _
                "Delivery Charge: " & Format$(lcurDelCharge, "$##,##0.00") & " ?", vbYesNo, App.Title)
        If ltresult = vbYes Then
            lstrValidation = "Document " & mOrder.OrderNumber & " has been posted to Invoice Number "
            lstrValidation = lstrValidation & DataCenter.PostOrder(mOrder, lcurDelCharge, ldtDelivered, Me.chkChargeTax.Value = vbChecked, lUseLocalTax, cmbSignature.Text)
            Me.Hide
            MsgBox lstrValidation, vbOKOnly Or vbInformation, App.Title
            Unload Me
        End If
    End If
End Sub


Private Sub lstlines_SelChange()
    Dim lLine As Line
    If lstlines.Row = 0 Then Exit Sub
    Set lLine = mOrder.Lines(lstlines.Row)
    txtQuantityDelivered.Text = lLine.ShippedNow
    
    lblOrdered = lLine.Ordered
    LblShipped = lLine.Shipped
    
    If lLine.Ordered - lLine.Posted = 0 Then
        txtQuantityDelivered.Enabled = False
    Else
        txtQuantityDelivered.Enabled = True
    End If
End Sub




