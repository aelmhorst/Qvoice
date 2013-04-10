VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "New Job"
   ClientHeight    =   8370
   ClientLeft      =   3540
   ClientTop       =   1815
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   11970
   Tag             =   "e"
   Begin VB.Frame frmOrder 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   6615
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   360
      Width           =   11775
      Begin qkorder.ucOERep OERepControl 
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   4800
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
      End
      Begin VB.CheckBox chkRush 
         Caption         =   "Rush Order"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Frame fraOrderDetails 
         Caption         =   "Order Specifics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5400
         TabIndex        =   25
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox cmbPOList 
            Height          =   360
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton cmdEditPO 
            Caption         =   "Edit / View"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   24
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblOrderNumber 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2160
            TabIndex        =   30
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblState 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2160
            TabIndex        =   28
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Order Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Orders ( Count of Items )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   3735
         End
      End
      Begin VB.TextBox txtTrackingCode 
         Height          =   330
         Left            =   6720
         MaxLength       =   20
         TabIndex        =   8
         Top             =   5160
         Width           =   3495
      End
      Begin VB.TextBox txtjobname 
         Height          =   330
         Left            =   6720
         TabIndex        =   3
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Frame FraCustomerLabel 
         Caption         =   "Customer Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   5055
         Begin VB.ComboBox cmbCustomer 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   1
            Text            =   "Pick a Customer from the list"
            Top             =   1920
            Width           =   4815
         End
         Begin VB.CommandButton cmdAlert 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4320
            Picture         =   "frmMain.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblPhoneNumber 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lblCustomerAddress 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   510
            Width           =   3855
         End
         Begin VB.Label lblCustomerAddress2 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   780
            Width           =   3855
         End
         Begin VB.Label lblCustomerOther 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   3855
         End
         Begin VB.Label lblDeliveryArea 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1050
            Width           =   3855
         End
      End
      Begin VB.TextBox txtPO 
         Height          =   330
         Left            =   6720
         TabIndex        =   4
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox txtShipto 
         Height          =   330
         Index           =   4
         Left            =   480
         MaxLength       =   50
         TabIndex        =   18
         Top             =   4680
         Width           =   4095
      End
      Begin VB.TextBox txtShipto 
         Height          =   330
         Index           =   5
         Left            =   480
         MaxLength       =   50
         TabIndex        =   19
         Top             =   5040
         Width           =   4095
      End
      Begin VB.TextBox txtShipto 
         Height          =   330
         Index           =   3
         Left            =   480
         MaxLength       =   50
         TabIndex        =   17
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox txtShipto 
         Height          =   330
         Index           =   1
         Left            =   480
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtShipto 
         Height          =   330
         Index           =   2
         Left            =   480
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3960
         Width           =   4095
      End
      Begin VB.ComboBox cmbOrderType 
         Height          =   360
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CheckBox chkCartoned 
         Caption         =   "Cartoned"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   5700
         Width           =   1215
      End
      Begin VB.CheckBox chkSplined 
         Caption         =   "Splined Miters"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   5700
         Width           =   1575
      End
      Begin VB.CheckBox chkPadded 
         Caption         =   "Padded"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   5700
         Width           =   975
      End
      Begin VB.CheckBox chkGroupOrder 
         Caption         =   "Put on Shipment"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   6120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSetShipLocation 
         Caption         =   "S&elect"
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   3120
         Width           =   2295
      End
      Begin qkorder.ucDate ucReqDate 
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   4440
         Width           =   3495
         _ExtentX        =   4471
         _ExtentY        =   661
      End
      Begin qkorder.ucDate ucEntryDate 
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   4080
         Width           =   3495
         _ExtentX        =   4683
         _ExtentY        =   661
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Price List"
         Height          =   255
         Left            =   5280
         TabIndex        =   70
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblPriceList 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6720
         TabIndex        =   69
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Tracking Code"
         Height          =   255
         Left            =   5280
         TabIndex        =   46
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label lblEntryDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Entry Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   45
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Name"
         Height          =   255
         Left            =   5280
         TabIndex        =   44
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Purchase Order"
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Salesperson"
         Height          =   255
         Left            =   5280
         TabIndex        =   42
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Requested Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   41
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Ship To:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Order Type"
         Height          =   255
         Left            =   5280
         TabIndex        =   39
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label txtBatchID 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6720
         TabIndex        =   38
         Top             =   5880
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.Frame frmOrder 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Index           =   2
      Left            =   0
      TabIndex        =   47
      Top             =   360
      Width           =   11700
      Begin VB.Frame FraLines 
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   1335
         Left            =   0
         TabIndex        =   48
         ToolTipText     =   "Delete Current Line"
         Top             =   120
         Width           =   11295
         Begin VB.ComboBox CmbLaminate 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   55
            ToolTipText     =   "Enter part of the laminate code or name and press [Enter]"
            Top             =   960
            Width           =   3045
         End
         Begin VB.ComboBox cmbSlab 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3240
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   270
            Width           =   5895
         End
         Begin VB.TextBox txtQuant 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   49
            Text            =   "1"
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txtLength 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   50
            Top             =   270
            Width           =   1575
         End
         Begin VB.TextBox txtCharges 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   95
            TabIndex        =   53
            Top             =   960
            Width           =   5895
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   57
            Top             =   300
            Width           =   1575
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   59
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   58
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Slab Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   56
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblChargeDescriptor 
            BackStyle       =   0  'Transparent
            Caption         =   "Additional Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   95
            TabIndex        =   54
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Slab Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   52
            Top             =   720
            Width           =   1695
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdLines 
         Height          =   5175
         Left            =   120
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1560
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9128
         _Version        =   393216
         Cols            =   5
         BackColor       =   -2147483624
         WordWrap        =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         MergeCells      =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrameTotals 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   62
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton cmdComment 
         Caption         =   "Co&mment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblOrderTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000017&
         Height          =   330
         Left            =   5400
         TabIndex        =   68
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblDiscountTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   330
         Left            =   8400
         TabIndex        =   67
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "List Price"
         Height          =   255
         Left            =   4320
         TabIndex        =   66
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   65
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000017&
         Height          =   330
         Left            =   3360
         TabIndex        =   63
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8070
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "4/9/2013"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9:19 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip MainTab 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      TabWidthStyle   =   2
      TabFixedWidth   =   8819
      TabMinWidth     =   2117
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Job Information"
            Object.ToolTipText     =   "View Job Details"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Job &Line Details"
            Object.ToolTipText     =   "View Line Item Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'   frmMain
'   Created By:     Andy Elmhorst
'   Purpose:        The main user interface screen for an order
'************************************************************************
Option Explicit
Option Compare Text

Private WithEvents mOrder           As Order
Attribute mOrder.VB_VarHelpID = -1
Private mbolDirty                   As Boolean
Private mbolNewOrder                As Boolean
Private mlngCurrentLine             As Long
Private mbolEditMode                As Boolean
Private mbolLoadingOrder            As Boolean

Private Const vbToolTipText = &H80000017


Public Sub SetShipTo(in_Line As Integer, in_String As String)
    txtShipto(in_Line).Text = in_String
End Sub


Public Sub PrintDeliveryReceipt()
    Dim aDialog As frmPrintOrderDocuments
    Set aDialog = GetOrderDocumentPrintDialog("printing delivery receipt")
    If Not aDialog Is Nothing Then
        aDialog.EnableDeliveryReceiptPrint
        aDialog.EnableLabelPrint
        aDialog.Show vbModal
    End If
End Sub

Public Sub PrintOrderDocument()
    Dim aDialog As frmPrintOrderDocuments
    Set aDialog = GetOrderDocumentPrintDialog("printing quote or order confirmation")
    If Not aDialog Is Nothing Then
        aDialog.EnableOrderConfirmationPrint
        aDialog.Show vbModal
    End If
End Sub

Public Sub PrintLabels()
    Dim aDialog As frmPrintOrderDocuments
    Set aDialog = GetOrderDocumentPrintDialog("printing labels")
    If Not aDialog Is Nothing Then
        aDialog.EnableLabelPrint
        aDialog.Show vbModal
    End If
End Sub

Public Sub PreviewOrderStatus()
    ReportPrinter.PrintGenericDocument Constants.REPORT_NAME_ORDER_SLAB_STATUS, _
        "{OrderHeader.iOrderID}=" & mOrder.OrderID, False, True, "Order Status", Globals.MainDocumentWindow.hwnd, False
End Sub



Private Function GetOrderDocumentPrintDialog(inAction As String) As frmPrintOrderDocuments
    If PromptBeforeContinuing(inAction) <> vbCancel Then
        If mOrder.OrderID = 0 Then
            MsgBox "Unable to print an unsaved order.", vbOKOnly Or vbInformation, App.Title
            Set GetOrderDocumentPrintDialog = Nothing
        Else
            Dim aDialog As New frmPrintOrderDocuments
            aDialog.SetOrderInfo Me.OrderInfo, Me.OrderID
            Set GetOrderDocumentPrintDialog = aDialog
        End If
    End If
    
End Function

Private Sub SetPrintMenuDescription()
If Not mOrder Is Nothing Then
    Globals.MainDocumentWindow.mnuPrintDoc.Enabled = Not (IsTemplate(mOrder.JobType) Or CreateReference(mOrder.JobType))
    'Dirty = True
End If
End Sub

Public Property Get Dirty() As Boolean
    Dirty = mbolDirty
End Property

Public Property Let Dirty(pbolDirtyFlag As Boolean)
If Not mbolLoadingOrder Then
    If mOrder.Editable Then
        mbolDirty = pbolDirtyFlag
    End If
End If
End Property



Public Property Let EditMode(InEditMode As Boolean)
    mbolEditMode = InEditMode
    cmdOK.Caption = IIf(mbolEditMode, "C&hange", "&Add")
End Property

Public Property Get EditMode() As Boolean
    EditMode = mbolEditMode
End Property




Private Sub chkCartoned_Click()
    Dirty = True
End Sub

Private Sub chkCartoned_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chkSplined.SetFocus
End Sub

Private Sub chkGroupOrder_Click()
    Dirty = True
End Sub

Private Sub chkGroupOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmbCustomer.SetFocus
End Sub

Private Sub chkPadded_Click()
    Dirty = True
End Sub

Private Sub chkPadded_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chkGroupOrder.SetFocus
End Sub

Private Sub chkRush_Click()
    Dirty = True
End Sub

Private Sub chkSplined_Click()
    Dirty = True
End Sub


Private Sub chkSplined_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chkPadded.SetFocus
End Sub

Private Sub cmbCustomer_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmbOrderType.SetFocus
Else
    AutoMatch cmbCustomer, KeyAscii
End If
End Sub

Private Sub cmbCustomer_LostFocus()
Dim lintCounter As Integer

If cmbCustomer.ListIndex = -1 Then
    With cmbCustomer
        .BackColor = vbToolTip
        .ForeColor = vbToolTipText
    End With
    Exit Sub
End If

With cmbCustomer
    .BackColor = vbWindowBackground
    .ForeColor = vbWindowText
End With
If mOrder.Customer.UniqueID <> cmbCustomer.ItemData(cmbCustomer.ListIndex) Then
    Dirty = True
    mOrder.Customerid = cmbCustomer.ItemData(cmbCustomer.ListIndex)
    ShowCustomerDataArea
    With mOrder.Customer
        If mOrder.Editable Then
            For lintCounter = 1 To 5
                txtShipto(lintCounter).Text = .AddressInfo.GetAddressLine(lintCounter)
            Next

            If .PopUpAlert Then ShowCustomerAlert

        End If
        chkCartoned.Value = IIf(.PrefCartoned, vbChecked, vbUnchecked)
        chkSplined.Value = IIf(.prefSplined, vbChecked, vbUnchecked)
        chkPadded.Value = IIf(.PrefPadded, vbChecked, vbUnchecked)
        chkGroupOrder.Value = IIf(.PrefGroupedOrders, vbChecked, vbUnchecked)
        RefreshSavedOrderDetails
        RefreshLineDetailsGrid
    End With
End If
End Sub




Private Sub ShowCustomerDataArea()
        
With mOrder.Customer
    FraCustomerLabel.Caption = Replace(.AddressInfo.Name, "&", "&&")
    lblPhoneNumber.Caption = Format$(.AddressInfo.Phone, "(000) ###-####")
    lblCustomerAddress = .AddressInfo.Address1
    lblCustomerAddress2 = .AddressInfo.City & ", " & .AddressInfo.State & "  " & .AddressInfo.Zip
    lblCustomerOther = "Tax: " & _
        Format$(.Tax, "#0.0%") & _
        "   Discount: " & Format$(.Discount, "#0.0%")
        
    lblDeliveryArea = .DeliveryDay & " - " & .DeliveryArea
    cmdAlert.Visible = .HasAlert
End With
Caption = mOrder.DescriptiveName


End Sub

Private Sub CmbLaminate_GotFocus()
CmbLaminate.SelStart = 0
CmbLaminate.SelLength = Len(CmbLaminate.Text)
End Sub

Private Sub cmbLaminate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    With CmbLaminate
        Select Case Len(.Text)
        Case Is = 0
             Exit Sub
        Case Is > 8
            If .ListIndex > -1 Then
                SendKeys "{TAB}"
            Else
                FillLaminateCombo .Text
            End If
        Case Else
            FillLaminateCombo .Text
        End Select
    End With
End If
End Sub
Private Sub FillLaminateCombo(pstrSearch As String)
    Dim rs          As Recordset
    Dim lstrString  As String
    Screen.MousePointer = MousePointerConstants.vbHourglass
    Set rs = DataCenter.GetColorMatches(pstrSearch, mOrder.PriceListID)
    Screen.MousePointer = MousePointerConstants.vbDefault
    
    With CmbLaminate
        lstrString = .Text
        .Clear
        Do Until rs.EOF
            .AddItem Format$(rs!vchLaminateCode, "@@@@@@@@!") & rs!vchLaminateDesc
            .ItemData(.NewIndex) = rs!iLaminateID
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        If .ListCount > 0 Then
            .ListIndex = 0
        Else
            .Text = lstrString
        End If
       ' .SelStart = Len(lstrString) - 1
       ' AutoMatch CmbLaminate, Asc(Right$(lstrString, 1))
       ' .SelStart = 0
    End With
End Sub

Private Sub cmbOERep_Change()
    Dirty = True
End Sub

Private Sub cmbOERep_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtTrackingCode.SetFocus
End Sub




Private Sub cmbOrderType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtjobname.SetFocus
End Sub

Private Sub cmbOrderType_LostFocus()
    If Not mOrder Is Nothing Then
        If Not mOrder.JobType = cmbOrderType.ItemData(cmbOrderType.ListIndex) Then
            mOrder.JobType = cmbOrderType.ItemData(cmbOrderType.ListIndex)
            SetPrintMenuDescription
            Dirty = True
            End If
    End If
End Sub


Private Sub cmbSlab_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtCharges.SetFocus
Else
    AutoMatch cmbSlab, KeyAscii
End If
End Sub

Private Sub cmbSlab_LostFocus()
With cmbSlab
    If .ListIndex = -1 Then
        .BackColor = vbToolTip
        .ForeColor = vbToolTipText
    Else
        .BackColor = vbWindowBackground
        .ForeColor = vbWindowText
    End If
End With
End Sub

Private Sub cmdAlert_Click()
    ShowCustomerAlert
End Sub
Private Sub ShowCustomerAlert()
    MsgBox mOrder.Customer.Alert, vbOKOnly Or vbInformation, App.Title
End Sub

Private Sub cmdCancel_Click()
If EditMode Then
    SwitchMode False
    RefreshLineDetailsGrid
End If
End Sub

Private Sub DeleteRow(Row As Integer)
Dim lbolOK  As Boolean

lbolOK = mOrder.DeleteLine(Row)

If Not lbolOK Then
    MsgBox mOrder, vbExclamation, _
        "Order Entry Validation Error"
    Exit Sub
Else
    Dirty = True
    grdLines.Rows = grdLines.Rows - 1
End If
RefreshLineDetailsGrid
RefreshOrderTotalsAndSpecialCharge
End Sub

Public Function Save() As Boolean
On Error GoTo errhandler:

If mOrder.Editable Then
    If Dirty Then
        If ValidateScreen Then
            mOrder.Save
            Dirty = False
            RefreshSavedOrderDetails
            Save = True
            
        Else
            Save = False
        End If
    End If
Else
    MsgBox "Cannot Save order that has already been posted!", vbExclamation, App.Title
    Save = False
End If
Exit Function
errhandler:
    Screen.MousePointer = MousePointerConstants.vbDefault
    HandleError "An error has occured during Order Save." & vbCrLf & _
            Err.Description & " - " & Err.Source, False
End Function
Private Sub ClearScreen()
Dim lconControl As Control

grdLines.Rows = 1

For Each lconControl In Controls

 Select Case TypeName(lconControl)
    Case Is = "TextBox"
        lconControl = ""
    Case Is = "ComboBox"
        lconControl.ListIndex = -1
        lconControl.Text = ""
 End Select
 
Next
End Sub



Private Sub cmdComment_Click()
    If Not mOrder Is Nothing Then
        With frmComments
            .Init mOrder
            .Show vbModal, Globals.MainDocumentWindow
        End With
    End If
    Dirty = True
End Sub

Private Sub cmdEditPO_Click()
    If Me.cmbPOList.ListIndex > -1 Then
        If cmbPOList.ItemData(cmbPOList.ListIndex) = -1 Then
            '// Create PO's for this order
            cmbPOList.Clear
            POEditing.HandlePOCreation mOrder, OERepControl.SelectedRepName, OERepControl.SelectedRepExtension
        Else
            Dim aController As New POPurchaseOrderEditController
            aController.POPurchaseOrderEditController cmbPOList.Text, cmbPOList.ItemData(cmbPOList.ListIndex)
        End If
    End If
End Sub

Private Sub cmdOK_Click()
Dim lcurlinetotal   As Currency
Dim lobjLine        As Line
Dim lcurLength      As Currency
Dim lcurWidth       As Currency
Dim lintx           As Integer
Dim lstrValidation  As String

Dirty = True

If txtLength Like "*[*]*" Then
    lintx = InStr(txtLength, "*")
    If IsNumeric(Left$(txtLength, lintx - 1)) Then
        lcurLength = CCur(Left$(txtLength, lintx - 1))
    Else
        lcurLength = 0@
    End If
    If IsNumeric(Mid$(txtLength, lintx + 1)) Then
        lcurWidth = CCur(Mid$(txtLength, lintx + 1))
    Else
        lcurWidth = 0@
    End If
    txtLength = Format$(lcurLength, "##0.000") & "*" & Format$(lcurWidth, "##0.000")
Else
    If IsNumeric(txtLength) Then
        lcurLength = CCur(txtLength)
    Else
        lcurLength = 0@
    End If
    lcurWidth = 0@
    txtLength = Format$(lcurLength, "##0.000")
End If

If lcurLength < 1 Then
    If Trim$(cmbSlab.Text) <> "[none]" Then lstrValidation = "Please indicate the length of this slab." & vbCrLf
End If


If Not LineFieldsValid(lstrValidation) Then
    MsgBox lstrValidation, vbOKOnly Or vbExclamation, App.Title
    Exit Sub
End If

If Not mOrder.Editable Then
    MsgBox "Order Line cannot be changed." & vbCrLf & _
        "This order has already been posted!", vbExclamation, _
        "Order Entry Validation Error"
    Exit Sub
End If

If EditMode Then
    Set lobjLine = mOrder.Lines(grdLines.Row)
    With lobjLine
        If CmbLaminate.ListIndex > -1 Then
            .Add mOrder, grdLines.Row, CLng(txtQuant), lcurLength, lcurWidth, _
                cmbSlab.ItemData(cmbSlab.ListIndex), CmbLaminate.ItemData(CmbLaminate.ListIndex)
        Else
            .Add mOrder, grdLines.Row, CLng(txtQuant), lcurLength, lcurWidth, _
                cmbSlab.ItemData(cmbSlab.ListIndex), 0
        End If
        ProcessChargeLine .Charges
    End With
    SwitchMode False
    RefreshLineDetailsGrid
Else
    If CmbLaminate.ListIndex > -1 Then
        Set lobjLine = mOrder.Lines.Add(CLng(txtQuant), lcurLength, lcurWidth, _
            cmbSlab.ItemData(cmbSlab.ListIndex), _
            CmbLaminate.ItemData(CmbLaminate.ListIndex))
    Else
        Set lobjLine = mOrder.Lines.Add(CLng(txtQuant), lcurLength, lcurWidth, _
            cmbSlab.ItemData(cmbSlab.ListIndex), 0)
    End If
    With lobjLine
    ProcessChargeLine .Charges
    FillGrid .LineNumber, .Ordered, .LineDescription, _
        .List, .Total
    End With
End If
    
txtCharges = ""
txtQuant = "1"
txtLength.SetFocus
RefreshOrderTotalsAndSpecialCharge
End Sub



Private Sub ProcessChargeLine(pobjCharges As qkorder.Charges)
Dim lvarcharges     As Variant
Dim lintUbound      As Integer
Dim lintTotal       As Integer
Dim lintx           As Integer
Dim lintDash        As Integer
Dim lstrAbbrev      As String
Dim lstrDesc        As String
Dim lcurPrice       As Currency
Dim lobjCharge      As Charge

'Make a variant array out of the charge line, based upon the
'current seperator character
lvarcharges = Split(Trim$(txtCharges), gstrSeperator)

lintUbound = UBound(lvarcharges)

For lintx = 0 To lintUbound
    
    lintDash = InStr(lvarcharges(lintx), "-")
    lstrAbbrev = Mid$(lvarcharges(lintx), lintDash + 1)
    If chkSplined = vbChecked Then
        If lstrAbbrev = "lm" Or lstrAbbrev = "rm" Or lstrAbbrev = "m" Then
            lstrAbbrev = lstrAbbrev & "s"
        End If
    End If
    If lstrAbbrev Like "'*" Or lstrAbbrev Like "misc" Then
        GoSub gsTextCharge
    Else
        Set lobjCharge = DataCenter.GetCharge(lstrAbbrev, mOrder.PriceListID)
    End If
    
    If lintDash > 0 Then
            lobjCharge.Quantity = CInt(Left$(lvarcharges(lintx), lintDash - 1))
    Else
            lobjCharge.Quantity = 1
    End If
    
    If Len(lobjCharge.Abbrev) > 0 Then
        lintTotal = lintTotal + 1
        If lintTotal > pobjCharges.Count Then
            With lobjCharge
                pobjCharges.Add .iChargeID, .Abbrev, .Description, .Price, .Quantity, .Lineal, .LengthAdjustment
            End With
        Else
           pobjCharges(lintTotal).Init lobjCharge.iChargeID, lobjCharge.Quantity, _
            lobjCharge.Abbrev, lobjCharge.Price, lobjCharge.Description, lobjCharge.Lineal, lobjCharge.LengthAdjustment
        End If
    End If
Next
Do While pobjCharges.Count > lintTotal
    pobjCharges.Remove pobjCharges.Count
Loop


Exit Sub
gsTextCharge:
    Set lobjCharge = New Charge
        If lstrAbbrev Like "'*" Then
            lobjCharge.Init 0, 0, "Text", 0, Mid$(lstrAbbrev, 2), False, 0@
        Else
            lstrDesc = InputBox("You have entered the code for a miscellaneous charge" & _
                vbCrLf & "Please enter the description below", "Charge Processing", "Description")
            
            Dim lstrPrice As String
            lstrPrice = _
                InputBox( _
                "Please enter the amount the charge for " & lstrDesc, _
                "Charge Processing", _
                "0.00")
                If IsNumeric(lstrPrice) Then
                    lcurPrice = CCur(lstrPrice)
                Else
                    lcurPrice = 0
                End If
                    
            lobjCharge.Init 0, 0, lstrAbbrev, lcurPrice, lstrDesc, False, 0@
        End If
    
Return
End Sub
Private Sub RefreshLineDetailsGrid()
Dim lintCounter As Integer
Dim lclsline As Line
Dim lbolShowBackorders

grdLines.Redraw = False
grdLines.Clear

lbolShowBackorders = Not mOrder.Editable

For lintCounter = 1 To mOrder.Lines.Count
    Set lclsline = mOrder.Lines(lintCounter)
    With lclsline
        If lbolShowBackorders And (.Ordered > .Shipped) Then
            FillGrid lintCounter, .Shipped & "/" & .Ordered, _
                .LineDescription, .List, .Total, True
        Else
            FillGrid lintCounter, CStr(.Ordered), _
                .LineDescription, .List, .Total
        End If
    End With
Next lintCounter

grdLines.Redraw = True
End Sub

Private Sub FillGrid(lintLine As Integer, _
                 lstrQuantity As String, _
                 lstrDescription As String, _
                 lsnglinetotal As Single, _
                 lsngCustomerPrice As Single, _
                 Optional in_boolAccentuate As Boolean = False)
                
    Dim lstrClipstring As String
    Dim lintx As Integer
    lstrClipstring = CStr(lintLine) & vbTab & _
                    lstrQuantity & vbTab & _
                    lstrDescription & vbTab & _
                    Format$(lsnglinetotal, "$##,##0.00") & vbTab & _
                    Format$(lsngCustomerPrice, "$##,##0.00")
    
    With grdLines
        
        If .Rows < lintLine + 1 Then .Rows = lintLine + 1
        .Row = lintLine
        
        For lintx = 0 To 4
            .Col = lintx
            Select Case lintx
            Case 0, 2
                .CellAlignment = flexAlignLeftTop
            Case 1
                .CellAlignment = flexAlignLeftTop
                If in_boolAccentuate Then
                   .CellFontBold = True
                   .CellForeColor = vbRed
                End If
            Case Else
                .CellAlignment = flexAlignLeftBottom
            End Select
        Next lintx
        
        .Col = 0
        .ColSel = 4
        .Clip = lstrClipstring
        If lintLine Mod 2 = 0 Then
            .CellBackColor = vbToolTip
        Else
            .CellBackColor = vbWindowBackground
        End If
        ' Funky algorithm for getting a valid line width
        Dim aHeight As Integer
        aHeight = Me.TextHeight(lstrDescription)
        .RowHeight(lintLine) = (((Me.TextWidth(lstrDescription) / .ColWidth(2)) Mod aHeight) + 1) * aHeight
        If lintLine > 5 Then .TopRow = lintLine - 5
        
    End With
End Sub

Private Sub cmdSetShipLocation_Click()
    Dim aFrm As frmSelectShipLocation
    Set aFrm = New frmSelectShipLocation
    aFrm.Init mOrder.Customer
    aFrm.Show vbModal
    If aFrm.AddressSelected Then
        Dim a As Integer
        For a = 1 To 5
            txtShipto(a).Text = aFrm.SelectedAddressInfo.GetAddressLine(a)
        Next
    End If
    Unload aFrm
End Sub




Private Sub CmbPOList_DropDown()
    If cmbPOList.ListCount = 0 And Not mbolNewOrder Then
        Dim aRS As Recordset
        Set aRS = DataCenter.getPurchaseOrderList(mOrder.OrderID)
        
        Do While Not aRS.EOF
            If aRS.Fields.Item(Constants.PURCHASE_ORDER_ID).Value < 0 Then
                cmbPOList.AddItem "Stock (" & aRS!iCount & ")"
            Else
                cmbPOList.AddItem aRS!PO & " (" & aRS!iCount & ")"
            End If
                cmbPOList.ItemData(cmbPOList.NewIndex) = aRS.Fields.Item(Constants.PURCHASE_ORDER_ID).Value
            aRS.MoveNext
        Loop
        
        'If cmbPOList.ListCount = 0 Then
            cmbPOList.AddItem "Create New . . ."
            cmbPOList.ItemData(cmbPOList.NewIndex) = -1
        'End If
    End If
    Me.cmdEditPO.Enabled = cmbPOList.ListCount > 0
End Sub



Private Sub Form_Activate()
    SetPrintMenuDescription
End Sub

Private Sub Form_Load()
    InitializeMe
    With Me
        .Left = GetSetting("qkorder", "Settings_" & .Name, "MainLeft", 1000)
        .Top = GetSetting("qkorder", "Settings_" & .Name, "MainTop", 1000)
        .Width = GetSetting("qkorder", "Settings_" & .Name, "MainWidth", 6500)
        .Height = GetSetting("qkorder", "Settings_" & .Name, "MainHeight", 6500)
        Set .ucReqDate.NextControl = OERepControl
    End With
    
End Sub
Private Sub InitializeMe()
    Dim rs As Recordset
    Dim lintCounter As Integer
    Dim lintMax     As Integer
    
    

    
    Set rs = Nothing
    
    'Initialize the Customer List dropdown
    Set rs = DataCenter.GetCustomerList
    With cmbCustomer
        Do Until rs.EOF
            .AddItem rs!vchCustomerName: .ItemData(.NewIndex) = rs!iCustomerID
            rs.MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
    
    'Initialize the Order Type dropdown list
    lintMax = UBound(OrderTypes)
    With cmbOrderType
        For lintCounter = 1 To lintMax
            .AddItem OrderTypes(lintCounter).Name
            .ItemData(.NewIndex) = OrderTypes(lintCounter).ID
        Next
        '.ListIndex = 0
    End With
    
    lblChargeDescriptor = "Additional Charges  (Seperated by: " & _
                gstrSeperator & " )"
    
    With grdLines
        .Row = 0
        .Col = 0
        .ColSel = 4
        .Clip = "#" & vbTab & "Q" & vbTab & _
                "Description" & vbTab & "Total" & vbTab & "Discounted"
        .ColWidth(0) = 400
        .ColWidth(1) = 400
        .ColWidth(2) = .Width - 2800 - 120
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        Me.Font.Size = .Font.Size
    End With
    
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ltresult As VbMsgBoxResult
        ltresult = PromptBeforeContinuing("closing")
        If ltresult = vbCancel Then Cancel = True
    If Not Cancel Then
        If Me.WindowState = vbNormal Then
            With Me
                SaveSetting "qkorder", "Settings_" & .Name, "LastState", CStr(Me.WindowState)
                SaveSetting "qkorder", "Settings_" & .Name, "MainLeft", Me.Left
                SaveSetting "qkorder", "Settings_" & .Name, "MainTop", Me.Top
                SaveSetting "qkorder", "Settings_" & .Name, "MainWidth", Me.Width
                SaveSetting "qkorder", "Settings_" & .Name, "MainHeight", Me.Height
            End With
        End If
    End If
    
    '// Only if this is a new order and it's been saved,
    '// and not a quote or template, we will do the PO creation
    If mbolNewOrder And mOrder.OrderID > 0 And POEditing.CanCreatePurchaseOrders(mOrder) Then
        POEditing.HandlePOCreation mOrder, OERepControl.SelectedRepName, OERepControl.SelectedRepExtension
    End If

End Sub

Public Sub NewJob(plngJobType As Long)
    'Set the order object up to receive a new order
    mbolNewOrder = True
    
    Set mOrder = New Order
    mOrder.NewOrder plngJobType
    
    Dirty = True
    
    mlngCurrentLine = 0
   ' MainTab.Index = 1
    frmOrder(1).ZOrder
    cmbOrderType.ListIndex = 0
    

End Sub



Private Sub RefreshOrderTotalsAndSpecialCharge()
'// Refreshes the order totals after they have changed
Dim lcCharge    As Charge
Dim lcurTotal   As Currency

With mOrder
    lblDiscount = Format$(.Customer.Discount * 100, "##.0") & "%"
    lblOrderTotal = Format$(.List, "$###,###.00")
    lblDiscountTotal = Format$((.Total), "$###,###.00")
    Set lcCharge = .SpecialCharge
End With
With lcCharge
    FillGrid mOrder.Lines.Count + 1, .Quantity, _
    .Description, .Price, _
    .Price * (mOrder.Customer.Multiplier)
End With
    
End Sub

Private Sub ShowMessage(Message As String)
sbStatus.Panels(1).Text = Message
sbStatus.Refresh
End Sub
Private Sub RefillFields(Emptythem As Boolean)
    Dim lcharge As Charge
    
    txtCharges = ""
    If Emptythem Then
        txtQuant = "1"
        txtLength = ""
        CmbLaminate = ""
        cmbSlab = ""
        txtLength.SetFocus
    Else
        With mOrder.Lines(mOrder.LineNumber)
            For Each lcharge In .Charges
                txtCharges = txtCharges & lcharge.Abbrev & gstrSeperator
            Next
            txtQuant = .Ordered
            txtLength = .SlabLength
            CmbLaminate.SetFocus
            SendKeys .LaminateCode
            SendKeys "~"
            DoEvents
            cmbSlab.SetFocus
            SendKeys .SlabCode
            SendKeys "~"
            DoEvents
            txtCharges.SetFocus
        End With
    End If
End Sub



Private Sub grdLines_Click()
    Static mlngRowID As Long
    Static currentTic As Long
    
    If grdLines.Row = mlngRowID And Timer - currentTic <= 1 Then
        If Not mOrder.Editable Then
            MsgBox "Order Line cannot be changed." & vbCrLf & _
                "This order has already been posted!", vbExclamation, _
                "Order Entry Validation Error"
            Exit Sub
        End If
        SwitchMode True
    Else
        mlngRowID = grdLines.Row
        currentTic = Timer
    End If
End Sub

Private Sub SwitchMode(Edit As Boolean)
If Edit Then
    If grdLines.Row <> grdLines.Rows - 1 Then
        mbolEditMode = True
        grdLines.Col = 0
        cmdOK.Caption = "C&hange Line " & grdLines.Text
        grdLines.ColSel = (grdLines.Cols - 1)
        grdLines.Enabled = False
        FraLines.ForeColor = vbRed
        mOrder.LineNumber = grdLines.Row
        RefillFields False
    End If
Else
    RefillFields True
    mbolEditMode = False
    FraLines.BackColor = vbButtonFace
    grdLines.Enabled = True
    cmdOK.Caption = "&Add"
End If
End Sub


Private Sub grdLines_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete _
    And grdLines.Row <> grdLines.Rows - 1 _
    Then DeleteRow grdLines.Row
End Sub

Private Sub MainTab_Click()
    Dim lstrValidation As String
    If MainTab.Tabs(2).Selected Then
        If FieldsValid(lstrValidation) Then
            SelectTab 2
        Else
            MainTab.Tabs(1).Selected = True
            MsgBox lstrValidation, vbOKOnly Or vbExclamation, App.Title
        End If
    Else
        SelectTab 1
    End If
End Sub

Private Sub MainTab_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And MainTab.Tabs(1).Selected Then
        MainTab.Tabs(2).Selected = True
    End If
End Sub

Private Sub mOrder_CustomerChanged()
    lblPriceList.Caption = DataCenter.GetPriceListName(mOrder.PriceListID)
    Dim rs As Recordset
       'Initialize the slab List dropdown
     With cmbSlab
    
        .Clear
        Set rs = DataCenter.GetSlabList(mOrder.PriceListID)
        Do Until rs.EOF
            .AddItem Format$(rs!vchSlabCode, "@@@@@@@!") & rs!vchSlabDesc: .ItemData(.NewIndex) = rs("iSlabID")
            rs.MoveNext
        Loop
        rs.Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub mOrder_Message(lstrMessage As String)
    ShowMessage lstrMessage
End Sub

Private Sub mOrder_Saving(strMessage As String, dblPercentDone As Double)
    sbStatus.Panels(1).Text = Format$(dblPercentDone, "#00.0") & "%  " & strMessage
End Sub



Private Function FieldsValid(pstrValidation As String)

    If cmbCustomer.ListIndex = -1 Then pstrValidation = "Please select a customer." & vbCrLf
    If cmbOrderType.ListIndex = -1 Then pstrValidation = pstrValidation & "Please select an order type." & vbCrLf
    If Not ucReqDate.isValid Then pstrValidation = pstrValidation & "Please enter a valid request date." & vbCrLf
    If OERepControl.SelectedRepID = 0 Then pstrValidation = pstrValidation & "Please select a salesperson." & vbCrLf
    
    FieldsValid = Len(pstrValidation) = 0
End Function
Private Function LineFieldsValid(pstrValidation As String)
    If Not IsNumeric(txtQuant) Then
        pstrValidation = pstrValidation & "Please indicate quantity." & vbCrLf
    ElseIf CInt(txtQuant) > 255 Then
        pstrValidation = pstrValidation & "Quantity must be less than 256" & vbCrLf
    End If
    If cmbSlab.ListIndex = -1 Then
        pstrValidation = pstrValidation & "Please select a slab type to use, or  "" "" for none." & vbCrLf
    Else
        If CmbLaminate.ListIndex = -1 And Trim$(cmbSlab.Text) <> "[none]" Then
            pstrValidation = pstrValidation & "Please select a laminate color for " & cmbSlab.Text
        End If
    End If
    LineFieldsValid = Len(pstrValidation) = 0
End Function



Private Sub OERepControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtTrackingCode.SetFocus
    End If
End Sub

Private Sub txtCharges_GotFocus()
    If Len(txtCharges) = 0 Then
        With txtCharges
            If chkCartoned.Value = vbChecked Then .Text = "ca/"
            If chkPadded.Value = vbChecked Then .Text = .Text & "pa/"
            .SelStart = Len(.Text)
        End With
    Else
        SelAllText txtCharges
    End If
End Sub

Private Sub txtCharges_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CmbLaminate.SetFocus
        KeyAscii = 0
    End If
End Sub



Private Sub txtjobname_Change()
    Dirty = True
End Sub
Public Property Get OrderInfo() As String
    OrderInfo = mOrder.DescriptiveName
End Property
Public Property Get OrderID() As Long
    OrderID = mOrder.OrderID
End Property
Private Sub txtjobname_GotFocus()
SelAllText txtjobname
End Sub

Private Sub txtjobname_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPO.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtjobname_LostFocus()
Caption = App.Title & " - " & mOrder.Customer.AddressInfo.Name & ":" & Me.txtjobname
End Sub

Private Sub txtLength_GotFocus()
SelAllText txtLength
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmbSlab.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtPO_Change()
    Dirty = True
End Sub

Private Sub txtPO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ucReqDate.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtQuant_GotFocus()
SelAllText txtQuant
End Sub

Private Sub txtQuant_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtLength.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtQuant_LostFocus()
With txtQuant
    If Not IsNumeric(txtQuant) Then
        .Text = ""
        .BackColor = vbToolTip
        .ForeColor = vbToolTipText
    Else
        .Text = CInt(txtQuant)
        .BackColor = vbWindowBackground
        .ForeColor = vbWindowText
    End If
End With
End Sub
Private Function fnverifyfields() As Boolean
fnverifyfields = True
End Function

Public Sub OpenOrder(JobID As Long)
Dim lintCounter As Integer
Dim llngBatchid As Long
Dim lclsline As Line
Dim lctrl As Control

mbolNewOrder = False

Set mOrder = New Order

mbolLoadingOrder = True

mOrder.Init JobID

RefreshLineDetailsGrid
RefreshOrderTotalsAndSpecialCharge
RefreshSavedOrderDetails

For lintCounter = 0 To (cmbCustomer.ListCount) - 1
    If cmbCustomer.ItemData(lintCounter) = mOrder.Customer.UniqueID Then
        cmbCustomer.ListIndex = lintCounter
        Exit For
    End If
Next

With mOrder
    For lintCounter = 1 To 5
        txtShipto(lintCounter) = .ShipTo(lintCounter)
    Next
    txtjobname = .GetProperty("vchJobName")
    txtPO = .PO
    ucReqDate.dtDate = .RequiredDate
    ucEntryDate.dtDate = .Entrydate
    txtTrackingCode = .Trackingcode
    For lintCounter = 0 To (cmbOrderType.ListCount - 1)
        If cmbOrderType.ItemData(lintCounter) = .JobType Then
            cmbOrderType.ListIndex = lintCounter
            Exit For
        End If
    Next
    OERepControl.SelectedRepID = .UserID
    chkCartoned = IIf(.Cartoned, vbChecked, vbUnchecked)
    chkSplined = IIf(.Splined, vbChecked, vbUnchecked)
    chkPadded = IIf(.Padded, vbChecked, vbUnchecked)
    chkGroupOrder = IIf(.GetProperty("tiGroup"), vbChecked, vbUnchecked)
    chkRush = IIf(.Rush, vbChecked, vbUnchecked)
    If IsNumeric(.GetProperty("iBatchID")) Then
        llngBatchid = CLng(.GetProperty("iBatchID"))
        If llngBatchid > 0 Then
            txtBatchID.Visible = True
            txtBatchID = "On Shipment " & llngBatchid
        End If
    End If
    lblPriceList = DataCenter.GetPriceListName(.PriceListID)
End With

mlngCurrentLine = 0

SelectTab 1
Dirty = False

If Not mOrder.Editable Then
    'Disable all controls
    Dim aCtrlType As String
    For Each lctrl In Me.Controls
        aCtrlType = TypeName(lctrl)
        If aCtrlType = "TextBox" Or aCtrlType = "CheckBox" Or aCtrlType = "ucDate" Then
            lctrl.Enabled = False
        ElseIf aCtrlType = "ComboBox" Then
            If Not lctrl = cmbPOList Then
                lctrl.Enabled = False
            End If
        End If
    Next
End If
ShowCustomerDataArea

mbolLoadingOrder = False

End Sub

Private Sub SelectTab(inTabIndex As Integer)
    If Not MainTab.Tabs(inTabIndex).Selected Then
        MainTab.Tabs(inTabIndex).Selected = True
    End If
    frmOrder(inTabIndex).ZOrder vbBringToFront
    frmOrder(inTabIndex).Enabled = True
    If inTabIndex = 1 Then
        frmOrder(2).Enabled = False
        SetFocusTo cmbCustomer
    Else
        SetFocusTo txtLength
        frmOrder(1).Enabled = False
    End If
End Sub

Private Sub SetFocusTo(inControl As Control)
    If inControl.Enabled Then
        inControl.SetFocus
    End If
End Sub


Public Function ValidateScreen() As Boolean
Dim lintCounter     As Integer
Dim lstrMessage     As String

If Not FieldsValid(lstrMessage) Then
    MsgBox lstrMessage, vbOKOnly Or vbExclamation, App.Title
    ValidateScreen = False
    Exit Function
End If

With mOrder
        .Customerid = cmbCustomer.ItemData(cmbCustomer.ListIndex)
        For lintCounter = 1 To 5
           .ShipTo(lintCounter) = txtShipto(lintCounter)
        Next
        .Cartoned = chkCartoned.Value = vbChecked
        .Splined = chkSplined.Value = vbChecked
        .Padded = chkPadded.Value = vbChecked
        .Rush = chkRush.Value = vbChecked
        .JobType = cmbOrderType.ItemData(cmbOrderType.ListIndex)
        .SetProperty "vchJobName", txtjobname
        .SetProperty "tiGroup", (chkGroupOrder.Value = vbChecked)
        .PO = txtPO
        .SetProperty "dtEntryDate", ucEntryDate.dtDate
        .RequiredDate = ucReqDate.dtDate
        .UserID = OERepControl.SelectedRepID
        .Trackingcode = txtTrackingCode
        Caption = .DescriptiveName
        lblOrderNumber = .OrderNumber
        'lblRefNo = .GetProperty("vchRefNumber")
        lblState = .StateDescription
End With
ValidateScreen = True
End Function

Private Sub RefreshSavedOrderDetails()
'// Refreshes some of the Order captions after an order has been saved
With mOrder
    Caption = .DescriptiveName
    lblOrderNumber = .OrderNumber
    'lblRefNo = .GetProperty("vchRefNumber")
    lblState = .StateDescription
End With
End Sub

Private Sub txtShipto_Change(Index As Integer)
    Dirty = True
End Sub

Private Sub txtShipto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index < 5 Then
            txtShipto(Index + 1).SetFocus
        Else
           chkCartoned.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub txtTrackingCode_Change()
    Dirty = True
End Sub

Private Function PromptBeforeContinuing(pstrFunction As String) As VbMsgBoxResult
Dim lResult As VbMsgBoxResult
    If Dirty Then
        lResult = MsgBox("Save Order before " & pstrFunction & "?", vbYesNoCancel, App.Title)
        If lResult = vbYes Then
            If Save = True Then
                PromptBeforeContinuing = vbYes
            Else
                PromptBeforeContinuing = vbCancel
            End If
        Else
            PromptBeforeContinuing = lResult
        End If
    Else
        PromptBeforeContinuing = vbYes
    End If
End Function


Private Sub txtTrackingCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MainTab.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub ucEntryDate_ValueChanged()
    Dirty = True
End Sub

Private Sub ucReqDate_ValueChanged()
    Dirty = True
End Sub
