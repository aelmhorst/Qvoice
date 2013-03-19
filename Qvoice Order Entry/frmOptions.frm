VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Specify Options for QK Order"
   ClientHeight    =   4305
   ClientLeft      =   2280
   ClientTop       =   3405
   ClientWidth     =   6585
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameScannerOptions 
      Caption         =   "Barcode Scanner 1 Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   6255
      Begin VB.ComboBox cmbPortNumber 
         Height          =   315
         ItemData        =   "frmOptions.frx":0442
         Left            =   240
         List            =   "frmOptions.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chkScanner1Enabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbScannerType 
         Height          =   315
         ItemData        =   "frmOptions.frx":0446
         Left            =   240
         List            =   "frmOptions.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblScanner 
         Caption         =   "Port"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblScanner 
         Caption         =   "Scanner Type"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frmGeneral 
      Caption         =   "Please specify the following options"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkShowSplash 
         Caption         =   "Show Splash Screen On Startup"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtSetting 
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   1
         TabIndex        =   0
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Seperator Key for Charges"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   510
         Width           =   1935
      End
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Select Case txtSetting(0).Text
        Case Is = "'"
            MsgBox "The apostrophe is reserved for text descriptions on the charge line" & _
            vbCrLf & "Please choose another character", vbOKOnly Or vbInformation, "Options"
            txtSetting(0) = gstrSeperator
            txtSetting(0).SetFocus
            Exit Sub
        Case Is = "-"
            MsgBox "The dash is reserved for charge quantities" & _
            vbCrLf & "Please choose another character", vbOKOnly Or vbInformation, "Options"
            txtSetting(0) = gstrSeperator
            txtSetting(0).SetFocus
            Exit Sub
    End Select
    If txtSetting(0).Tag = "Changed" And Len(Trim$(txtSetting(0))) <> 0 Then
        SaveSetting "QK", "General", "Seperator", txtSetting(0).Text
        gstrSeperator = txtSetting(0).Text
    End If
    
    Settings.ShowSplash = (chkShowSplash.Value = vbChecked)
    
    
    '// Save the flic scanner settings
    If cmbPortNumber.ListIndex > -1 Then
        Dim aFlicSetting As New FlicSettings
        aFlicSetting.Enabled = chkScanner1Enabled.Value
        aFlicSetting.PortNumber = cmbPortNumber.ListIndex + 1
        Settings.SaveBarCodeScannerSettings 1, aFlicSetting.ToString()
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtSetting(0) = gstrSeperator
    txtSetting(0).Tag = ""
    chkShowSplash.Value = IIf(Settings.ShowSplash, vbChecked, vbUnchecked)
    
    '// Initialize the barcode scanner options
    With cmbScannerType
        .AddItem "Flic Tethered / Batch Scanner"
        .ListIndex = 0
    End With
    
    Dim aCounter As Integer
    For aCounter = 1 To 10
        Me.cmbPortNumber.AddItem "Com Port " & aCounter
        Me.cmbPortNumber.ItemData(Me.cmbPortNumber.NewIndex) = aCounter
    Next
    
   
    Dim aFlicSetting As FlicSettings
    Set aFlicSetting = MainModule.GetFlicSettings(1)
    If Not aFlicSetting Is Nothing Then
        cmbPortNumber.ListIndex = aFlicSetting.PortNumber - 1
        chkScanner1Enabled.Value = IIf(aFlicSetting.Enabled, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
    End If
   
End Sub

            

Private Sub txtSetting_Change(Index As Integer)
txtSetting(Index).Tag = "Changed"
End Sub

Private Sub txtSetting_GotFocus(Index As Integer)
With txtSetting(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
