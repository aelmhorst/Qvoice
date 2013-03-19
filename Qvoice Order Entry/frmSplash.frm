VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblDiagnostic 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblLicense 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   6135
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblCopyRight 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2000 - 2002 WebWide Services"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CanUnload As Boolean

Private Sub Form_Click()
    If CanUnload Then Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyRight.Caption = App.LegalCopyright
    lblLicense.Caption = App.Comments
    lblDiagnostic.Caption = "Database: " & Settings.ConnectionString
    Picture = LoadResPicture(101, vbResBitmap)
End Sub
Public Property Let EnableUnload(pbolEnable As Boolean)
    CanUnload = pbolEnable
    lblDiagnostic.Visible = pbolEnable
End Property





