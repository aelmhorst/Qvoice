VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2430
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5130
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1435.724
   ScaleMode       =   0  'User
   ScaleWidth      =   4816.792
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   360
      Picture         =   "frmLogin.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   840
      Width           =   480
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2490
      TabIndex        =   1
      Top             =   735
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2490
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1125
      Width           =   2325
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Please enter your username and password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   1305
      TabIndex        =   2
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim aCurrentUser As CurrentUser
    Set aCurrentUser = DataCenter.GetUser(Me.txtUserName.Text, Me.txtPassword.Text)
    If aCurrentUser Is Nothing Then
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    Else
        Set Globals.LoggedInUser = aCurrentUser
        LoginSucceeded = True
        Me.Hide
    End If
End Sub
