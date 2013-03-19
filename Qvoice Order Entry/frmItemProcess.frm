VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemProcess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing Scan Items"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar mProgress 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmItemProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Property Let Max(inMax As Integer)
    mProgress.Max = inMax
End Property

Public Property Let Value(inValue As Integer)
    mProgress.Value = inValue
    Me.Refresh
End Property

Public Property Let ProgressCaption(inCaption As String)
    Me.Caption = inCaption
    Me.Refresh
End Property

