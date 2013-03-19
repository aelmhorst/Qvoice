VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMainDoctor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Qvoice Data Doctor"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMainDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateIndexList 
      Caption         =   "Create Inde&x Def"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   600
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRepair 
      Caption         =   "Repair Database Indexes"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmMainDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCreateIndexList_Click()
    With dlgFiles
        .DialogTitle = "Select Qvoice Database"
        .Filter = "Access Databases (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With
     
    If Len(dlgFiles.FileName) > 0 Then
        CreateIndexDef dlgFiles.FileName
    End If
End Sub

Private Sub cmdRepair_Click()
    With dlgFiles
        .DialogTitle = "Select Qvoice Database"
        .Filter = "Access Databases (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With
     
    If Len(dlgFiles.FileName) > 0 Then
        RebuildIndexes dlgFiles.FileName
    End If
End Sub

Public Sub Status(inStatus As String)
    lblStatus.Caption = inStatus
    lblStatus.Refresh
End Sub
