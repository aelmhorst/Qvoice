VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMainDoctor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Qvoice Database Updater"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   Icon            =   "frmMainDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   1440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Query Defs as XML File"
      FileName        =   "QueryDefs.xml"
   End
   Begin VB.CommandButton cmdGetUpdate 
      Caption         =   "GetQueryDefs"
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearDb 
      Caption         =   "Zap Db"
      Height          =   735
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   600
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRepair 
      Caption         =   "Update Database"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
End
Attribute VB_Name = "frmMainDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClearDb_Click()
    With dlgFiles
        .DialogTitle = "Select Qvoice Database To Clear Out"
        .Filter = "Access Databases (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With
     
    If Len(dlgFiles.FileName) > 0 Then
        If MsgBox("Are you sure you want to permanently delete this data", vbYesNo, App.Title) = vbYes Then
            ModZap.ZapDb dlgFiles.FileName
            MsgBox "Database has been zapped", vbOKOnly, App.Title
        End If
    End If
    Unload Me
End Sub

Private Sub cmdGetUpdate_Click()
        With dlgFiles
        .DialogTitle = "Select Qvoice Database"
        .Filter = "Access Databases (*.mdb)|*.mdb"
        .FilterIndex = 1
        .ShowOpen
    End With
     
    If Len(dlgFiles.FileName) > 0 Then
        With dlgSave
            .DialogTitle = "Save Output As"
            .Filter = "Xml Files (*.xml|*.xml"
            .FilterIndex = 1
            .ShowSave
        End With
        If Len(dlgSave.FileName) > 0 Then
            QueryDefSQLCreator.CreateUpdateSQL dlgFiles.FileName, #9/1/2003#, dlgSave.FileName
        End If
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
        ModData.UpdateDatabase dlgFiles.FileName
    End If
    Unload Me
End Sub

Public Sub Status(inStatus As String)
    lblStatus.Caption = inStatus
    lblStatus.Refresh
End Sub
