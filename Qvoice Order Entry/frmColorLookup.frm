VERSION 5.00
Begin VB.Form frmColorLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laminate Lookup Wizard"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmColorLookup 
      Caption         =   "Laminate Details"
      Height          =   4095
      Left            =   4680
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
      Begin VB.Label Label7 
         Caption         =   "Upcharges:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblSESPrice 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "SES"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblLaminatePrice 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Laminate"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblSlabPrice 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Slab"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblUpcharge 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Upcharge Code"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblBrand 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Brand"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblColorCode 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Color Code"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ListBox lstResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laminate Lookup"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.OptionButton optSort 
         Caption         =   "&Brand"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optSort 
         Caption         =   "&Name"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optSort 
         Caption         =   "&Code"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Sort By"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Enter part of a color name or code and click ""Go"""
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "frmColorLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type LaminateType
    Code As String
    Description As String
    Brand As String
    UpchargeDesc As String
    SlabUpcharge As Currency
    LaminateUpcharge As Currency
    SESUpcharge As Currency
End Type
Private mLamTypes() As LaminateType
Private mintLamCount As Integer

Private Sub cmdGo_Click()
FillResultsList txtSearch.Text
End Sub

Private Sub FillResultsList(pstrSearchText As String)
Dim rs As Recordset
Dim X As Integer
Set rs = DataCenter.GetColorMatches(pstrSearchText, "qklgColorMatchesEx")

mintLamCount = rs.RecordCount
X = 0

ReDim mLamTypes(0 To mintLamCount)
Me.lstResults.Clear

With rs
    Do Until .EOF
        X = X + 1
        mLamTypes(X).Code = !vchLaminateCode
        mLamTypes(X).Description = !vchLaminateDesc
        mLamTypes(X).UpchargeDesc = !vchColorCodeDesc
        mLamTypes(X).Brand = !vchBrandDescription
        mLamTypes(X).SlabUpcharge = !mSlabUpCharge
        mLamTypes(X).LaminateUpcharge = !mLaminateUpcharge
        mLamTypes(X).SESUpcharge = !mSquareEdgeUpcharge
        Select Case True
        Case optSort(0)
            lstResults.AddItem mLamTypes(X).Code & " - " & mLamTypes(X).Description
        Case optSort(1)
            lstResults.AddItem mLamTypes(X).Description & " (" & mLamTypes(X).Code & ")"
        Case optSort(2)
            lstResults.AddItem mLamTypes(X).Brand & " - " & mLamTypes(X).Description
        End Select
        lstResults.ItemData(lstResults.NewIndex) = X
        .MoveNext
    Loop
    

    
End With
End Sub

Sub FillLaminateDetails(lintIndex As Integer)

With mLamTypes(lintIndex)
    Me.lblBrand = .Brand
    Me.lblColorCode = .Code
    Me.lblLaminatePrice = FormatCurrency(.LaminateUpcharge)
    Me.lblName = .Description
    Me.lblSESPrice = FormatCurrency(.SESUpcharge)
    Me.lblSlabPrice = FormatCurrency(.SlabUpcharge)
    Me.lblUpcharge = .UpchargeDesc

End With
End Sub


Private Sub Form_Load()
Dim lintSetting
    SelAllText txtSearch
      
    lintSetting = CInt(GetSetting(App.EXEName, "General", "DefaultColorSortKey", "0"))
    If lintSetting < 3 Then
        optSort(lintSetting) = True
    End If
      
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lstrsetting
    Select Case True
        Case optSort(0)
            lstrsetting = 0
        Case optSort(1)
            lstrsetting = 1
        Case optSort(2)
            lstrsetting = 2
    End Select
SaveSetting App.EXEName, "General", "DefaultColorSortKey", lstrsetting

End Sub

Private Sub lstResults_Click()
    FillLaminateDetails CInt(lstResults.ItemData(lstResults.ListIndex))
End Sub
