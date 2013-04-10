VERSION 5.00
Begin VB.Form frmColorLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laminate Lookup Wizard"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
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
   ScaleHeight     =   7515
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUpcharges 
      Height          =   1260
      Left            =   4680
      TabIndex        =   26
      Top             =   6000
      Width           =   3975
   End
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
         Top             =   3000
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
      Begin VB.ComboBox cmbPriceList 
         Height          =   360
         Left            =   4680
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   1080
         Width           =   3615
      End
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
    LaminateId As Long
    Code As String
    Description As String
    Brand As String
    UpchargeDesc As String
    Upcharge As Currency
End Type
Private mLamTypes() As LaminateType
Private mintLamCount As Integer

Private Sub cmdGo_Click()
FillResultsList txtSearch.Text
End Sub

Private Sub FillResultsList(pstrSearchText As String)
Dim rs As Recordset
Dim X As Integer

    Screen.MousePointer = MousePointerConstants.vbHourglass
    Set rs = DataCenter.GetColorMatches(pstrSearchText, cmbPriceList.ItemData(cmbPriceList.ListIndex))
    Screen.MousePointer = MousePointerConstants.vbDefault
 
    mintLamCount = rs.RecordCount
    X = 0
    
    ReDim mLamTypes(0 To mintLamCount)
    Me.lstResults.Clear
    
    With rs
    Do Until .EOF
        X = X + 1
        mLamTypes(X).LaminateId = !iLaminateID
        mLamTypes(X).Code = !vchLaminateCode
        mLamTypes(X).Description = !vchLaminateDesc
        mLamTypes(X).UpchargeDesc = !vchColorCodeDescription
        mLamTypes(X).Brand = !vchBrandDescription
        mLamTypes(X).Upcharge = !flJobUpcharge
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
        Me.lblName = .Description
        Me.lblUpcharge = .UpchargeDesc
    End With
    
    RefreshUpchargeDetails lintIndex
    
    
End Sub

Private Sub RefreshUpchargeDetails(lintIndex As Integer)
    lstUpcharges.Clear
    Dim rs As Recordset
    Set rs = DataCenter.GetUpChargeDetails(mLamTypes(lintIndex).LaminateId, cmbPriceList.ItemData(cmbPriceList.ListIndex))
    While Not rs.EOF
        lstUpcharges.AddItem rs!vchSlabTypeDescription & ";" & FormatCurrency(rs!mSlabUpCharge)
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    


End Sub

Private Sub Form_Load()
Dim lintSetting
    SelAllText txtSearch
      
    lintSetting = CInt(GetSetting(App.EXEName, "General", "DefaultColorSortKey", "0"))
    If lintSetting < 3 Then
        optSort(lintSetting) = True
    End If
    
    Dim rs As Recordset
    Set rs = DataCenter.GetPriceLists
    While Not rs.EOF
        cmbPriceList.AddItem rs!vchPriceListDesc
        cmbPriceList.ItemData(cmbPriceList.ListCount - 1) = rs!iPriceListID
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    cmbPriceList.ListIndex = 0
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


