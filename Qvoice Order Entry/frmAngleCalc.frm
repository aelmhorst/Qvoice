VERSION 5.00
Begin VB.Form frmAngleCalc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Angle Calculator"
   ClientHeight    =   5385
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAngleCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton commCalcoddangle 
      Caption         =   "&Calculate"
      Height          =   372
      Left            =   1440
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton commClearangle 
      Caption         =   "C&lear"
      Height          =   372
      Left            =   3360
      TabIndex        =   5
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Finished"
      Height          =   372
      Left            =   5280
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   852
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   852
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   4080
      TabIndex        =   3
      Top             =   3480
      Width           =   852
   End
   Begin VB.Label lblAngleA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblAngleA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   1920
      X2              =   6720
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   1200
      X2              =   1920
      Y1              =   4560
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderStyle     =   2  'Dash
      X1              =   6720
      X2              =   1200
      Y1              =   1920
      Y2              =   4560
   End
   Begin VB.Label lblPrompt 
      Caption         =   $"frmAngleCalc.frx":030A
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance left from corner"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance right from corner"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance Across Diagonal"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblAngleA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmAngleCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub commClearangle_Click()
    ClearAllFields
End Sub
Private Sub ClearAllFields()
Dim l As Integer
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    For l = 0 To 2
        lblAngleA(l).Caption = ""
    Next
End Sub

Private Sub Text1_gotfocus()
SelAllText Text1
End Sub
Private Sub Text2_gotfocus()
SelAllText Text2
End Sub
Private Sub Text3_gotfocus()
SelAllText Text3
End Sub
Private Sub commCalcoddangle_Click()
    Dim Angle1, Angle2, Angle3 As Double
    
    On Error GoTo errhandler:
    
    If Not (IsNumeric(Text1.Text) Or IsNumeric(Text2.Text) Or IsNumeric(Text3.Text)) Then
        MsgBox "Please enter real numbers only.", vbInformation, App.Title
        Exit Sub
    End If
    
    Angle1 = CDbl(Text1.Text)
    Angle2 = CDbl(Text2.Text)
    Angle3 = CDbl(Text3.Text)
    
    'Warning: funky code
    '(This was written in my junior days and I'm to lazy to fix it)
    lblAngleA(0).Caption = CalculateResult(Angle1, Angle2, Angle3)
    lblAngleA(1).Caption = CalculateResult(Angle3, Angle1, Angle2)
    lblAngleA(2).Caption = CalculateResult(Angle2, Angle3, Angle1)
    
    Exit Sub
    
errhandler:
        MsgBox "Invalid Values Specified", vbExclamation, App.Title
End Sub

Private Function CalculateResult(ByVal Angle1 As Double, ByVal Angle2 As Double, ByVal Diagonal As Double) As String
    Dim OddAngle As Double
    'Warning: funky code

    OddAngle = Angle1 ^ 2 + Angle2 ^ 2 - Diagonal ^ 2
    OddAngle = OddAngle / 2 / Angle1 / Angle2
    OddAngle = Atn(-OddAngle / Sqr(-OddAngle * OddAngle + 1)) + 2 * Atn(1)
    OddAngle = OddAngle * 180 / 3.1415926
    OddAngle = Round(OddAngle, 3)
    CalculateResult = OddAngle & "°"

End Function

Private Sub cmdDone_Click()
    ClearAllFields
    Unload Me
End Sub





