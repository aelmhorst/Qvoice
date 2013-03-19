VERSION 5.00
Begin VB.UserControl ucDate 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   405
   ScaleWidth      =   2325
   Begin VB.TextBox txtDate 
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "ucDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mbolIsInEdit As Boolean
Private mDate As Date
Private mChanged As Boolean
Private mNextControl As Control

Event ValueChanged()



Public Property Set NextControl(in_NextControl As Control)
    Set mNextControl = in_NextControl
End Property

Public Property Get dtDate() As Date
    dtDate = mDate
End Property
Public Property Let dtDate(newDate As Date)
    mDate = newDate
    txtDate.Text = Format$(newDate, "MMM DD, YYYY")
End Property

Public Function isValid() As Boolean
Dim mbolvalid As Boolean
mbolvalid = False

If Len(txtDate) > 0 Then
    If IsDate(txtDate) Then mbolvalid = True
End If
isValid = mbolvalid

End Function

Public Property Let Enabled(inEnabled As Boolean)
    txtDate.Enabled = inEnabled
End Property


Private Sub txtDate_GotFocus()
    If mbolIsInEdit Then
        mbolIsInEdit = False
    Else
        mChanged = False
        txtDate.Text = FormatDateTime(mDate, vbShortDate)
        SelAllText txtDate
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
Dim lintAdd         As Integer

If KeyAscii = vbKeyReturn And Not mNextControl Is Nothing Then
    mNextControl.SetFocus
    KeyAscii = 0
    Exit Sub
End If

Select Case KeyAscii
    Case Is = vbKeyAdd - 64
        lintAdd = 1
    Case Is = vbKeySubtract - 64
        lintAdd = -1
'    Case Is = vbKeyMultiply - 64
'        lintAdd = 7
'    Case Is = vbKeyDivide - 64
'        lintAdd = -7
End Select

If lintAdd <> 0 And isValid Then
    mChanged = True
    mDate = DateAdd("d", lintAdd, mDate)
    txtDate.Text = FormatDateTime(mDate, vbShortDate)
    KeyAscii = 0
End If

End Sub

Private Sub txtDate_LostFocus()
If Not IsDate(txtDate.Text) Then
    mbolIsInEdit = True
    txtDate.SetFocus
    SelAllText txtDate
    Beep
Else
    txtDate.Text = Format$(txtDate.Text, "MMM DD, YYYY")
    If mDate <> CDate(txtDate) Then
        RaiseEvent ValueChanged
        mDate = CDate(txtDate)
    ElseIf mChanged Then
        RaiseEvent ValueChanged
    End If
    mChanged = False
End If

End Sub
Private Sub UserControl_Initialize()
    mDate = CDate(Format$(Now(), "MM/DD/YYYY"))
    txtDate.Text = Format$(mDate, "MMM DD, YYYY")
End Sub


Private Sub UserControl_Resize()
    txtDate.Move 0, 0, Width, Height
End Sub
