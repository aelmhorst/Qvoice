VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mbolIsInEdit As Boolean
Private WithEvents DateControl As TextBox
Attribute DateControl.VB_VarHelpID = -1

Public Sub Init(TextObject As TextBox)
Set DateControl = TextObject
If Not IsDate(DateControl.Text) Then
    DateControl.Text = Format$(Now, "MMM DD, YYYY")
End If
End Sub

Private Sub DateControl_GotFocus()
If mbolIsInEdit Then
    mbolIsInEdit = False
Else
    If Not IsDate(DateControl.Text) Then
        DateControl.Text = FormatDateTime(Now(), vbShortDate)
    Else
        DateControl.Text = FormatDateTime(CDate(DateControl.Text), vbShortDate)
    End If
End If
End Sub

Private Sub DateControl_KeyPress(KeyAscii As Integer)
Dim lintAdd         As Integer

Select Case KeyAscii
    Case Is = vbKeyAdd - 64
        lintAdd = 1
    Case Is = vbKeySubtract - 64
        lintAdd = -1
    Case Is = vbKeyReturn
        lintAdd = 0
        DateControl.Text = FormatDateTime(Now(), vbShortDate)
End Select

If lintAdd <> 0 And IsDate(DateControl.Text) Then
    DateControl.Text = FormatDateTime(DateAdd("d", lintAdd, CDate(DateControl.Text)), vbShortDate)
    KeyAscii = 0
End If

End Sub

Private Sub DateControl_LostFocus()
If Not IsDate(DateControl.Text) Then
    'DateControl.Text = Format$(Now, "mm/dd/yyyy")
    mbolIsInEdit = True
    DateControl.SetFocus
    SelAllText DateControl
    Beep
Else
    DateControl.Text = Format$(DateControl.Text, "MMM DD, YYYY")
End If
End Sub
