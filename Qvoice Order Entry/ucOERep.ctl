VERSION 5.00
Begin VB.UserControl ucOERep 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ScaleHeight     =   405
   ScaleWidth      =   5625
   Begin VB.ComboBox cmbOERep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "ucOERep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event KeyPress(KeyAscii As Integer)
Private mREPID As Long



Private Sub cmbOERep_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        RaiseEvent KeyPress(KeyAscii)
    End If
End Sub


Private Sub UserControl_Resize()
    cmbOERep.Move 0, 0, Width
End Sub


Private Sub UserControl_Show()
    'Initialize the OeRep DropDown
    Dim rs As Recordset
    If Ambient.UserMode Then
        Set rs = DataCenter.GetOEReps
        With cmbOERep
            Do Until rs.EOF
                .AddItem rs!vchFirstName & " " & rs!vchLastName & Space$(5) & "x" & rs.Fields.Item(Constants.O_REP_EXTENSION).Value
                .ItemData(.NewIndex) = rs!iRepID
                rs.MoveNext
            Loop
        End With
        If mREPID > 0 Then
            SelectedRepID = mREPID
        Else
            Dim aRepID As Long
            aRepID = CLng(GetSetting(App.EXEName, "General", "DefaultOERepID", "0"))
            If aRepID > (cmbOERep.ListCount - 1) Then aRepID = 0
            cmbOERep.ListIndex = aRepID
        End If
    Else
        cmbOERep.AddItem "Example User" & Space$(5) & "x100"
        cmbOERep.ListIndex = 0
    End If
End Sub

Public Property Get SelectedRepID() As Long
    If cmbOERep.ListIndex = -1 Then
        SelectedRepID = 0
    Else
        SelectedRepID = cmbOERep.ItemData(cmbOERep.ListIndex)
    End If
End Property

Public Property Let SelectedRepID(inID As Long)
    Dim aCounter As Integer
    If cmbOERep.ListCount > 0 Then
        For aCounter = 0 To (cmbOERep.ListCount - 1)
            If cmbOERep.ItemData(aCounter) = inID Then
                cmbOERep.ListIndex = aCounter
                Exit For
            End If
        Next
        mREPID = inID
    Else
        mREPID = inID
    End If
End Property

Public Property Get SelectedRepName() As String
    Dim aParts() As String
    aParts = Split(cmbOERep.Text, "x")
    SelectedRepName = Trim$(aParts(0))
End Property

Public Property Get SelectedRepExtension() As String
    Dim aParts() As String
    aParts = Split(cmbOERep.Text, "x")
    SelectedRepExtension = Trim$(aParts(1))
End Property

Private Sub UserControl_Terminate()
    SaveSetting App.EXEName, "General", "DefaultOERepID", cmbOERep.ListIndex
End Sub
