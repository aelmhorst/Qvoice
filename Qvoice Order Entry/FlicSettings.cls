VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlicSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mPortNumber As Integer
Private mEnabled As Boolean

Public Sub FromString(inString As String)
    Dim aSetting As Variant
    aSetting = Split(inString, ",")
    mEnabled = aSetting(0)
    mPortNumber = aSetting(1)
End Sub

Public Function ToString() As String
    ToString = Join(Array(mEnabled, mPortNumber), ",")
End Function


Public Property Get PortNumber() As Integer
    PortNumber = mPortNumber
End Property

Public Property Let PortNumber(inNumber As Integer)
    mPortNumber = inNumber
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(inEnabled As Boolean)
    mEnabled = inEnabled
End Property
