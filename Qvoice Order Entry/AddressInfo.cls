VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public PreAddressInfo As String
Public Name As String
Public Address1 As String
Public Address2 As String
Public Phone As String
Public City As String
Public State As String
Public Zip As String


Public Property Get Line1() As String
    Line1 = ContactInfo
End Property
Public Property Get Line2() As String
    Line2 = Name
End Property
Public Property Get Line3() As String
    Line3 = Address1
End Property
Public Property Get Line4() As String
    If Len(Address2) > 0 Then
        Line4 = Address2
    Else
        Line4 = BuildLastLine
    End If
End Property
Public Property Get Line5() As String
    If Len(Address2) > 0 Then
        Line5 = BuildLastLine
    Else
        Line5 = ""
    End If
End Property

Public Function GetAddressLine(in_Line As Integer)
    Select Case in_Line
        Case 1: GetAddressLine = Line1
        Case 2: GetAddressLine = Line2
        Case 3: GetAddressLine = Line3
        Case 4: GetAddressLine = Line4
        Case 5: GetAddressLine = Line5
    End Select
End Function

Private Function BuildLastLine() As String
    BuildLastLine = City & ", " & State & "  " & Zip
End Function



