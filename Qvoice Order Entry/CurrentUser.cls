VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrentUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum Role
    Office = 1
    Shop = 2
    Admin = 3
End Enum

Private mUserID As Integer
Private mFirstName As String
Private mLastName As String
Private mRole As Role


Public Sub Init( _
    inUserID As Integer, _
    inFirstName As String, _
    inLastName As String, _
    inRole As Role _
   )
    
    mUserID = inUserID
    mFirstName = inFirstName
    mLastName = inLastName
    mRole = inRole
    
End Sub

Public Property Get FirstName() As String
    FirstName = mFirstName
End Property

Public Property Get LastName() As String
    LastName = mLastName
End Property

Public Property Get UserID() As Integer
    UserID = mUserID
End Property

Public Property Get Role() As Role
    Role = mRole
End Property

