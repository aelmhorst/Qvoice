VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarBarCode As String 'local copy
Private mvarDateScanned As Date 'local copy
Public Property Let DateScanned(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateScanned = 5
    mvarDateScanned = vData
End Property


Public Property Get DateScanned() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateScanned
    DateScanned = mvarDateScanned
End Property



Public Property Let BarCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BarCode = 5
    mvarBarCode = vData
End Property


Public Property Get BarCode() As String
Attribute BarCode.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BarCode
    BarCode = mvarBarCode
End Property



