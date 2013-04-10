Attribute VB_Name = "Settings"
Option Explicit


'// Declares
Private Declare Function GetPrivateProfileString _
    Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString _
    Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
   
'// Strings
Private mstrConnect            As String



'//----------------------------------------------------------------------------------------------
'// Settings Class
'//----------------------------------------------------------------------------------------------
Public Property Get ShowSplash() As Boolean
    ShowSplash = Convert.StringToBool(GetIniSetting("General", "ShowSplash", "TRUE"))
End Property

Public Property Let ShowSplash(Value As Boolean)
    SaveIniSetting "General", "ShowSplash", Convert.BoolToString(Value)
End Property

Public Function GetBarCodeScannerSettings(inScannerNumber As Integer) As String
    GetBarCodeScannerSettings = GetIniSetting("BarCodeScanners", CStr(inScannerNumber), "")
End Function

Public Sub SaveBarCodeScannerSettings(inScannerNumber As Integer, inSettings As String)
    SaveIniSetting "BarCodeScanners", CStr(inScannerNumber), inSettings
End Sub

Public Property Get ConnectionString() As String
    If Len(mstrConnect) = 0 Then
        mstrConnect = GetIniSetting("General", "Connect", "DSN=QKORDER;")
    End If
    ConnectionString = mstrConnect
End Property

Public Function GetReportPath(inKey As String) As String
    GetReportPath = GetIniSetting("reports", inKey, "")
End Function

Public Function GetPrinter(inKey As String) As String
    GetPrinter = GetIniSetting("printers", inKey, "")
End Function

'********************************************************************************************
'   private Function SaveIniSetting(Section As String, Key As String,Value as String) As String
'   Created By:     Andy Elmhorst
'   Purpose:        Writes a setting to the Ini file for this application
'********************************************************************************************
Private Function SaveIniSetting(Section As String, Key As String, Value As String) As String
    Dim lstrworkstring As String
    lstrworkstring = Space$(255)
    Call WritePrivateProfileString(Section, Key, Value, App.Path & "\qkoe.INI")
End Function

'********************************************************************************************
'   private Function GetIniSetting(Section As String, Key As String, Default As String) As String
'   Created By:     Andy Elmhorst
'   Purpose:        Retrieves a string from the Ini file for this application
'********************************************************************************************
Private Function GetIniSetting(Section As String, Key As String, Default As String) As String
    Dim lstrworkstring As String
    Dim lngNsize As Long
    
    lstrworkstring = Space$(255)
    lngNsize = Len(lstrworkstring)
    
    Call GetPrivateProfileString(Section, Key, Default, lstrworkstring, lngNsize, App.Path & "\qkoe.ini")
    
    GetIniSetting = Left$(lstrworkstring, Len(Trim$(lstrworkstring)) - 1)
End Function



