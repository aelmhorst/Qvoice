Attribute VB_Name = "MainModule"
Option Explicit

'Public Strings
Public gstrSeperator    As String
Public gbolUniqueDocIDs As Boolean

'Public Enums
Public Enum eUserFunction
    NewJob = 1
    LoadJob = 2
End Enum

Public Enum OpenOrderReason
    ForEditing = 1
    ForPosting = 2
    ForCreatingPOs = 3
End Enum


    
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Integer

Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = -1

'Types
Public Type OrderType
    ID As Integer
    Name As String
    Caption As String
    CanPost As Boolean
    IsTemplate As Boolean
    CreateReference As Boolean
End Type



Public Function CreateReference(iOrderType As Integer) As Boolean
    CreateReference = GetOrderType(iOrderType).CreateReference
End Function

Public Function IsTemplate(iOrderType As Integer) As Boolean
    IsTemplate = GetOrderType(iOrderType).IsTemplate
End Function

Public Function GetOrderType(inOrderTypeID As Integer) As OrderType
    Dim lintMax         As Integer
    Dim lintCounter     As Integer
    Dim aOrderType      As OrderType
    
    lintMax = UBound(OrderTypes)
    For lintCounter = 1 To lintMax
        If OrderTypes(lintCounter).ID = inOrderTypeID Then
            aOrderType = OrderTypes(lintCounter)
            Exit For
        End If
    Next
    GetOrderType = aOrderType
End Function


Sub Main()
    Dim aStartTime As Double
    aStartTime = Timer
    
    Dim aShowSplash As Boolean
    aShowSplash = Settings.ShowSplash
    
    Dim aSplashScreen As frmSplash
    If aShowSplash Then
        Set aSplashScreen = New frmSplash
        With aSplashScreen
            .Show
            .Refresh
        End With
    End If
    
    Initialize
    If aShowSplash Then
        While Timer - aStartTime < 1
            DoEvents
        Wend
        Unload aSplashScreen
        Set aSplashScreen = Nothing
    End If
    
'    If LoginUser Then
'        If Globals.LoggedInUser.Role = Shop Then
'            '// TODO: Show the shop main window
'        Else
'            '// This is an office worker, show the office main screen
            Dim aMainWindow As frmMDI
            Set aMainWindow = New frmMDI
            Set Globals.MainDocumentWindow = aMainWindow
            Load aMainWindow
            aMainWindow.PostInitialize
            aMainWindow.Show
'        End If
'    End If
End Sub

'Private Function LoginUser() As Boolean
'    Dim aLogin As New frmLogin
'    aLogin.Show vbModal
'    LoginUser = aLogin.LoginSucceeded
'    Unload aLogin
'End Function


'********************************************************************************************
'   Private Sub Initialize()
'   Created By:     Andy Elmhorst
'   Purpose:        Sets up the global variables and connects to the database
'********************************************************************************************
Private Sub Initialize()
    Dim rs              As Recordset
    Dim lintCounter     As Integer
    gstrSeperator = GetSetting("QK", "General", "Seperator", "/")
    
    Set rs = DataCenter.LoadOrderTypes
    lintCounter = 1
    With rs
        ReDim OrderTypes(1 To .RecordCount)
        Do Until .EOF
            OrderTypes(lintCounter).ID = !iOrderType
            OrderTypes(lintCounter).Name = !vchOrderTypeDesc
            OrderTypes(lintCounter).Caption = !vchFormCaption & ""
            OrderTypes(lintCounter).CanPost = !tiPost
            OrderTypes(lintCounter).IsTemplate = !tiTemplate
            OrderTypes(lintCounter).CreateReference = !tiReference
            lintCounter = lintCounter + 1
            .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub

Public Function ToWords(plngNumber As Long) As String

Select Case plngNumber
    Case Is > 999999
        ToWords = ToWords(Int(plngNumber / 1000000)) & _
            " million, " & ToWords(plngNumber Mod 1000000)
    Case Is > 999
        ToWords = ToWords(Int(plngNumber / 1000)) & _
            " thousand, " & ToWords(plngNumber Mod 1000)
    Case Is > 99
        ToWords = ToWords(Int(plngNumber / 100)) & _
            " hundred and " & ToWords(plngNumber Mod 100)
    Case Is > 19
        If plngNumber > 79 And plngNumber < 90 Then
            ToWords = "eighty-" & ToWords(plngNumber Mod 10)
        ElseIf plngNumber < 60 And plngNumber > 49 Then
            ToWords = "fifty-" & _
                ToWords(plngNumber Mod 10)
        ElseIf plngNumber > 39 Then
            ToWords = ToWords(Int(plngNumber / 10)) & _
                "ty-" & ToWords(plngNumber Mod 10)
        ElseIf plngNumber > 29 Then
            ToWords = "thirty-" & _
                ToWords(plngNumber Mod 10)
        Else
            ToWords = "twenty-" & _
                ToWords(plngNumber Mod 10)
        End If
        
    Case Is = 1
        ToWords = "one"
    Case Is = 2
        ToWords = "two"
    Case Is = 3
         ToWords = "three"
    Case Is = 4
         ToWords = "four"
    Case Is = 5
         ToWords = "five"
    Case Is = 6
         ToWords = "six"
    Case Is = 7
         ToWords = "seven"
    Case Is = 8
         ToWords = "eight"
    Case Is = 9
         ToWords = "nine"
    Case Is = 10
         ToWords = "ten"
    Case Is = 11
         ToWords = "eleven"
    Case Is = 12
         ToWords = "twelve"
    Case Is = 13
         ToWords = "thirteen"
    Case Is = 15
         ToWords = "fifteen"
    Case Is < 1
    Case Else
         ToWords = ToWords(plngNumber - 10) & "teen"
End Select
End Function


Public Sub AutoMatch(cbo As ComboBox, KeyAscii As Integer)
    Dim sbuffer As String
    Dim lretval As Long
    
    sbuffer = Left$(cbo.Text, cbo.SelStart) & Chr(KeyAscii)
    lretval = SendMessage((cbo.hwnd), CB_FINDSTRING, -1, ByVal sbuffer)
    
    If lretval <> CB_ERR Then
        cbo.ListIndex = lretval
        cbo.Text = cbo.List(lretval)
        cbo.SelStart = Len(sbuffer)
        cbo.SelLength = Len(cbo.Text)
        KeyAscii = 0
    End If
End Sub

Public Sub HandleError(pstrErrorString As String, pbolFatal As Boolean)
    
    LogIt Err.Number & " - " & Err.Description & " : " & Err.Source & " : " & Err.LastDllError
    
    If Len(pstrErrorString) > 0 Then MsgBox pstrErrorString, vbCritical, App.Title
    
    If pbolFatal Then Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub LogIt(pstrLogString As String)
    Dim linthwnd As Integer
    linthwnd = FreeFile()

    Open App.Path & "\" & App.EXEName & ".log" For Append Shared As #linthwnd
    Print #linthwnd, FormatDateTime(Now), pstrLogString
    Close #linthwnd

End Sub

Public Sub SelAllText(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub





#If DBUG Then
    Public Sub PerfLog(inMessage As String, inIndent As Integer)
    Static linthwnd As Integer
    If linthwnd = 0 Then
        linthwnd = FreeFile()
        Open App.Path & "\" & App.EXEName & "_Perf.log" For Append As #linthwnd
    End If
    
    Print #linthwnd, Format$(Now, "hh:mm:ssss") & Space$(inIndent * 3) & inMessage
    
    End Sub
#End If

Public Function CleanDocumentName(in_Name As String, in_Prefix As String) As String
    Dim lintx As Integer
    Dim lNewStr As String
    
    'Convert to numeric only
    For lintx = 1 To Len(in_Name)
        If IsNumeric(Mid$(in_Name, lintx, 1)) Then
             lNewStr = lNewStr & Mid$(in_Name, lintx, 1)
        End If
    Next
    If Len(lNewStr) > 0 Then
        If Len(lNewStr) > 7 Then lNewStr = Left$(lNewStr, 7)
        lNewStr = Format$(CLng(lNewStr), "000000")
    End If
    CleanDocumentName = in_Prefix & lNewStr
End Function

Public Function CleanDBString(in_String As String, in_Delimiter As String)
    
    CleanDBString = Replace(in_String, in_Delimiter, String$(2, in_Delimiter), 1, Compare:=vbTextCompare)
    
End Function

Public Function GetFlicSettings(inScannerNumber As Integer) As FlicSettings
    Dim aSetting As String
    aSetting = Settings.GetBarCodeScannerSettings(inScannerNumber)
    If Len(aSetting) > 0 Then
        Dim aFlicSetting As FlicSettings
        Set aFlicSetting = New FlicSettings
        aFlicSetting.FromString (aSetting)
        Set GetFlicSettings = aFlicSetting
    End If
End Function




