VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlicSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FlicSignalLevel
    flicBeep = 1
    flicFlash = 2
    flicFlashAndBeep = 3
    flicDoubleBeep = 4
End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Const READ_TIMEOUT = 1

Private WithEvents mScannerPort As MSComm
Attribute mScannerPort.VB_VarHelpID = -1
Private mReceivedItems As ScanItems
Private mLastEventTimeStamp As Single
Private mPortNumber As Integer

'// (CONST) strings used for communication
Private CODE_SOH As String
Private CODE_STX As String
Private CODE_ETX As String
Private CODE_ACK As String
Private CODE_CR As String


Public Event OnScansReceived(inScanItems As ScanItems, ByRef outIsHandled As Boolean)
Public Event OnRead(inInfo As String)
Public Event OnWrite(inInfo As String)

Private Sub Class_Initialize()
    CODE_SOH = Chr$(1)
    CODE_STX = Chr$(2)
    CODE_ETX = Chr$(3)
    CODE_ACK = Chr$(6)
    CODE_CR = Chr$(13)
    
    Set mReceivedItems = New ScanItems
End Sub

Public Sub FlicSession_cctr(inComm As MSComm, inSettings As FlicSettings)
    Set mScannerPort = inComm
    mPortNumber = inSettings.PortNumber
End Sub

Public Sub Connect()
     With mScannerPort
        If .PortOpen Then Exit Sub
        .CommPort = mPortNumber
        .Settings = "4800,N,8,2"
        .DTREnable = True '// TODO: Try setting this to false
        .RThreshold = 1
        ' Tell the control to read entire buffer when Input is used.
        .InputLen = 0
        ' Open the port.
        .PortOpen = True
        
    End With
End Sub

Private Sub WakeIfSleeping()
    If Timer - mLastEventTimeStamp > 2 Then
        Wake
    End If
End Sub

Public Sub Disconnect()
    With mScannerPort
        If .PortOpen Then
            .PortOpen = False
        End If
    End With
End Sub

Private Sub Wake()
    ' Send the attention command to the modem.
    WriteLine Chr$(1) ' Ensure that
    Sleep 1000
    WriteLine "W" & vbCr
   
    Dim aResponse As String
    aResponse = ReadLine(vbLf)
    RaiseEvent OnRead(aResponse)

    aResponse = ReadLine(vbLf)
    RaiseEvent OnRead(aResponse)
End Sub

Public Sub Signal(inType As FlicSignalLevel)
    If inType = flicDoubleBeep Then
        Dim aCommand As String
        aCommand = CODE_SOH & "T3" & CODE_CR
        SendAckedCommand aCommand
        SendAckedCommand aCommand
    Else
        SendAckedCommand CODE_SOH & "T" & CInt(inType) & CODE_CR
    End If
End Sub

Private Sub SendAckedCommand(inCommand As String)
    WakeIfSleeping
    WriteLine inCommand
    Dim aResponse As String
    aResponse = ReadLine(CODE_ACK)
    RaiseEvent OnRead(aResponse)
End Sub

Private Sub Class_Terminate()
    If Not mScannerPort Is Nothing Then
        Disconnect
        Set mScannerPort = Nothing
    End If
End Sub

Private Sub mScannerPort_OnComm()
    ReceiveScans
End Sub

Private Sub ReceiveScans()
    Dim aResponse As String
    Do
        aResponse = ReadLine(vbLf)
        If Len(aResponse) > 0 Then
            If aResponse <> CODE_ACK Then
                WriteLine CODE_ACK
            End If
            If Left$(aResponse, 1) = CODE_STX Then
                '// this is a bar code
                mReceivedItems.Add Mid$(aResponse, 2, Len(aResponse) - 3), Now
            Else
                RaiseEvent OnRead(aResponse)
            End If
        End If
    Loop While Len(aResponse) > 0
    If mReceivedItems.Count > 0 Then
        Dim aIsHandled As Boolean
        RaiseEvent OnScansReceived(mReceivedItems, aIsHandled)
        If aIsHandled Then
            While mReceivedItems.Count > 0
                mReceivedItems.Remove 1
            Wend
        End If
    End If
End Sub

Private Function SendAck()
    mScannerPort.Output = CODE_ACK
End Function

Private Sub WriteLine(inLine As String)
    RaiseEvent OnWrite(inLine)
    mScannerPort.Output = inLine
    mLastEventTimeStamp = Timer
End Sub


Private Function ReadLine(inEndChar As String) As String
    '// TODO: GetCSTRING class
    Dim aReceivedString As String
    Dim aStartOfRead As Single
    Dim aIsReading As Boolean
    aIsReading = True
    aStartOfRead = Timer
    
    While aIsReading
        aReceivedString = aReceivedString & mScannerPort.Input
        If (InStr(1, aReceivedString, inEndChar)) Then
            aIsReading = False
        ElseIf Timer - aStartOfRead > READ_TIMEOUT Then
            aIsReading = False
        End If
    Wend
    ReadLine = aReceivedString
    mLastEventTimeStamp = Timer
End Function
