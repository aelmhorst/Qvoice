VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NumberToCode128(inVal As Long, inPrefix As String, inFormatString As String) As String
    Dim aString As New CString
    aString.Append inPrefix
    aString.Append Format$(inVal, inFormatString)
    NumberToCode128 = ToCode128A(aString)
End Function

Private Function ToCode128A(inString As CString) As String
    Dim aCounter As Integer
    Dim aChecksum As Long
    Dim aVal As Integer
    
    aChecksum = 103 '// Barcode 128 Subset A starting checksum
    '
    ' Scan the string and add character value times position
    '
    For aCounter = 1 To inString.Length
        ' Get the numeric value of the character and subtract
        ' 32 to shift (the space character, ASCII value 32, has
        ' a numeric value of 0 as far as Code 128 is concerned)
        aVal = inString.Char(aCounter)
        
        If aVal < 127 Then
            aVal = aVal - 32
        Else
            aVal = aVal - 70 '103
        End If
        
        aChecksum = aChecksum + (aVal * aCounter)
        
        '// Now work on output string for Elfring Code 128 font support
        If aVal = 0 Then
            inString.Char(aCounter) = 228
        ElseIf aVal = 2 Then
            inString.Char(aCounter) = 226
        End If
    Next aCounter
    '
    ' Calculate the Modulo 103 checksum
    '
    aChecksum = aChecksum Mod 103
    '
    ' Now convert this number to a character.  This conversion
    ' takes into account the particular mapping of the font
    ' being used (this example is for the font published by
    ' Elfring Software.
    '
    
    If aChecksum > 90 Then
        aChecksum = aChecksum + 70 '103
    ElseIf aChecksum > 0 Then
        aChecksum = aChecksum + 32
    Else
        aChecksum = 174
    End If
    
    inString.Append Chr$(aChecksum)
    '
    ' Now format the final output string: start character,
    ' data, check character, and stop character
    '
    ToCode128A = "{" & inString.Value & "~"
End Function

