Attribute VB_Name = "EFBAR128"

Private BarTextOut As String
Private BarTextIn As String
Private BarTextInA As String
Private BarTextInB As String
Private TempString As String
Private BarTempOut As String
Private BarCodeOut As String
Private Sum As Long
Private II As Integer
Private ThisChar As String
Private CharValue As Long
Private CheckSumValue As Integer
Private CheckSum As String
Private Subset As Integer
Private StartChar As String
Private Weighting As Integer
Private UCC As Integer

' Copyright 2000-2003 by Elfring Fonts Inc. All rights reserved. This code
' may not be modified or altered in any way.
' Modified 2/13/03 for Word?Excel/Access quote problem workaround
'  put duplicate quote bar code character in slot 203, have code here replace
'  a quote character (34) with character 203, elimintating weird mailmerge bugs

'Functions in this file:
' Bar128A(Text)     -> convert text to bar code 128 subset A
' Bar128Aucc(Text)  -> convert text to bar code 128 subset A, UCC/EAN
' Bar128B(Text)     -> convert text to bar code 128 subset B
' Bar128Bucc(Text)  -> convert text to bar code 128 subset B, UCC/EAN
' Bar128C(Text)     -> convert text to bar code 128 subset C
' Bar128Cucc(Text)  -> convert text to bar code 128 subset C, UCC/EAN

' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for subset A only. It adds the
' start character, scans and converts data, adds a checksum and a stop
' character. Note that lower case letters are interpreted as control
' characters in subset A!
'------------------------------------------------------
Public Function Bar128A(BarTextInA As String) As String
  Bar128A = Bar128AB(BarTextInA, 0)
End Function


' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for UCC/EAN, subset A only.
' It adds the two start characters, scans and converts data, adds a checksum
' and a stop character. Note that lower case letters are interpreted as control
' characters in subset A!
'---------------------------------------------------------
Public Function Bar128Aucc(BarTextInA As String) As String
  ' Add FNC1 to beginning of string
  TempString = Chr(172) & BarTextInA
  Bar128Aucc = Bar128AB(TempString, 0)
End Function


' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for subset B only. It adds the
' start character, scans and converts data, adds a checksum and a stop
' character.
'------------------------------------------------------
Public Function Bar128B(BarTextInB As String) As String
  Bar128B = Bar128AB(BarTextInB, 1)
End Function


' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for UCC/EAN, subset B only.
' It adds the two start characters, scans and converts data, adds a checksum
' and a stop character.
'---------------------------------------------------------
Public Function Bar128Bucc(BarTextInB As String) As String
  ' Add FNC1 to beginning of string
  TempString = Chr(172) & BarTextInB
  Bar128Bucc = Bar128AB(TempString, 1)
End Function


'-----------------------------------------------------------------------------
' Convert input string to bar code 128 A or B format, Pass Subset 0 = A, 1 = B
'-----------------------------------------------------------------------------
Public Function Bar128AB(BarTextIn As String, Subset As Integer) As String

' Initialize input and output strings
BarTextOut = ""
BarTextIn = RTrim(LTrim(BarTextIn))

' Set up for the subset we are in
If Subset = 0 Then
  Sum = 103
  StartChar = "{"
Else
  Sum = 104
  StartChar = "|"
End If

' Calculate the checksum, mod 103 and build output string
For II = 1 To Len(BarTextIn)
  'Find the ASCII value of the current character
  ThisChar = (Asc(Mid(BarTextIn, II, 1)))
  'Calculate the bar code 128 value
  If ThisChar < 127 Then
    CharValue = ThisChar - 32
  Else
    CharValue = ThisChar - 103
  End If
  'add this value to sum for checksum work
  Sum = Sum + (CharValue * II)

  'Now work on output string, no spaces in TrueType fonts, quotes replaced for Word mailmerge bug
  If Mid(BarTextIn, II, 1) = " " Then
    BarTextOut = BarTextOut & Chr(228)
  ElseIf Asc(Mid(BarTextIn, II, 1)) = 34 Then
    BarTextOut = BarTextOut & Chr(226)
  Else
    BarTextOut = BarTextOut & Mid(BarTextIn, II, 1)
  End If
Next II

' Find the remainder when Sum is divided by 103
CheckSumValue = (Sum Mod 103)
' Translate that value to an ASCII character
If CheckSumValue > 90 Then
  CheckSum = Chr(CheckSumValue + 103)
ElseIf CheckSumValue > 0 Then
  CheckSum = Chr(CheckSumValue + 32)
Else
  CheckSum = Chr(228)
End If

'Build ouput string, trailing space is for Windows rasterization bug
BarTempOut = StartChar & BarTextOut & CheckSum & "~ "

'Return the string
Bar128AB = BarTempOut
End Function


' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for subset C only. It adds the
' start character, throws away any non-numeric data, adds a leading zero if
' there aren't an even number of digits, scans and converts data into data
' pairs, adds a checksum and a stop character.
'-----------------------------------------------------
Public Function Bar128C(BarTextIn As String) As String
  Bar128C = Bar128SubsetC(BarTextIn, 0)
End Function

' This function converts a string into a format compatible with Elfring
' Fonts Inc bar codes. This conversion is for UCC/EAN subset C only. It adds
' the two start characters, throws away any non-numeric data, adds a leading
' zero if there aren't an even number of digits, scans and converts data into
' data pairs, adds a checksum and a stop character.
'--------------------------------------------------------
Public Function Bar128Cucc(BarTextIn As String) As String
  Bar128Cucc = Bar128SubsetC(BarTextIn, 1)
End Function

'---------------------------------------------------------------------------
' Convert input string to bar code 128 C format, Pass UCC 0 = no, 1 = yes
'---------------------------------------------------------------------------
Public Function Bar128SubsetC(BarTextIn As String, UCC As Integer) As String

' Initialize input and output strings
BarTextOut = ""
BarTextIn = RTrim(LTrim(BarTextIn))

' Throw away non-numeric data
TempString = ""
For I = 1 To Len(BarTextIn)
  If IsNumeric(Mid(BarTextIn, I, 1)) Then
    TempString = TempString & Mid(BarTextIn, I, 1)
  End If
Next I

' If not an even number of digits, add a leading 0
If (Len(TempString) Mod 2) = 1 Then
  TempString = "0" & TempString
End If

' If UCC = 0, then normal start, otherwise UCC/EAN start
If UCC = 0 Then
  Sum = 105
  StartChar = "}"
  Weighting = 1
Else
  Sum = 207
  StartChar = "}²"
  Weighting = 2
End If

' Calculate the checksum, mod 103 and build output string
For I = 1 To Len(TempString) Step 2
    'Break string into pairs of digits and get value
    CharValue = Mid(TempString, I, 2)
    'Multiply value times weighting and add to sum
    Sum = Sum + (CharValue * Weighting)
    Weighting = Weighting + 1

    'translate value to ASCII and save in BarTextOut
    If CharValue < 90 Then
      BarTextOut = BarTextOut & Chr(CharValue + 33)
    Else
      BarTextOut = BarTextOut & Chr(CharValue + 104)
    End If
Next I

' Find the remainder when Sum is divided by 103
CheckSumValue = (Sum Mod 103)
' Translate that value to an ASCII character
If CheckSumValue < 90 Then
  CheckSum = Chr(CheckSumValue + 33)
Else
  CheckSum = Chr(CheckSumValue + 104)
End If

'Build ouput string, trailing space for Windows rasterization bug
BarTempOut = StartChar & BarTextOut & CheckSum & "~ "

' Replace all quote characters with duplicate bar code character in slot 226
' fixing MS Word mailmerge bug
BarCodeOut = ""
For II = 1 To Len(BarTempOut)
  'Find the ASCII value of the current character
  ThisChar = (Asc(Mid(BarTempOut, II, 1)))
  If ThisChar = 34 Then
    BarCodeOut = BarCodeOut & Chr(226)
  Else
    BarCodeOut = BarCodeOut & Mid(BarTempOut, II, 1)
  End If
Next II

'Return the string
Bar128SubsetC = BarCodeOut
End Function

