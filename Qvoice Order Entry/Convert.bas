Attribute VB_Name = "Convert"
Option Explicit

Public Function BoolToString(in_Bool As Boolean)
    BoolToString = IIf(in_Bool, "TRUE", "FALSE")
End Function

Public Function StringToBool(inValue As String)
    StringToBool = (UCase$(Trim$(inValue)) = "TRUE")
End Function
