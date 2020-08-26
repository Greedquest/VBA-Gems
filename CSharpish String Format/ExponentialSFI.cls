VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExponentialSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormatString(ByVal specifier As StringFormatSpecifier) As String
    Err.Raise 5                                  'TODO implement interface member
End Function

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormattedValue = GetFormattedValue(value, specifier)
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase(specifier.identifier) = "E"
End Function

Public Function GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
        
    If LenB(formatGroup) > 2 And specifier.Precision > 0 Then
        GetFormattedValue = Format$(value, "0." & String$(specifier.Precision - 1, Chr$(48)) & "#" & specifier.identifier & "-#")
    
    ElseIf LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
        Err.Raise ERR_FORMAT_EXCEPTION, ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
        
    Else
        GetFormattedValue = Format$(value, "0.00000#" & specifier.identifier & "-#") 'defaults to 6 decimals
        
    End If
    


End Function

'Should be static method really
Public Function ParseExponent(ByVal formattedValue As String) As Long
    ParseExponent = Mid(formattedValue, InStr(1, UCase$(formattedValue), "E-") + 1, Len(formattedValue) - InStr(1, UCase$(formattedValue), "E-"))
End Function

