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
    Err.Raise 5                                  'TODO implement interface member
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    Err.Raise 5                                  'TODO implement interface member
End Function

Public Function GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String

    '    Dim result As String
    '    Dim exponentialNotation As String
    '    Dim power As Integer
    '    Dim exponentialFormat As New ExponentialSFI
    '    Dim fixedPointFormat As New FixedPointSFI
    '    Dim decimalFormat As New DecimalSFI
    '
    '    Dim formatSpecifier As New StringFormatSpecifier
    '    formatSpecifier.Alignment = specifier.Alignment
    '    formatSpecifier.CustomSpecifier = specifier.CustomSpecifier
    '
    '    If StringMatchesAny(TypeName(value), "Integer", "Long") Then
    '
    '        formatSpecifier.identifier = IIf(specifier.identifier = "G", "D", "d")
    '        result = decimalFormat.GetFormattedValue(value, formatSpecifier)
    '
    '    ElseIf TypeName(value) = "Double" Then
    '
    '        formatSpecifier.identifier = IIf(specifier.identifier = "G", "E", "e")
    '        exponentialNotation = exponentialFormat.GetFormattedValue(value, formatSpecifier)
    '        power = exponentialFormat.GetPower(exponentialNotation)
    '
    '        If power > -5 And Abs(power) < specifier.Precision Then
    '
    '            formatSpecifier.identifier = IIf(specifier.identifier = "G", "F", "f")
    '            result = fixedPointFormat.GetFormattedValue(value, formatSpecifier)
    '
    '        Else
    '
    '            result = exponentialNotation
    '
    '        End If
    '
    '    End If
    '
    '    GetFormattedValue = result

End Function

'Should be static method really
Public Function ParseExponent(ByVal formattedValue As String) As Long
    ParseExponent = Mid(formattedValue, InStr(1, UCase$(formattedValue), "E-") + 1, Len(formattedValue) - InStr(1, UCase$(formattedValue), "E-"))
End Function

