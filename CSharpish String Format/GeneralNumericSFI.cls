VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneralNumericSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormatString(specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormatString = vbNullString
End Function

Private Function IStringFormatIdentifier_GetFormattedValue(value As Variant, specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormattedValue = GetFormattedValue(value, specifier)
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "G"
End Function

Public Function GetFormattedValue(value As Variant, specifier As StringFormatSpecifier) As String
    Dim result As String
    Dim exponentialNotation As String
    Dim power As Integer
    Dim exponentialFormat As New ExponentialSFI
    Dim fixedPointFormat As New FixedPointSFI
    Dim decimalFormat As New DecimalSFI

    Dim formatSpecifier As New StringFormatSpecifier
    formatSpecifier.Alignment = specifier.Alignment
    formatSpecifier.CustomSpecifier = specifier.CustomSpecifier

    If StringMatchesAny(TypeName(value), "Integer", "Long") Then

        formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "D")
        result = decimalFormat.GetFormattedValue(value, formatSpecifier)

    ElseIf TypeName(value) = "Double" Then

        formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "E")
        exponentialNotation = exponentialFormat.GetFormattedValue(value, formatSpecifier)
        power = exponentialFormat.GetPower(exponentialNotation)

        If power > -5 And Abs(power) < specifier.Precision Then

            formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "F")
            result = fixedPointFormat.GetFormattedValue(value, formatSpecifier)

        Else

            result = exponentialNotation

        End If

    End If

    GetFormattedValue = result

End Function

