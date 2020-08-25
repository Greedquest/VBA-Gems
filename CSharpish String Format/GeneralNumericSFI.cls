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

Private Function IStringFormatIdentifier_GetFormatString(ByVal specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormatString = vbNullString
End Function

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormattedValue = GetFormattedValue(value, specifier)
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "G"
End Function

Public Function GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    Dim result As String
    Dim exponentialNotation As String
    
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
        
        'Check the power returned by this call, since General
        'only uses exponential notation in certain circumstances
        Dim exponent As Long
        exponent = exponentialFormat.ParseExponent(exponentialNotation)
        If exponent > -5 And Abs(exponent) < specifier.Precision Then
            formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "F")
            result = fixedPointFormat.GetFormattedValue(value, formatSpecifier)

        Else
            result = exponentialNotation

        End If

    End If
    
    GetFormattedValue = result
    
End Function

