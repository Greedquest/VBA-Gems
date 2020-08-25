VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter")
Option Explicit

Private Type tSpecifier
    Index As Integer
    identifier As String
    alignmentSpecifier As Integer
    precisionSpecifier As Integer
    CustomSpecifier As String
End Type

Private this As tSpecifier

Public Property Get Index() As Integer
    Index = this.Index
End Property

Public Property Let Index(value As Integer)
    this.Index = value
End Property

Public Property Get identifier() As String
    identifier = this.identifier
End Property

Public Property Let identifier(value As String)
    this.identifier = value
End Property

Public Property Get Alignment() As Integer
    Alignment = this.alignmentSpecifier
End Property

Public Property Let Alignment(value As Integer)
    this.alignmentSpecifier = value
End Property

Public Property Get Precision() As Integer
    Precision = this.precisionSpecifier
End Property

Public Property Get CustomSpecifier() As String
    CustomSpecifier = this.CustomSpecifier
End Property

Public Property Let CustomSpecifier(value As String)
    this.CustomSpecifier = value
    If IsNumeric(value) And Val(value) <> 0 Then this.precisionSpecifier = CInt(value)
End Property

Public Function ToString() As String
    ToString = "{" & this.Index & _
               IIf(this.alignmentSpecifier <> 0, _
                   "," & this.alignmentSpecifier, vbNullString) & _
                   IIf(this.identifier <> vbNullString, _
                       ":" & this.identifier, vbNullString) & _
                       IIf(this.CustomSpecifier <> vbNullString, _
                           this.CustomSpecifier, vbNullString) & "}"
End Function


