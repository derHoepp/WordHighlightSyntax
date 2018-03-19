VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Highlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mDocument As Document

Public Sub CreateCharacterStyles()
    Dim stl As Style
    Dim styleNames As Variant
    Dim styleColors As Variant
    Dim i As Long
    styleNames = Array("Keyword", "DataType", "String", "Comment")
    styleColors = Array(wdColorDarkBlue, wdColorOrange, wdColorLightBlue, wdColorGreen)
    If mDocument Is Nothing Then
        Err.Raise 90815, "Highlighter.CreateCharacterStyles", "No Document Set"
    Else
        For i = LBound(styleNames) To UBound(styleNames)
            Set stl = Nothing
            On Error Resume Next
                Set stl = mDocument.Styles(styleNames(i))
            On Error GoTo 0
            If stl Is Nothing Then
                Set stl = mDocument.Styles.Add(styleNames(i), wdStyleTypeCharacter)
            End If
            stl.Font.TextColor = styleColors(i)
        Next i
    End If
End Sub

Public Sub CreateCodeParagraphStyle()
    Dim stl As Style
    If mDocument Is Nothing Then
        Err.Raise 90815, "Highlighter.CreateCodeParagraphStyle", "No Document Set"
    Else
        Set stl = Nothing
        On Error Resume Next
            Set stl = mDocument.Styles("Code")
        On Error GoTo 0
        If stl Is Nothing Then
            Set stl = mDocument.Styles.Add("Code", wdStyleTypeParagraph)
        End If
        With stl
            .BaseStyle = mDocument.Styles(wdStyleNormal)
            .Font.Name = "Courier New"
            .Font.Size = "11"
            .NoProofing = True
            With stl.ParagraphFormat
                .Hyphenation = False
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        End With
    End If
End Sub

Public Property Set Document(doc As Document)
    Set mDocument = doc
End Property

Public Property Get Document() As Document
    Set Document = mDocument
End Property
