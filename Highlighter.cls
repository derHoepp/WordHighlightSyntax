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
Private mName As String
Private mCodeSections As CodeSections
Private mKeyWordDictionary As KeyWordDictionary
Private mOptions As HighlighterOptions
Private mAllKeyWords As Keywords

Public Property Get CodeSections() As CodeSections
    Set CodeSections = mCodeSections
End Property

Public Property Get Name() As String
    Name = mName & " in " & Me.Document.Name
End Property

Public Property Get Options() As HighlighterOptions
    Set Options = mOptions
End Property

Public Property Get KeyWordDictionary() As KeyWordDictionary
    Set KeyWordDictionary = mKeyWordDictionary
End Property

Public Sub AddKeyword(kw As Keyword)
    mAllKeyWords.Add kw
End Sub

Public Sub RecreateBuiltInStyles()
    mCreateCharacterStyles
    mCreateCodeParagraphStyle
End Sub

Private Sub mCreateCharacterStyles()
    Dim styleNames As Variant
    Dim styleColors As Variant
    Dim i As Long
    styleNames = Array("Keyword", "DataType", "String", "Comment")
    styleColors = Array(wdColorDarkBlue, wdColorOrange, wdColorLightBlue, wdColorGreen)
    If mDocument Is Nothing Then
        Err.Raise 90815, "Highlighter.CreateCharacterStyles", "No Document Set"
    Else
        For i = LBound(styleNames) To UBound(styleNames)
            CreateCustomCharacterStyle styleNames(i), styleColors(i)
        Next i
    End If
End Sub

Private Sub mCreateCodeParagraphStyle()
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
            .Font.Size = "10"
            .NoProofing = True
            With stl.ParagraphFormat
                .LeftIndent = 10
                .RightIndent = 10
                .Hyphenation = False
                .SpaceBefore = 0
                .SpaceAfter = 0
                .Shading.BackgroundPatternColor = wdColorGray10
            End With
        End With
    End If
End Sub

Public Sub CreateCustomCharacterStyle(ByVal StyleName As String, ByVal TextColor As WdColor)
    Dim stl As Style
    Dim i As Long
    If mDocument Is Nothing Then
        Err.Raise 90817, "Highlighter.CreateCustomCharacterStyle", "No Document Set"
    Else
        Set stl = Nothing
        On Error Resume Next
            Set stl = mDocument.Styles(StyleName)
        On Error GoTo 0
        If stl Is Nothing Then
            Set stl = mDocument.Styles.Add(StyleName, wdStyleTypeCharacter)
        End If
        stl.Font.TextColor = TextColor
    End If
End Sub

Public Property Set Document(doc As Document)
    Set mDocument = doc
    RecreateBuiltInStyles
End Property

Public Property Get Document() As Document
    Set Document = mDocument
End Property

Public Sub ParseDocument()

'READINGMARK: Working here

    Dim tmpCodeSection As CodeSection
    Dim para As Paragraph
    Set tmpCodeSection = New CodeSection
    tmpCodeSection.Init Me
    For Each para In mDocument.StoryRanges(wdMainTextStory).Paragraphs
        If para.Style = "Code" Then
            tmpCodeSection.AddParagraph para
        Else
            If tmpCodeSection.ParagraphCount > 0 Then
                mCodeSections.Add tmpCodeSection
                Set tmpCodeSection = New CodeSection
                tmpCodeSection.Init Me
            End If
        End If
    Next para
    If tmpCodeSection.ParagraphCount > 0 Then
        mCodeSections.Add tmpCodeSection
        Set tmpCodeSection = Nothing
    End If
    'Todo: Start Parsing All Sections
End Sub

Public Sub StyleAllKeywords()
'TODO: Only when Parsed
'TODO: Unstyle all
    Dim i As Long
    For i = 1 To mAllKeyWords.Count
        mAllKeyWords.Item(i).StyleKeyword
    Next i
End Sub

Private Sub Class_Initialize()
    Set mCodeSections = New CodeSections
    mCodeSections.Init Me
    If mDocument Is Nothing Then Set mDocument = ThisDocument
    mName = "SyntaxHighlighter"
    RecreateBuiltInStyles
    Set mOptions = New HighlighterOptions
    mOptions.Init Me
    Set mAllKeyWords = New Keywords
    mAllKeyWords.Init Me
    Set mKeyWordDictionary = New KeyWordDictionary
    mKeyWordDictionary.Init Me
End Sub
