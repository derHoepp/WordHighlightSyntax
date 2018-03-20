Attribute VB_Name = "AllNewTests"
Option Explicit

Sub AllTests()
    TestHighlighterStyleCreation
    TestHighlighterCodeSectionDetection
End Sub

Sub TestHighlighterStyleCreation()
    Dim Highli As New Highlighter
    Set Highli.Document = ThisDocument
    
    Debug.Assert Not ThisDocument.Styles("Code") Is Nothing
    Debug.Assert ThisDocument.Styles("Code").Type = wdStyleTypeParagraph
    Debug.Assert Not ThisDocument.Styles("Comment") Is Nothing
    Debug.Assert ThisDocument.Styles("Comment").Type = wdStyleTypeCharacter
    Debug.Assert ThisDocument.Styles("DataType").Font.TextColor = wdColorOrange
End Sub

Sub TestHighlighterCodeSectionDetection()
    Dim Highli As New Highlighter
    Set Highli.Document = ThisDocument
    Highli.ParseDocument
    Debug.Assert Highli.CodeSections.Count = 2
    Debug.Assert Highli.CodeSections.Item(2).ParagraphCount = 8
    Debug.Assert Highli.CodeSections.Item(1).HasChanged = False
    ThisDocument.Paragraphs(2).Range.Text = Replace(ThisDocument.Paragraphs(2).Range.Text, "Integer", "Long")
    Debug.Assert Highli.CodeSections.Item(1).HasChanged = True
    Debug.Assert Highli.CodeSections.Item(2).HasChanged = False
End Sub
