Attribute VB_Name = "Tests"
Option Explicit

Sub TestAll()
    TestOneLineKeys
    'TestFormatingWithCertainDocumentStructure 'Currently wrong doc
    'TestParsingParagraph 'Currently wrong doc
    TestCreateCharStyles
    TestParsingParagraphAndFormatting
    
End Sub

Sub TestOneLineKeys()
    Dim myLine As New cLine
    Dim mySecondLine As New cLine
    myLine.ParseText "Dim myLine As Integer"
    Debug.Assert myLine.Keywords.Count = 3
    Debug.Assert myLine.Keywords.Item(2).Tag = "As"
    Debug.Assert myLine.Keywords.Item(3).KeywordType = "DataType"
    mySecondLine.ParseText "For i = LBound(myArr) To UBound(myArr)"
    Debug.Assert mySecondLine.Keywords.Count = 4
    Debug.Assert mySecondLine.Keywords.Item(2).Tag = "LBound"
    Debug.Assert mySecondLine.Keywords.Item(1).Start = 1
    Debug.Assert mySecondLine.Keywords.Item(4).Start = 26
End Sub

Sub TestFormatingWithCertainDocumentStructure()
    'Current Document does no longer support this Test, as Paragraphs have moved
    Dim myFirstLine As New cLine
    Dim mySecondLine As New cLine
    Dim i As Long
    myFirstLine.ParseText ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1).Range.Text
    Debug.Assert myFirstLine.Keywords.Count = 3
    Debug.Assert myFirstLine.Keywords.Item(2).Tag = "As"
    Debug.Assert myFirstLine.Keywords.Item(3).KeywordType = "DataType"
    For i = 1 To myFirstLine.Keywords.Count
        With ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1)
            ThisDocument.Range(.Range.Start + myFirstLine.Keywords.Item(i).Start - 1, .Range.Start + myFirstLine.Keywords.Item(i).Ende - 1).Style = myFirstLine.Keywords.Item(i).KeywordType
        End With
    Next i
    
    mySecondLine.ParseText ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2).Range.Text
    Debug.Assert mySecondLine.Keywords.Count = 4
    
    For i = 1 To mySecondLine.Keywords.Count
        With ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2)
            ThisDocument.Range(.Range.Start + mySecondLine.Keywords.Item(i).Start - 1, .Range.Start + mySecondLine.Keywords.Item(i).Ende - 1).Style = mySecondLine.Keywords.Item(i).KeywordType
        End With
    Next i
End Sub

Sub TestParsingParagraph()
    'Current Document does no longer support this Test, as Paragraphs have moved
    Dim myLine As New cLine
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1)
    Debug.Assert myLine.Keywords.Count = 3
    Debug.Assert myLine.Keywords.Item(2).Tag = "As"
    Debug.Assert myLine.Keywords.Item(3).KeywordType = "DataType"
    Set myLine = New cLine
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2)
    Debug.Assert myLine.Keywords.Count = 4
    Debug.Assert myLine.Keywords.Item(2).Tag = "LBound"
    Debug.Assert myLine.Keywords.Item(1).Start = ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2).Range.Start
    Debug.Assert myLine.Keywords.Item(4).Start = 47
End Sub

Sub TestParsingParagraphAndFormatting()
    Dim myLine As New cLine
    Dim pars As Variant
    Dim j As Long
    Dim i As Long
    Dim er As ErrObject
    pars = Array(2, 3, 6, 7, 8, 9, 10, 11, 12, 13)
    'First: Clear Formatting
    ThisDocument.StoryRanges(wdMainTextStory).Select
    On Error Resume Next
        Selection.ClearCharacterStyle 'Creates an Error if no characterstyle is set
        Set er = Err
        If er.Number = 4605 Or er.Number = 0 Then
            er.Clear
        Else
            Stop
        End If
    On Error GoTo 0
    Selection.EndOf
    
    For j = LBound(pars) To UBound(pars)
        Set myLine = New cLine
        myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(pars(j))
        For i = 1 To myLine.Keywords.Count
            With myLine.Keywords.Item(i)
                .Range.Style = .KeywordType
            End With
        Next i
    Next j
End Sub

Sub TestCreateCharStyles()
'Tested and working
    Dim Highli As New Highlighter
    Set Highli.Document = ThisDocument
    Highli.CreateCharacterStyles
    Highli.CreateCodeParagraphStyle
End Sub
