Attribute VB_Name = "Tests"
Option Explicit

Sub Test_All()
    Test_OneLineKeys
    test_ParsingParagraph
End Sub

Sub Test_OneLineKeys()
    Dim myLine As New cLine
    Dim mySecondLine As New cLine
    myLine.ParseText "Dim myLine As Integer"
    Debug.Assert myLine.Keywords.Count = 3
    Debug.Assert myLine.Keywords.item(2).Tag = "As"
    Debug.Assert myLine.Keywords.item(3).KeywordType = "DataType"
    mySecondLine.ParseText "For i = LBound(myArr) To UBound(myArr)"
    Debug.Assert mySecondLine.Keywords.Count = 4
    Debug.Assert mySecondLine.Keywords.item(2).Tag = "LBound"
    Debug.Assert mySecondLine.Keywords.item(1).Start = 1
    Debug.Assert mySecondLine.Keywords.item(4).Start = 26
End Sub

Sub Test_Formating()
    Dim myFirstLine As New cLine
    Dim mySecondLine As New cLine
    Dim i As Long
    myFirstLine.ParseText ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1).Range.Text
    Debug.Assert myFirstLine.Keywords.Count = 3
    Debug.Assert myFirstLine.Keywords.item(2).Tag = "As"
    Debug.Assert myFirstLine.Keywords.item(3).KeywordType = "DataType"
    For i = 1 To myFirstLine.Keywords.Count
        With ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1)
            ThisDocument.Range(.Range.Start + myFirstLine.Keywords.item(i).Start - 1, .Range.Start + myFirstLine.Keywords.item(i).Ende - 1).Style = myFirstLine.Keywords.item(i).KeywordType
        End With
    Next i
    
    mySecondLine.ParseText ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2).Range.Text
    Debug.Assert mySecondLine.Keywords.Count = 4
    
    For i = 1 To mySecondLine.Keywords.Count
        With ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2)
            ThisDocument.Range(.Range.Start + mySecondLine.Keywords.item(i).Start - 1, .Range.Start + mySecondLine.Keywords.item(i).Ende - 1).Style = mySecondLine.Keywords.item(i).KeywordType
        End With
    Next i
End Sub

Sub test_ParsingParagraph()
    Dim myLine As New cLine
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1)
    Debug.Assert myLine.Keywords.Count = 3
    Debug.Assert myLine.Keywords.item(2).Tag = "As"
    Debug.Assert myLine.Keywords.item(3).KeywordType = "DataType"
    Set myLine = New cLine
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2)
    Debug.Assert myLine.Keywords.Count = 4
    Debug.Assert myLine.Keywords.item(2).Tag = "LBound"
    Debug.Assert myLine.Keywords.item(1).Start = ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2).Range.Start
    Debug.Assert myLine.Keywords.item(4).Start = 47
End Sub

Sub test_ParsingParagraphAndFormatting()
    Dim myLine As New cLine
    Dim i As Long
    Dim er As ErrObject
    'First: Clear Formatting
    ThisDocument.StoryRanges(wdMainTextStory).Select
    On Error Resume Next
        Selection.ClearCharacterStyle 'Creates an Error if no characterstyle is set
        Set er = Err
        If er.Number = 4605 Then
            er.Clear
        Else
            Stop
        End If
    On Error GoTo 0
    Selection.EndOf
    
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(1)
    For i = 1 To myLine.Keywords.Count
        With myLine.Keywords.item(i)
            .Range.Style = .KeywordType
        End With
    Next i
    Set myLine = New cLine
    myLine.ParseParagraph ThisDocument.StoryRanges(wdMainTextStory).Paragraphs(2)
    For i = 1 To myLine.Keywords.Count
        With myLine.Keywords.item(i)
            .Range.Style = .KeywordType
        End With
    Next i
End Sub

Sub TestCreateCharStyles()
'Tested and working
    Dim highli As New Highlighter
    Set highli.Document = ThisDocument
    highli.CreateCharacterStyles
    highli.CreateCodeParagraphStyle
End Sub
