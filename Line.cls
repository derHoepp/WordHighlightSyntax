VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRawtext As String
Private mKeywords As cKeywords
'Private mKeyWordDict As Collection 'Obsolete -> Centralized in Highlighter
Private mParagraph As Paragraph
Private mHighlighter As Highlighter

Public Sub ParseParagraph(para As Paragraph)
    Dim wd As Range
    Dim tmpKeyWord As cKeyword
    Dim actCategory As String
    
    For Each wd In para.Range.Words
        'The Range of a word includes spacecharacters.
        On Error Resume Next
            actCategory = mHighlighter.KeyWordDictionary.Item(Replace(wd.Text, " ", ""))
        On Error GoTo 0
        If Not actCategory = vbNullString Then
            Set tmpKeyWord = New cKeyword
            Set tmpKeyWord.Range = wd
            tmpKeyWord.KeywordType = actCategory
            mKeywords.Add tmpKeyWord
        End If
        actCategory = vbNullString
    Next wd
End Sub

Public Sub ParseText(RawLine As String)
    Dim myArr As Variant
    Dim i As Long
    Dim tmpLen As Long
    Dim actWord As String
    Dim tmpKeyWord As cKeyword
    
    tmpLen = 1
    mRawtext = vbNullString
    mRawtext = RawLine
    initKeyWordDict
    
    myArr = Split(Replace(Replace(Replace(Replace(RawLine, "(", " "), ")", " "), vbCr, " "), vbLf, " "), Space(1))
    For i = LBound(myArr) To UBound(myArr)
        On Error Resume Next
            actWord = mKeyWordDict.Item(myArr(i))
        On Error GoTo 0
        If Not actWord = vbNullString Then
            Set tmpKeyWord = New cKeyword
            tmpKeyWord.Start = tmpLen
            tmpKeyWord.Tag = myArr(i)
            tmpKeyWord.KeywordType = actWord
            mKeywords.Add tmpKeyWord
        End If
        tmpLen = tmpLen + Len(myArr(i)) + 1 'To determine the start
        actWord = vbNullString
    Next i
End Sub

Public Property Get Keywords() As cKeywords
    Set Keywords = mKeywords
End Property

Private Sub initKeyWordDict()
   'Obsolete -> Centralized in Highlighter
End Sub

Private Sub Class_Initialize()
    Set mKeywords = New cKeywords
    Set mKeyWordDict = New Collection
End Sub
