VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRawtext As String
Private mKeywords As cKeywords
Private mKeywordDict As Collection
Private mParagraph As Paragraph

Public Sub ParseParagraph(Para As Paragraph)
    Dim wd As Range
    Dim tmpKeyWord As cKeyword
    Dim actCategory As String
    
    initKeyWordDict
    For Each wd In Para.Range.Words
        'The Range of a word includes spacecharacters.
        On Error Resume Next
            actCategory = mKeywordDict.item(Replace(wd.Text, " ", ""))
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
            actWord = mKeywordDict.item(myArr(i))
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
    Dim KeyWordIni As String
    Dim KWArray As Variant
    Dim KeyWordGroup As String
    Dim i As Long
    KeyWordIni = CreateObject("Scripting.FilesystemObject").OpenTextfile("C:\Daten\GitRepos\WordHighlightSyntax\VBASyntax.ini").ReadAll
    KWArray = Split(KeyWordIni, vbCrLf)
    If Left(KeyWordIni, 1) = "[" Then
        For i = LBound(KWArray) To UBound(KWArray)
            If KWArray(i) Like "[[]*]" Then
                KeyWordGroup = Mid(KWArray(i), 2, Len(KWArray(i)) - 2)
            Else
                mKeywordDict.Add KeyWordGroup, KWArray(i)
            End If
        Next i
    Else
        Err.Raise 9002, "initKeyWordDict", "No Ini-File"
    End If
    
End Sub

Private Sub Class_Initialize()
    Set mKeywords = New cKeywords
    Set mKeywordDict = New Collection
End Sub
