VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText As String
Private mStart As Long
Private mRange As Word.Range
Private mKeywordType As String

Public Property Let Tag(val As String)
    mText = val
End Property

Public Property Get Tag() As String
    Tag = mText
End Property

Public Property Let KeywordType(val As String)
    mKeywordType = val
End Property
Public Property Get KeywordType() As String
    KeywordType = mKeywordType
End Property

Public Property Let Start(val As Long)
    mStart = val
End Property

Public Property Get Start() As Long
    Start = mStart
End Property

Public Property Get Ende() As Long
    Ende = Start + Len(mText)
End Property

Public Property Get Range() As Range
    Set Range = mRange

End Property

Public Property Set Range(rng As Range)
    Set mRange = rng
    mText = Replace(mRange.Text, " ", "") 'The Range contains spaces, but shouldn't
    mStart = mRange.Start
    
End Property
