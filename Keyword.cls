VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Keyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText As String
Private mStart As Long
Private mRange As Word.Range
Private mCategory As String

Public Property Get Text() As String
    Tag = mText
End Property

Public Property Let Category(val As String)
    mCategory = val
End Property
Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Get Range() As Range
    Set Range = mRange
End Property

Public Property Set Range(rng As Range)
    Set mRange = rng
    mText = Replace(mRange.Text, " ", "") 'The Range contains spaces, but shouldn't
End Property

Public Sub StyleKeyword()
    mRange.Style = mCategory
End Sub
