VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Keywords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeywords As Collection
Private mHighlighter As Highlighter

Public Property Get myHighlighter() As Highlighter
    Set myHighlighter = mHighlighter
End Property

Public Sub Add(Item As Keyword)
    mKeywords.Add Item
End Sub

Public Property Get Count() As Long
    Count = mKeywords.Count
End Property

Public Property Get Item(index As Long) As Keyword
    If mKeywords.Count < 1 Then
        Set Item = Nothing
    Else
        Set Item = mKeywords(index)
    End If
End Property

Public Sub Init(Highli As Highlighter)

End Sub

Private Sub Class_Initialize()
    If mKeywords Is Nothing Then Set mKeywords = New Collection
End Sub
