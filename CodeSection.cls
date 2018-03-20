VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mParagraphs As Collection
Private mText As String
Private mHighlighter As Highlighter
Public Property Get myHighlighter() As Highlighter
    Set myHighlighter = mHighlighter
End Property

Public Sub AddParagraph(p As Paragraph)
    mParagraphs.Add p
    mText = mText & p.Range.Text
End Sub

Public Property Get ParagraphCount() As Long
    ParagraphCount = mParagraphs.Count
End Property

Public Property Get HasChanged() As Boolean
    Dim currentText As String
    Dim p As Paragraph
    For Each p In mParagraphs
        currentText = currentText & p.Range.Text
    Next p
    HasChanged = currentText <> mText
    'Todo: Implement Strategy to replace mText with CurrentText after dealing with changes
End Property

Public Sub Init(Highli As Highlighter)
    Set mHighlighter = Highli
End Sub

Private Sub Class_Initialize()
    Set mParagraphs = New Collection
End Sub
