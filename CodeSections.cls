VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mCollection As Collection
Dim mHighlighter As Highlighter

Public Property Get myHighlighter() As Highlighter
    Set myHighlighter = mHighlighter
End Property

Public Sub Add(Item As CodeSection)
    mCollection.Add Item
End Sub
Public Property Get Count() As Long
    Count = mCollection.Count
End Property

Public Property Get Item(index As Long) As CodeSection
    Set Item = mCollection(index)
End Property

Public Sub Init(Highli As Highlighter)
    Set mHighlighter = Highli
End Sub


Private Sub Class_Initialize()
    Set mCollection = New Collection
End Sub
