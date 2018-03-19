VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKeywords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeywords As Collection

Public Sub Add(item As cKeyword)
    mKeywords.Add item
End Sub

Public Property Get Count() As Long
    Count = mKeywords.Count
End Property

Public Property Get item(Index As Long) As cKeyword
    If mKeywords.Count < 1 Then
        Set item = Nothing
    Else
        Set item = mKeywords(Index)
    End If
End Property

Private Sub Class_Initialize()
    If mKeywords Is Nothing Then Set mKeywords = New Collection
End Sub
