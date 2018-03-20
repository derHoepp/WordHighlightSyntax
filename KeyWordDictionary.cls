VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyWordDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mHighlighter As Highlighter
Dim mKeyWordDict As Collection

Public Sub Init(Highli As Highlighter)
    Dim KWArray As Variant
    Dim KeyWordIni As String
    Dim KeyWordGroup As String
    Dim i As Long
    Set mHighlighter = Highli
    KeyWordIni = CreateObject("Scripting.FilesystemObject").OpenTextfile(Highli.Options.KeywordIniFileName).ReadAll
    KWArray = Split(KeyWordIni, vbCrLf)
    If Left(KeyWordIni, 1) = "[" Then
        For i = LBound(KWArray) To UBound(KWArray)
            If KWArray(i) Like "[[]*]" Then
                KeyWordGroup = Mid(KWArray(i), 2, Len(KWArray(i)) - 2)
            Else
                mKeyWordDict.Add KeyWordGroup, KWArray(i)
            End If
        Next i
    Else
        Err.Raise 9002, "KeyWordDictionary.Init", "No Ini-File"
    End If
End Sub

Public Property Get Item(key As String) As String
    Item = vbNullString
    On Error Resume Next
        Item = mKeyWordDict.Item(key)
    On Error GoTo 0
End Property

Private Sub Class_Initialize()
    Set mKeyWordDict = New Collection
End Sub
