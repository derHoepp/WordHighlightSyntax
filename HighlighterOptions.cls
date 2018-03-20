VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighlighterOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIniFileName As String
Private mHighliter As Highlighter

Public Property Get myHighlighter() As Highlighter
    Set myHighlighter = mHighliter
End Property

Public Property Get KeywordIniFileName() As String
    KeywordIniFileName = mIniFileName
End Property

Public Sub Init(Highli As Highlighter)
    If Dir(Highli.Document.Path & "\VBASyntax.ini") = vbNullString Then
        Err.Raise "90819", "HighliterOptions.Init", "Inifile not found in DocumentDirectory"
    Else
        mIniFileName = Highli.Document.Path & "\VBASyntax.ini"
    End If
    Set mHighliter = Highli
End Sub
