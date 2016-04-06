VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rng As Range
Private cmnt As Comment
Private txt As String
Private type_of_comment As E_TYPE_OF_COMMENT

Public Function getTxt() As String
    getTxt = txt
End Function


Public Function typeOfComment() As E_TYPE_OF_COMMENT
    typeOfComment = type_of_comment
End Function


Public Sub init(ByRef r As Range)
    Set rng = r
    Set cmnt = rng.AddComment()
End Sub

Public Sub initWithTxt(ByRef r As Range, txt As String)
    Set rng = r
    rng.AddComment txt
    Set cmnt = rng.Comment
End Sub

Public Sub overwriteTxt(txt As String)
    cmnt.Text txt, , True
End Sub

Public Sub appendTxt(txt As String)
    cmnt.Text txt, Len(cmnt.Text), False
End Sub

Public Sub adjustSizeOfThisCmnt()
    cmnt.Shape.TextFrame.AutoSize = True
End Sub
