VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

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
