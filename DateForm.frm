VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   Caption         =   "Date"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1905
   OleObjectBlob   =   "DateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Sub BtnSubmit_Click()
    hide
    Selection.Value = Me.DTPicker1.Value
End Sub

Private Sub DTPicker1_AfterUpdate()

    Selection.Value = Me.DTPicker1.Value
End Sub


Private Sub DTPicker1_DblClick()
    hide
    Selection.Value = Me.DTPicker1.Value
End Sub
