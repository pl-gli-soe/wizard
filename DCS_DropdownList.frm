VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DCS_DropdownList 
   Caption         =   "DCS"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   OleObjectBlob   =   "DCS_DropdownList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DCS_DropdownList"
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


Public dcs_handler As DCSHandler


Private Sub BtnSubmit_Click()


    With Me.ListBox2
        .Selected(0) = True
    End With

    Me.hide
    Set Me.dcs_handler.r = ActiveCell
    
    If Me.ListBox2.List(Me.ListBox2.ListCount - 1) = ThisWorkbook.Sheets(DCS_SHEET_NAME).Range("A2").Value Then
        Me.dcs_handler.r.Value = ""
    Else
        Me.dcs_handler.r.Value = Me.ListBox2.List(Me.ListBox2.ListCount - 1)
    End If
End Sub

Private Sub ListBox1_Click()
    ' Hide
    Me.dcs_handler.work_around_listbox1_and_listbox2
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.hide
    Set Me.dcs_handler.r = ActiveCell
    
    If Me.ListBox2.Value = ThisWorkbook.Sheets(DCS_SHEET_NAME).Range("A2").Value Then
        Me.dcs_handler.r.Value = ""
    Else
        Me.dcs_handler.r.Value = Me.ListBox2.Value
    End If
End Sub

Private Sub TextBox1_Change()
    Me.dcs_handler.work_around_listbox1_and_listbox2
End Sub


Private Sub UserForm_Initialize()
    Set dcs_handler = New DCSHandler
    Set Me.dcs_handler.r = Selection
    Me.TextBox1.Value = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set dcs_handler = Nothing
    Me.hide
End Sub
