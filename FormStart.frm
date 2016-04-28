VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormStart 
   Caption         =   "Start"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormStart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormStart"
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


Private Sub BtnAddData_Click()
    inner_dopisz
End Sub

Private Sub BtnPodmien_Click()
    hide
    add_data E_NADPISZ
End Sub

Private Sub BtnDetails_Click()
    
    hide
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
    
    If Me.Caption Like "*EDIT*" Then
        detailsfill False
    ElseIf Me.Caption Like "*NEW*" Then
        detailsfill True
    End If
End Sub


Private Sub detailsfill(new_details_keywords As Boolean)

    Dim dh As DetailsHandler
    Set dh = New DetailsHandler
    
    ' arg is for is it new definition of keywords in project details
    dh.init_wizard_for_details new_details_keywords
    
    Set dh = Nothing
End Sub

Private Sub BtnValid_Click()
    hide
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
    my_validation ThisWorkbook.Sheets(MASTER_SHEET_NAME)
End Sub

