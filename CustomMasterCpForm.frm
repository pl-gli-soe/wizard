VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomMasterCpForm 
   Caption         =   "Take data from raw T/D"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5820
   OleObjectBlob   =   "CustomMasterCpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomMasterCpForm"
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

Public podmien_handler As PodmienHandler

Private Sub BtnClear_Click()
    'quick_labels_assign
    'respawn_listbox
    
    
    For x = 0 To Me.ListBoxRawData.ListCount - 1
        Me.ListBoxRawData.Selected(x) = False
    Next x
End Sub

Private Sub BtnEnd_Click()
    hide
    With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
        .Range("d2") = 0
        .Range("d3") = ""
    End With
End Sub

Private Sub BtnNxt_Click()
    
    
    ' jednoczesnie zapisuje dane
    work_around_custom_copy_worksheet
    
    
    ' na koniec powieksz
    Dim tr As Range
    Set tr = ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("a1").End(xlDown)
    If ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2") = tr Then
        
        ' do nothing - you are at the end of list so now you making some custom stuff
    Else
        nizej
    End If
    
    
    If ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2") < 2 Then
        Me.BtnPrev.Enabled = False
    Else
        Me.BtnPrev.Enabled = True
    End If
    
    quick_labels_assign
    
    respawn_listbox
    
End Sub

Private Sub nizej()


    With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
    
        .Range("d2") = .Range("d2") + 1
        
        If .Cells(.Range("d2"), 7) <> "x" Then
            nizej
        End If
    End With
End Sub


Private Sub wyzej()


    With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
        .Range("d2") = .Range("d2") - 1
        
        If .Cells(.Range("d2"), 7) <> "x" Then
            wyzej
        End If
    End With
End Sub

Private Sub BtnPrev_Click()

    ' na koniec powieksz
    wyzej
        
    If ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2") < 2 Then
        Me.BtnPrev.Enabled = False
    Else
        Me.BtnPrev.Enabled = True
    End If
    
    quick_labels_assign
    
    respawn_listbox

End Sub

Private Sub quick_labels_assign()

    With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
    
        CustomMasterCpForm.Label1.Caption = _
            .Cells(.Range("d2"), 2)
        CustomMasterCpForm.Label2.Caption = WizardMain.NIC_NIE_WYBRANO_TXT
        
        
        If CustomMasterCpForm.Label1.Caption = CStr(.Range("B28")) Then
            
            CustomMasterCpForm.ListBoxRawData.MultiSelect = fmMultiSelectMulti
        Else
            CustomMasterCpForm.ListBoxRawData.MultiSelect = fmMultiSelectSingle
        End If
    
    End With
End Sub


Private Sub respawn_listbox()
    
    
    Dim ssh As Worksheet
    Set ssh = podmien_handler.source_workbook.Sheets(WizardMain.RAW_TXT)
    
    Me.ListBoxRawData.Clear
    Dim rr As Range
    Set rr = ssh.Range("A1")
    i = 1
    Do
        
        Me.ListBoxRawData.AddItem rr
    
        Set rr = rr.Offset(0, 1)
        
    Loop Until Trim(CStr(rr)) = ""
    
End Sub


Private Sub BtnUstaw_Click()
    work_around_custom_copy_worksheet
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    hide
    ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2") = 0
End Sub
