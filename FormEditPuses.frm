VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEditPuses 
   Caption         =   "Edytuj PUSy"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   OleObjectBlob   =   "FormEditPuses.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormEditPuses"
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


Private pickups_handler As PickupsHandler

Public am_i_visible As Boolean


Public Function get_pickups_handler() As PickupsHandler
    Set get_pickups_handler = pickups_handler
End Function

Private Sub BtnCancel_Click()
    hide
End Sub

Private Sub BtnDelete_Click()
    answer = MsgBox("Czy jestes pewien?", vbYesNo)
    
    
    'CzekajForm.show
    If answer = vbYes Then
        
        
        
        pus_name = zwroc_wyselekcjonowany_pus_name()
        If pus_name <> "" Then
            usun_wybrany_pickup_po_potwierdzeniu CStr(pus_name)
            inner_init
        Else
            MsgBox "nie wybrano!"
        End If
        
    Else
        ' nop
    End If
    
    'CzekajForm.hide
End Sub

Private Function zwroc_wyselekcjonowany_pus_name()

    zwroc_wyselekcjonowany_pus_name = ""
    
    For x = 0 To Me.ListBoxPUSes.ListCount - 1
        
        If Me.ListBoxPUSes.Selected(x) Then
            zwroc_wyselekcjonowany_pus_name = Me.ListBoxPUSes.List(x)
            Exit For
        End If
        
    Next x
End Function

Private Sub BtnDodajPN_Click()


    Dim msh As Worksheet, r As Range

    With DodajPNForm
    
        If Me.ListBoxPUSes.ListCount > 0 And Me.ListBoxINDX.ListCount > 0 Then
            .TextBoxPUSName = Me.ListBoxPUSes.Value
            .DTPickerPUSDate.Value = Me.ListBoxPickupDate.List(0)
            .DTPickerDelDate.Value = Me.ListBoxDelDate.List(0)
            .TextBoxPtrn.Value = ""
            .TextBoxBufferForIndx0.Value = Me.ListBoxINDX.List(0)
            .ListBoxINDX.Clear
    
            Set msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
            Set r = msh.Cells(2, WizardMain.pn)
            Do
                If Me.ListBoxINDX.List(0) Like _
                    "*" & CStr(msh.Cells(r.Row, WizardMain.duns)) & "," & CStr(msh.Cells(r.Row, WizardMain.fup_code)) Then
                    
                        .ListBoxINDX.AddItem CStr(msh.Cells(r.Row, WizardMain.pn)) & _
                            "," & CStr(msh.Cells(r.Row, WizardMain.duns)) & _
                            "," & CStr(msh.Cells(r.Row, WizardMain.fup_code))
                End If
                
                WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
                
            Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    
            .show
        Else
            MsgBox "lista PUSow oraz lista PNow jest pusta nie ma co zrobic"
        End If
    End With
End Sub

Private Sub inner_edit()
    Set pickups_handler = New PickupsHandler
    
    pus_name = ""
    With pickups_handler
        .connect_with_form_pickups E_EDIT, Me
        '.quick_layout_config
        '.adjust_content_if_selection_changed
        x = .edit_puses
    End With
    
    inner_init
    
    If x >= 0 Then
        Me.ListBoxPUSes.Selected(x) = True
    End If
End Sub



Public Sub inner_init()
    
    If Me.Visible Then
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        Me.am_i_visible = True
     Else
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
        Me.am_i_visible = False
    End If
    
    
    Set pickups_handler = New PickupsHandler
    
    With pickups_handler
        .connect_with_form_pickups E_EDIT, Me
        .fill_edit_listboxes
        .adjust_content_if_selection_changed
    
    
    End With
    
    
End Sub

Private Sub BtnUsunPN_Click()
        
    'CzekajForm.show
    With pickups_handler
        .connect_with_form_pickups E_ADD, Me
        '.quick_layout_config
        '.adjust_content_if_selection_changed
        
        .remove_this_pn
    End With
    'CzekajForm.hide
    
End Sub


Private Sub ListBoxPUSes_Change()
    ' zmiana na listbox
    
    Set pickups_handler = New PickupsHandler
    
    
    
    With pickups_handler
        .connect_with_form_pickups E_ADD, Me
        '.quick_layout_config
        '.adjust_content_if_selection_changed
        
        
        .pus_listbox_change
        .fill_labels
    End With
End Sub

Private Sub ListBoxPUSes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    ' bedzie to jedyna procedura, ktora nie jest sterowana z poziomu
    ' niezaleznego obiektu
    Me.hide
    With ZmienPUSaForm
        ' date picker
        For x = 0 To Me.ListBoxDelDate.ListCount - 1
        
            If Me.ListBoxPUSes.Selected(x) = True Then
            
            
                ' podwojnie by zachowac pierwotne wartosci
                .TextBoxPUSName.Value = Me.ListBoxPUSes.List(x)
                .TextBoxPUSName2.Value = .TextBoxPUSName
                
                .DTPickerPUSDate.Value = Me.ListBoxPickupDate.List(x)
                .DTPickerPUSDate2.Value = .DTPickerPUSDate
                
                .DTPickerDelDate.Value = Me.ListBoxDelDate.List(x)
                .DTPickerDelDate2.Value = .DTPickerDelDate
                Exit For
            End If
        Next x
        .show
    End With
    
    
End Sub

Private Sub ListBoxQty_Click()


    Set pickups_handler = New PickupsHandler
    With pickups_handler
        .connect_with_form_pickups E_EDIT, Me
        .listbox_clicked
    End With

End Sub

Private Sub ListBoxQty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


    ' MsgBox Int(KeyAscii)
    klick_enter_now_edit = False

    Set pickups_handler = New PickupsHandler
    With pickups_handler
        .connect_with_form_pickups E_EDIT, Me
        klick_enter_now_edit = .edit_qty_key_pressed(KeyAscii)
    End With
    
    
    If CBool(klick_enter_now_edit) Then
        inner_edit
        ' ThisWorkbook.Save
    End If
End Sub

Private Sub TextBox1_Change()

    ' juz mi sie nazwy nie chcialo zmieniac
    ' to jest textbox do patternu zeby ograniczyc mozliwosci
    ' listy i zeby sie miescilaw w widoku usera w ogole :)

    Set pickups_handler = New PickupsHandler
    With pickups_handler
        .connect_with_form_pickups E_EDIT, Me
        .zmniejsz_liste_indx_poprzez_ptrn
    End With

End Sub

Private Sub UserForm_Initialize()
    inner_init
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.am_i_visible = False
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
End Sub

Private Sub UserForm_Terminate()
    Me.am_i_visible = False
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
End Sub
