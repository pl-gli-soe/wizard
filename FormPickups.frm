VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPickups 
   Caption         =   "Add orders"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   OleObjectBlob   =   "FormPickups.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPickups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private pickups_handler As PickupsHandler

Public am_i_visible As Boolean


Public Function get_pickups_handler() As PickupsHandler
    Set get_pickups_handler = pickups_handler
End Function

Private Sub BtnDodaj_Click()
    
    
    ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Activate
    
    Set pickups_handler = New PickupsHandler
    
    
    With pickups_handler
        .connect_with_form_pickups E_ADD, Me
        '.quick_layout_config
        '.adjust_content_if_selection_changed
        
        .dodaj
    End With

End Sub

Private Sub CheckBoxOnlyFMAResp_Click()

    If ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1 Then
        pickups_handler.adjust_content_if_selection_changed
    End If
    
End Sub

Private Sub CheckBoxWorkOnlyOnVisibleRows_Click()
    
    If ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1 Then
        pickups_handler.adjust_content_if_selection_changed
    End If
End Sub

Private Sub ComboBoxPN_Change()
    If ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1 Then
        pickups_handler.adjust_content_if_selection_changed
    End If
End Sub

Private Sub ComboBoxSourceDUNS_Change()
    If ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1 Then
        pickups_handler.adjust_content_if_selection_changed
    End If
End Sub



Private Sub ListBoxCurrPusQty_Click()

    Me.TextBoxChangePUSQty.Value = Me.ListBoxCurrPusQty.Value
    
    Set pickups_handler = New PickupsHandler
    
    With pickups_handler
        .connect_with_form_pickups E_ADD, Me
        '.quick_layout_config
        '.adjust_content_if_selection_changed
        
        .add_form_listbox_qty_sth_selected
    End With

End Sub

Private Sub MultiPage_Change()

    pickups_handler.quick_layout_config
End Sub


Private Sub TextBoxChangePUSQty_Change()
    
    'Me.ListBoxCurrPusQty.value = Me.TextBoxChangePUSQty.Value
    For i = 0 To Me.ListBoxCurrPusQty.ListCount - 1
        If Me.ListBoxCurrPusQty.Selected(i) Then
            Me.ListBoxCurrPusQty.List(i) = Me.TextBoxChangePUSQty.Value
        End If
    Next i
End Sub

Private Sub TextBoxMaskForPusNumber_Change()
    
    Me.TextBoxPusName1.Value = Me.TextBoxMaskForPusNumber.Value
End Sub

Private Sub TextBoxPusName1_Change()
    Me.TextBoxMaskForPusNumber.Value = Me.TextBoxPusName1.Value
End Sub

Private Sub UserForm_Initialize()
    'Dim d As Date
    'd = Date
    'Me.MultiPage.Pages.sele
    'Me.MultiPage.Pages.Item("PageDUNS").DTPickerDeliveryDate.Value = CStr(d)
    'Me.MultiPage.Pages.Item("PageDUNS").DTPickerPickUpDate.Value = CStr(d)
    
    inner_init
    
    
End Sub

Public Sub inner_init()
    
    If Me.Visible Then
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
     Else
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
    End If
    
    
    Set pickups_handler = New PickupsHandler
    
    With pickups_handler
        .connect_with_form_pickups E_ADD, Me
        .fill_source_checkbox
        .quick_layout_config
        .adjust_content_if_selection_changed
    
    
    End With
    
    
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.am_i_visible = False
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
End Sub

Private Sub UserForm_Terminate()
    Me.am_i_visible = False
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
End Sub
