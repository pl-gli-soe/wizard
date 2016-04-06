VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailsFormsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private current_form As Variant
Private rng_pointer As Range
Private det As Worksheet


' to jest zmienna publiczna i nalezy
' ja uruchomic zaraz po konstruktorze
Public new_definition As Boolean



' slot pod tymczasowy strinh ktory wspolpracuje z zmienna d ponizej
Private tmp As String
' data z ktorej korzystam przez cale zycie obiektu
Private d As Date

Public Function get_current_form() As Variant
    Set get_current_form = current_form
End Function



Private Sub Class_Initialize()

    
    Set current_form = Nothing
    Set det = ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
    Set rng_pointer = det.Range("C1")
    'rng_pointer.Parent.Unprotect 123
    
    clear_all_pointers
End Sub

Private Sub Class_Terminate()
    Set current_form = Nothing
    'rng_pointer.Parent.Protect 123
End Sub


Public Sub init(ByRef e As E_NEW_PROJECT_ITEM, ByRef dh As DetailsHandler)


    clear_all_pointers

    If e = PLT Then
        
        logic_for_init_in_txtbox_form e, dh, "__", PIERWSZY
        
    ElseIf e = PROJECT Then
        
        logic_for_init_in_txtbox_form e, dh, "", SRODEK
        
    ElseIf e = BIW_GA Then
    
        logic_for_biw_ga_directly e, dh
    ElseIf e = MY Then
    
        logic_for_init_in_txtbox_form e, dh, "00.00", SRODEK
        
    ElseIf e = PHAZE Then
        
        logic_for_init_in_txtbox_form e, dh, "", SRODEK
        
        
    ElseIf e = BOM Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_CW, e, dh, SRODEK
    
    ElseIf e = PICKUP_DATE Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_DATE, e, dh, SRODEK
        
    ElseIf e = PPAP_GATE Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_DATE, e, dh, SRODEK
        
    ElseIf e = mrd Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_CW, e, dh, SRODEK
        
    ElseIf e = BUILD_START Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_CW, e, dh, SRODEK
        
    ElseIf e = BUILD_END Then
    

        logic_for_init_in_dtpicker_form WizardMain.E_DC_CW, e, dh, SRODEK
    
    
    ElseIf e = KOORDYNATOR Then
        logic_for_init_in_txtbox_form e, dh, "", SRODEK
        
        
    ElseIf e = E_ACTIVE Then
    
        logic_for_active_directly E_ACTIVE, dh
    
    ElseIf e = CAPACITY_CHECK Then
    
        logic_for_init_in_dtpicker_form WizardMain.E_DC_CW, e, dh, SRODEK
        
    ElseIf e = E_MRD_DATE Then
    
        logic_for_init_in_dtpicker_form E_DC_DATE, e, dh, SRODEK
        
    ElseIf e = E_MRD_REG_ROUTES Then
    
        logic_for_init_in_dtpicker_form E_DC_DATE, e, dh, SRODEK
        
    ElseIf e = E_PLATFORM Then
        
        logic_for_init_in_combobox_form e, dh, "", SRODEK
        
    ElseIf e = E_TRANSPORTATION_ACCOUNT_NUMBER Then
    
        logic_for_init_in_combobox_form e, dh, "", SRODEK
        
    ElseIf e = E_UNIQUE_ID Then
    
        logic_for_init_in_txtbox_form e, dh, "", ostatni
    
    Else
        MsgBox "to jest e poza zasiegiem - nigdy ten msgbox nie moze sie pojawic!"
    End If
    
    current_form.connect_with_details_handler dh
    current_form.show
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
End Sub

Private Sub logic_for_active_directly(ByRef e As E_NEW_PROJECT_ITEM, ByRef dh As DetailsHandler)


    ' rng_pointer.Offset(e - 2, 0).Value = ""
    rng_pointer.Offset(e - 1, 0).Value = POINTER
    
    Set current_form = WizardToggle
    WizardToggle.BtnPrev.Enabled = True
    WizardToggle.BtnNext.Enabled = True
    WizardToggle.BtnPrev.Visible = True
    WizardToggle.BtnNext.Visible = True
    WizardToggle.LabelQ.Caption = "Czy projekt jest aktywny: "
    ' WizardToggle.ToggleButtonInput.Value = 1
    ' WizardToggle.ToggleButtonInput.Caption = "GA"
    
    WizardToggle.e = E_ACTIVE
    WizardToggle.ToggleButtonInput.Caption = "YES"
    
    If new_definition Then
        WizardToggle.ToggleButtonInput.Value = 1
        WizardToggle.ToggleButtonInput.Caption = "YES"
    Else
    
        If rng_pointer.Offset(e - 1, -1).Value = "YES" Then
            WizardToggle.ToggleButtonInput.Value = 1
            WizardToggle.ToggleButtonInput.Caption = "YES"
            
        ElseIf rng_pointer.Offset(e - 1, -1).Value = "NO" Then
            WizardToggle.ToggleButtonInput.Value = 0
            WizardToggle.ToggleButtonInput.Caption = "NO"
            

        End If
        
    End If
    
    WizardToggle.LabelDesc.Caption = CStr(e)
    WizardToggle.LabelDesc.Visible = False

End Sub

Private Sub logic_for_biw_ga_directly(ByRef e As E_NEW_PROJECT_ITEM, ByRef dh As DetailsHandler)

    ' rng_pointer.Offset(e - 2, 0).Value = ""
    rng_pointer.Offset(e - 1, 0).Value = POINTER
    
    Set current_form = WizardToggle
    WizardToggle.BtnPrev.Enabled = True
    WizardToggle.BtnNext.Enabled = True
    WizardToggle.BtnPrev.Visible = True
    WizardToggle.BtnNext.Visible = True
    WizardToggle.LabelQ.Caption = "Wybierz typ projektu: "
    ' WizardToggle.ToggleButtonInput.Value = 1
    ' WizardToggle.ToggleButtonInput.Caption = "GA"
    
    WizardToggle.e = BIW_GA
    WizardToggle.ToggleButtonInput.Caption = "GA"
    
    If new_definition Then
        WizardToggle.ToggleButtonInput.Value = 1
        WizardToggle.ToggleButtonInput.Caption = "GA"
    Else
    
        If rng_pointer.Offset(e - 1, -1).Value = "BIW" Then
            WizardToggle.ToggleButtonInput.Value = 0
            WizardToggle.ToggleButtonInput.Caption = "BIW"
            
        ElseIf rng_pointer.Offset(e - 1, -1).Value = "GA" Then
            WizardToggle.ToggleButtonInput.Value = 1
            WizardToggle.ToggleButtonInput.Caption = "GA"
            

        End If
        
    End If
    
    WizardToggle.LabelDesc.Caption = CStr(e)
    WizardToggle.LabelDesc.Visible = False
End Sub


Private Sub logic_for_init_in_combobox_form(ByRef e As E_NEW_PROJECT_ITEM, _
    ByRef dh As DetailsHandler, def_txt As String, btn_label As E_DETAILS_WIZARD_ORDER)
    
    
    With WizardComboBox
    
        ' nic przed tym nie ma!
        ' rng_pointer.Offset(e - 2, 0).Value = ""
        rng_pointer.Offset(e - 1, 0).Value = POINTER
        Set current_form = WizardComboBox
        .BtnPrev.Enabled = True
        .BtnNext.Enabled = True
        
        ' najpierw na sile widoczny
        .BtnPrev.Visible = True
        
        ' a potem opcjonalnie chowamy
        If btn_label = SRODEK Then
            .BtnPrev.Visible = True
            
        ElseIf btn_label = PIERWSZY Then
            ' a potem opcjonalnie chowamy
            .BtnPrev.Visible = False
            .BtnPrev.Enabled = False
        End If
        .BtnNext.Visible = True
        
        If btn_label = ostatni Then
            .BtnNext.Caption = "Finish"
        Else
            .BtnNext.Caption = "Next"
        End If
        
        
        add_content_to_combobox_from_cache e
        
        
        .LabelQ.Caption = "Zdefiniuj " & CStr(rng_pointer.Offset(e - 1, -2).Value) & ": "
        
        If new_definition Then
            If CStr(Trim(def_txt)) = "" Then
                On Error Resume Next
                .ComboBoxInput.Value = CStr(.ComboBoxInput.List(1))
                On Error Resume Next
                .ComboBoxInput.Value = CStr(.ComboBoxInput.List(0))
            Else
                .ComboBoxInput.Value = CStr(def_txt)
            End If
        Else
            .ComboBoxInput.Value = rng_pointer.Offset(e - 1, -1).Value
        End If
        
        With .TextBoxBufor
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        .ComboBoxInput.SetFocus
        
        .LabelDesc.Caption = CStr(e)
        .LabelDesc.Visible = False
            
        
        
    End With
End Sub


Private Sub add_content_to_combobox_from_cache(e As E_NEW_PROJECT_ITEM)


    Dim c As Worksheet, cr As Range
    
    Dim proj As Range, faza As Range, platorm As Range
    Dim cont As Range
    
    Dim pod_txt As Range
    

    
    
    Set c = ThisWorkbook.Sheets(CACHE)
    
    
    Set pod_txt = c.Range("A2")
            
    Set proj = det.Range("project")
    Set faza = det.Range("faza")
    Set platform = det.Range("platform")
    
    With WizardComboBox
    
        .ComboBoxInput.Clear
        .ComboBoxInput.Value = ""
    
        
        ' 17th row
        If e = E_PLATFORM Then

            ' brakuje numerkow w wersji 3.8x
            If (UCase(Left(CStr(faza), 3)) = "PPV") Or (UCase(Left(CStr(faza), 5)) = "MVBNS") Then
            
                Set cr = c.Range("b2")
                
                If Trim(cr) = "" Then
                    .ComboBoxInput.AddItem G_NO_PLATFORMS_TXT
                Else
                
                    Do
                        .ComboBoxInput.AddItem cr
                        Set cr = cr.Offset(1, 0)
                    Loop Until Trim(cr) = ""
                
                End If
            Else
                .ComboBoxInput.AddItem G_NO_PLATFORMS_TXT
            End If
        
        
        ' 18th row
        ElseIf e = E_TRANSPORTATION_ACCOUNT_NUMBER Then
            
            
            ' faza zostala wybrana
            ' nazwa projektu
            ' platforma zotala wybrana
            ' maly calc do comboboxu i moze od razu tez value
            
            
            
            
            If Len(CStr(proj)) >= MIN_LEN_PROJ_NAME Then
            
            
                txt = ""
            
            
                If UCase(Left(CStr(faza), 3)) = "PPV" Then
                    Set cont = c.Range("E4")
                    
                    txt = cont.Offset(0, 2)
                    
                    If cont.Offset(0, 1) = "P" Then
                    
                        last_2_num_plat = Right(Split(CStr(platform), " ")(0), 2)
                        txt = Replace(txt, "qq", last_2_num_plat)
                        
                        txt = Replace(txt, "xxxxx", "0" & CStr(Left(proj, 4)))
                    End If
                    
                ' lefty sa po to gdyby po prawej stornie byly numerki
                ElseIf UCase(Left(CStr(faza), 5)) = "MVBNS" Then
                    Set cont = c.Range("E5")
                    
                    txt = cont.Offset(0, 2)
                    
                    If cont.Offset(0, 1) = "P" Then
                    
                        last_2_num_plat = Right(Split(CStr(platform), " ")(0), 2)
                        txt = Replace(txt, "qq", last_2_num_plat)
                        
                        txt = Replace(txt, "xxxxx", "0" & CStr(Left(proj, 4)))
                    End If
                    
                    
                Else
                    Set cont = c.Range("E6")
                    
                    txt = Left(proj, 4) & " " & CStr(faza)
                End If
                
                pod_txt.Value = CStr(txt)
            
                ' .ComboBoxInput.Clear
            
                If Trim(c.Range("A2")) = "" Then
                    .ComboBoxInput.AddItem G_NO_LIST_IN_CACHE
                Else
                
                    Set cr = c.Range("a2")
                    
                    Do
                        .ComboBoxInput.AddItem cr
                        Set cr = cr.Offset(1, 0)
                    Loop Until Trim(cr) = ""
                    
                    ' .ComboBoxInput.Value = .ComboBoxInput.List(1)
                End If
            End If
        
        Else
        
            MsgBox "ten msgbox nigdy nie moze sie pokazac: add_content_to_combobox_from_cache"
        
        End If
        
    End With
End Sub


Private Sub logic_for_init_in_txtbox_form(ByRef e As E_NEW_PROJECT_ITEM, _
    ByRef dh As DetailsHandler, def_txt As String, btn_label As E_DETAILS_WIZARD_ORDER)


    ' nic przed tym nie ma!
    ' rng_pointer.Offset(e - 2, 0).Value = ""
    rng_pointer.Offset(e - 1, 0).Value = POINTER
    Set current_form = WizardTxtBoxCurrent
    WizardTxtBoxCurrent.BtnPrev.Enabled = True
    WizardTxtBoxCurrent.BtnNext.Enabled = True
    
    ' najpierw na sile widoczny
    WizardTxtBoxCurrent.BtnPrev.Visible = True
    
    ' a potem opcjonalnie chowamy
    If btn_label = SRODEK Then
        WizardTxtBoxCurrent.BtnPrev.Visible = True
        
    ElseIf btn_label = PIERWSZY Then
        ' a potem opcjonalnie chowamy
        WizardTxtBoxCurrent.BtnPrev.Visible = False
        WizardTxtBoxCurrent.BtnPrev.Enabled = False
    End If
    WizardTxtBoxCurrent.BtnNext.Visible = True
    
    If btn_label = ostatni Then
        WizardTxtBoxCurrent.BtnNext.Caption = "Finish"
    Else
        WizardTxtBoxCurrent.BtnNext.Caption = "Next"
    End If
    
    WizardTxtBoxCurrent.LabelQ.Caption = "Zdefiniuj " & CStr(rng_pointer.Offset(e - 1, -2).Value) & ": "
    
    If new_definition Then
        WizardTxtBoxCurrent.TextBoxInput.Value = CStr(def_txt)
    Else
        WizardTxtBoxCurrent.TextBoxInput.Value = rng_pointer.Offset(e - 1, -1).Value
    End If
    
    With WizardTxtBoxCurrent.TextBoxBufor
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    On Error Resume Next
    WizardTxtBoxCurrent.TextBoxInput.SetFocus
    
    WizardTxtBoxCurrent.LabelDesc.Caption = CStr(e)
    WizardTxtBoxCurrent.LabelDesc.Visible = False
    
    
    
    If e = E_UNIQUE_ID Then
        With WizardTxtBoxCurrent.TextBoxInput
            .Enabled = False
            .Value = generate_unique_id(rng_pointer.Parent.Range("B1"))
        End With
    Else
        With WizardTxtBoxCurrent.TextBoxInput
            .Enabled = True
        End With
        
    End If
    
End Sub




Private Sub logic_for_init_in_dtpicker_form(dc As E_DATE_OR_CW, _
    ByRef e As E_NEW_PROJECT_ITEM, ByRef dh As DetailsHandler, btn_label As E_DETAILS_WIZARD_ORDER)

    With WizardDatePicker
    
        ' rng_pointer.Offset(e - 2, 0).Value = ""
        rng_pointer.Offset(e - 1, 0).Value = POINTER
    
        Set current_form = WizardDatePicker
        .BtnPrev.Enabled = True
        .BtnNext.Enabled = True
        
        If btn_label = PIERWSZY Then
            .BtnPrev.Visible = False
        ElseIf btn_label = SRODEK Then
            .BtnPrev.Visible = True
        Else
            .BtnPrev.Visible = True
        End If
        
        .CheckBoxDateAvail.Value = False
        
        .BtnNext.Visible = True
        If btn_label = ostatni Then
            .BtnNext.Caption = "Finish"
        Else
            .BtnNext.Caption = "Next"
        End If
        .LabelQ.Caption = "Zdefiniuj " & CStr(rng_pointer.Offset(e - 1, -2).Value) & ": "
    
        If dc = E_DC_CW Then
            
            If .CheckBoxDateAvail.Value = False Then
            
                If new_definition Then
                    ' nic nie musi - cala logika jest zaimplementowana w formularzu
                    ' .DTPickerInput
                Else
                    'Dim tmp As String
                    
                    If CStr(rng_pointer.Offset(e - 1, -1).Value) = TBD Then
                        .CheckBoxDateAvail.Value = True
                    Else
                        tmp = Replace(Replace(CStr(rng_pointer.Offset(e - 1, -1).Value), "CW", ""), "Y", "")
                        
                        ' Dim d As Date
                        If Len(tmp) = 6 Then
                            d = CDate(fnDateFromWeek(Left(tmp, 4), Right(tmp, 2), 2))
                            
                        ' opcja edycji
                        ElseIf Len(tmp) = 5 Then
                            MsgBox "jakims cudem pojawil sie schemat 5 elementowy - skontaktuj sie z autorem teraz!"
                            d = CDate(fnDateFromWeek(Left(tmp, 4), Right(tmp, 1), 2))
                        Else
                            MsgBox " nie ma takiego formatu: Y____CW__"
                        End If
                        .DTPickerInput.Value = d
                        If Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 2 Then
                            .LabelCW.Caption = "Y" & CStr(Year(d)) & "CW" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
                        ElseIf Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 1 Then
                            .LabelCW.Caption = "Y" & CStr(Year(d)) & "CW0" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
                        Else
                            MsgBox "details forms manager - logic for init in dtpicker - nie powinno sie pojawic!"
                        End If
                    End If
                End If
            Else
                
            End If
            
            
        ElseIf dc = E_DC_DATE Then
            
            
            If .CheckBoxDateAvail.Value = False Then
                
                If new_definition Then
                    ' nic nie musi - cala logika jest zaimplementowana w formularzu
                    ' WizardDatePicker.DTPickerInput
                Else
                    
                    If CStr(rng_pointer.Offset(e - 1, -1).Value) = TBD Then
                        .CheckBoxDateAvail.Value = True
                    Else
                        
                        If IsDate(rng_pointer.Offset(e - 1, -1).Value) Then
                            d = CDate(rng_pointer.Offset(e - 1, -1).Value)
                        Else
                            d = Date
                        End If
                        
                        WizardDatePicker.DTPickerInput.Value = d
                        If Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 2 Then
                            WizardDatePicker.LabelCW.Caption = "Y" & CStr(Year(d)) & "CW" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
                        ElseIf Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 1 Then
                            WizardDatePicker.LabelCW.Caption = "Y" & CStr(Year(d)) & "CW0" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
                        Else
                            MsgBox "CW01 a CW1 - nie powinno sie pokazac"
                        End If
                        ' WizardDatePicker.LabelCW.Caption = CStr(d)
                    End If
                End If
            End If
            
            
        End If
        
        
        
        WizardDatePicker.LabelDesc.Caption = CStr(e)
        WizardDatePicker.LabelDesc.Visible = False
    
    End With
End Sub


Private Sub clear_all_pointers()


    Dim tmp As Range
    
    ' zdefiniowany narazie poczatek
    Set tmp = rng_pointer
    
    Set tmp = rng_pointer.Parent.Range(rng_pointer, rng_pointer.Offset(0, -2).End(xlDown).Offset(0, 2))
    
    ' ThisWorkbook.Sheets(DETAILS_SHEET_NAME).Unprotect 123
    tmp.Clear
    ' ThisWorkbook.Sheets(DETAILS_SHEET_NAME).Protect 123
End Sub

