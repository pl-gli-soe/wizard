VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickupsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta powstala na potrzeby odzielenia interfejsu od
' logiki biznesowej - dzieki temu kod w formularzu bardzo sie odchudzil i zachowalem
' std ktory zalozylem we wszystkich wizardach :D


' po kilku przemyslaniach na tej klasie obsluze oprocz dodaj
' dodam rowniez logike edit - no i jak juz to moze tez delete?

' z racji tego ze mam juz gotowa zmienna pod przyjecie pracy na form
' current_form - to moge miec je dwie rozne
' a chodzi wlasciwie ino o adjust i sub edit zamiast dodaj
'
' no tak tak
' teraz to juz w ogole jest jazda
' pod current form raz jest FormEditPickups a raz FormPickups (to jest do dodawania tylko)

' aktualnie jest juz ponad 1000 linii wiec nie malo :)

' udpate z grudnia 2015 - jest juz ponad 1500 linii


Private current_form As Variant


Private e As E_MASTER_MANDATORY_COLUMNS
Private column_source As E_MASTER_MANDATORY_COLUMNS

Private r As Range
Private czy_moge_dodac_nowy_pickup As Boolean
Private czy_chociaz_raz_pickup_zostal_dodany As Boolean
Private czy_pojawily_sie_zerowe_badz_ujemne_pusy As Boolean
Private silowy_duplikat_nazwy_pusa As Boolean


Private tmp_for_pn As String
Private tm_for_duns As String
Private tmp_for_fupcode As String


Private ghost_activated As Boolean

' =========================================================

Public co_poszlo_zle As String
Private error_sh As Worksheet



' =========================================================
' synchro z jenny dane z jenny
Private jr As Range
' =========================================================

Public Sub setJr(r As Range)
    Set jr = r
    
    ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Activate
End Sub



Public Sub connect_with_form_pickups(e As E_ADD_EDIT_PUSES, arg As Variant)
    ' okazuje sie ze pierwszy argument narazie jest obsoletowy i super nie potrzebny
    Set current_form = arg
End Sub


Public Sub hide_edit_thing_and_prepare_for_adding()
    ' w sumie tak jak w konstruktorze
    Set current_form = Nothing
    co_poszlo_zle = ""
End Sub


Public Sub fill_labels()
    
    
    ' current_form.name = FormPickups
    ' DTPickerDUNS_DeliveryDate
    
    ' current_form.name = FormEditPuses
    ' ListBoxDelDate
    
    With current_form
        .LabelDelConf.Caption = WizardMain.PREFIX_TXT_EDIT_PUS_DEL_CONF & CStr(Cells(ActiveCell.Row, WizardMain.Delivery_confirmation))
        ' obsolete - ten element juz nie istnieje
        ' .LabelFirstPUSDate.Caption = _
        ' WizardMain.PREFIX_TXT_EDIT_PUS_FST_PUS_DATE & CStr(Cells(ActiveCell.Row, WizardMain.First_Confirmed_PUS_Date))
        
        If .Name = "FormPickups" Then
        
            If match_cw(CStr(Cells(ActiveCell.Row, WizardMain.Delivery_confirmation)), CStr(.DTPickerDUNS_DeliveryDate.Value)) Then
                .LabelDelConf.ForeColor = RGB(0, 0, 0)
            Else
                .LabelDelConf.ForeColor = RGB(255, 0, 0)
                
                'answer = MsgBox("Delivery Date przekracza Del Conf, czy dopasowac Del Conf?", vbYesNo)
                '
                'If answer = vbYes Then
                '
                '
                '    For q = 0 To .ListBoxPNsFromThisDuns.ListCount - 1
                '
                '       ' ListBoxIndx
                '        pn_for_fill_labels = Split(.ListBoxPNsFromThisDuns.List(q), WizardMain.COMMA)(0)
                '
                '       Set r = Range(Cells(1, WizardMain.pn), Cells(WizardMain.POLOWA_CAPACITY_ARKUSZA, WizardMain.pn))
                '        Set r = r.Find(pn_for_fill_labels)
                '
                '        If Not r Is Nothing Then
                '
                '            nowa_wartosc_w_del_conf_range CDate(.DTPickerDUNS_DeliveryDate.Value), Cells(ActiveCell.Row, WizardMain.Delivery_confirmation)
                '        End If
                '    Next q
                'End If
                
            End If
        ElseIf .Name = "FormEditPuses" Then
        
            .LabelDelConf.ForeColor = RGB(0, 0, 0)
            v = ""
        
            For x = 0 To .ListBoxDelDate.ListCount - 1
                v = .ListBoxDelDate.List(x)
                
                If v <> "" Then
                    If match_cw(CStr(Cells(ActiveCell.Row, WizardMain.Delivery_confirmation)), CStr(v)) Then
                        .LabelDelConf.ForeColor = RGB(0, 0, 0)
                    Else
                    

                        .LabelDelConf.ForeColor = RGB(255, 0, 0)
                        
                        'Debug.Print CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(v))))
                        'Debug.Print CStr(Cells(ActiveCell.Row, WizardMain.Delivery_confirmation))
                        'Debug.Print CStr(.ListBoxPUSes.List(x))
                        
                        
                        'answer = MsgBox("Delivery Date (CW: " & _
                        '    CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(v)))) & _
                        '    ") przekracza Del Conf: " & _
                        '    CStr(Cells(ActiveCell.Row, WizardMain.Delivery_confirmation)) & _
                        '    ", czy dopasowac Del Conf dla " & _
                        '    CStr(.ListBoxPUSes.List(x)) & _
                        '    "?", vbYesNo)
                
                        'If answer = vbYes Then
                        '    For q = 0 To .ListBoxIndx.ListCount - 1
                        '
                        '        ' ListBoxIndx
                        '        pn_for_fill_labels = Split(.ListBoxIndx.List(q), WizardMain.COMMA)(0)
                        '
                        '        Set r = Range(Cells(1, WizardMain.pn), Cells(WizardMain.POLOWA_CAPACITY_ARKUSZA, WizardMain.pn))
                        '        Set r = r.Find(pn_for_fill_labels)
                        '
                        '        If Not r Is Nothing Then
                        '
                        '            nowa_wartosc_w_del_conf_range CDate(v), Cells(r.Row, WizardMain.Delivery_confirmation)
                        '        End If
                        '    Next q
                        'End If

                        
                        
                        ' Exit For
                    End If
                End If
                
            Next x
            
            
        Else
            ' no operation
        End If
    End With
End Sub


Private Sub nowa_wartosc_w_del_conf_range(dd As Date, ByRef r As Range)
    
    tmp_dd = CDate(dd)
    
    tmp_del_conf_arr = Split(r, " ")
    
    tmp_txt = ""
    For x = LBound(tmp_del_conf_arr) To UBound(tmp_del_conf_arr) - 1
        tmp_txt = tmp_txt & tmp_del_conf_arr(x)
    Next x
    
    cw = CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(tmp_dd)))
    
    If Len(cw) = 1 Then
        cw = "0" & CStr(cw)
    End If

    tmp_txt = tmp_txt & " Y" & CStr(Year(tmp_dd)) & "CW" & CStr(cw)
    
    r.Value = tmp_txt
    
    
    ' teraz pozostale - dla tego PUS'a
    ' nowy_schemat_offsetu_w_arkuszu_pickups
    
    
End Sub

Private Function match_cw(ycw As String, del As String) As Boolean
    match_cw = False
    
    If IsDate(del) Then
    
        tmp = Replace(Replace(ycw, "Y", ""), "CW", "")
        
        If Len(Trim(tmp)) = 5 Then
            tmp = Left(tmp, 4) & "0" & Right(tmp, 1)
        End If
        
        If Len(Trim(tmp)) = 6 Then
        
            cw = CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(del))))
            If Len(cw) = 1 Then
                cw = "0" & CStr(cw)
            End If
            
            int_del = CStr(Year(CDate(del))) & CStr(cw)
        
            If CLng(int_del) <= CLng(tmp) Then
            
                match_cw = True
            End If
            
        End If
    End If
End Function


Public Sub fill_edit_listboxes()


    Dim psh As Worksheet, i  As Range, arr(1 To 256) As String, msh As Worksheet
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    
    

    With current_form
        ' z tych elementow to w sumie do uzupelnienia tylko ListBoxPuses:
        .ListBoxPUSes.Clear
        .ListBoxPickupDate.Clear
        .ListBoxDelDate.Clear
        .ListBoxIndx.Clear
        .ListBoxQty.Clear
        
        str_pn = CStr(msh.Cells(ActiveCell.Row, WizardMain.pn))
        str_duns = CStr(msh.Cells(ActiveCell.Row, WizardMain.duns))
        str_fup = CStr(msh.Cells(ActiveCell.Row, WizardMain.fup_code))
        str_indx = str_pn & "," & str_duns & "," & str_fup
        pus_name = ""
        Set i = psh.Range("A2")
        Do
            
            ' sprawdz czy indx jest matched z odpowiednimi danymi
            ' ----------------------------------------------------------
            If CStr(Trim(str_indx)) = CStr(Trim(i)) Then
                pus_name = CStr(i.Offset(0, 7))
                
                If sprawdz_czy_mozesz_dodac_czy_lista_pusta_czy_moze_jest_juz_duplikat(current_form, CStr(pus_name)) Then
                    .ListBoxPUSes.AddItem pus_name
                    ' uzupelnij_indx_i_qty CStr(pus_name), .ListBoxIndxQty, psh
                    .ListBoxPickupDate.AddItem Format(CStr(i.Offset(0, 4)), "yyyy-mm-dd")
                    .ListBoxDelDate.AddItem Format(CStr(i.Offset(0, 5)), "yyyy-mm-dd")
                    ' .ListBoxIndxQty.AddItem CStr(i.Offset(0, 6))
                    ' Exit Do
                End If
            End If
            
            ' global sub
            
            WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups i
        Loop Until i.Row = WizardMain.CAPACITY_ARKUSZA
    End With
End Sub

Private Function sprawdz_czy_mozesz_dodac_czy_lista_pusta_czy_moze_jest_juz_duplikat(ByRef cf As Variant, pus_name As String) As Boolean
    sprawdz_czy_mozesz_dodac_czy_lista_pusta_czy_moze_jest_juz_duplikat = True
    
    With cf
        If .ListBoxPUSes.ListCount = 0 Then
            sprawdz_czy_mozesz_dodac_czy_lista_pusta_czy_moze_jest_juz_duplikat = True
        ElseIf .ListBoxPUSes.ListCount > 0 Then
        
            For x = 0 To .ListBoxPUSes.ListCount - 1
            
                If .ListBoxPUSes.List(x) = pus_name Then
                    sprawdz_czy_mozesz_dodac_czy_lista_pusta_czy_moze_jest_juz_duplikat = False
                    Exit Function
                End If
                
            Next x
        End If
    End With
End Function

Public Sub disable_checkbox()
    current_form.ComboBoxPN.Enabled = False
    current_form.ComboBoxSourceDUNS.Enabled = False
End Sub

Public Sub fill_source_checkbox()



    With current_form.ComboBoxPN
    
        ' ----------------------------------------------
        
        ' Me.ComboBoxSource.Clear
        ' Me.ComboBoxSource.AddItem Item
        
        
        .Clear
        e = Total_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        e = MRD1_Ordered_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        e = MRD2_Ordered_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        
        .AddItem ALL_ORDERED_QTY
        
        .AddItem GHOST
        
        
        .Value = ALL_ORDERED_QTY
    End With
    
    
    With current_form.ComboBoxSourceDUNS
    
        ' ----------------------------------------------
        
        ' Me.ComboBoxSource.Clear
        ' Me.ComboBoxSource.AddItem Item
        
        
        .Clear
        e = Total_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        e = MRD1_Ordered_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        e = MRD2_Ordered_QTY
        .AddItem CStr(ThisWorkbook.Sheets(MASTER_SHEET_NAME).Cells(1, Int(e)).Value)
        
        .AddItem ALL_ORDERED_QTY
        
        .AddItem GHOST
        
        
        .Value = ALL_ORDERED_QTY
    End With
    
    
    e = MRD2_Ordered_QTY
    column_source = MRD2_Ordered_QTY
        ' ----------------------------------------------
    
    
End Sub


Public Sub listbox_clicked()
    Me.connect_with_form_pickups E_EDIT, FormEditPuses
    
    With current_form
    
    
        For x = 0 To .ListBoxIndx.ListCount - 1
            .ListBoxIndx.Selected(x) = False
        Next x
    
        For x = 0 To .ListBoxQty.ListCount - 1
            
                If .ListBoxQty.Selected(x) Then
                           
                     .ListBoxIndx.Selected(x) = True
                   
                End If
            Next x
    End With
End Sub

Public Sub remove_this_pn()

    Me.connect_with_form_pickups E_EDIT, FormEditPuses
    
    
    Dim czy_zostalo_cos_usuniete As Boolean
    czy_zostalo_cos_usuniete = False
    
    Dim psh As Worksheet, r As Range, tmp As Range
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set r = psh.Cells(2, 1)
    
    
    With current_form
        
        For x = 0 To .ListBoxIndx.ListCount - 1
            
            If .ListBoxIndx.Selected(x) = True Then
            
                Do
                    If r = .ListBoxIndx.List(x) Then
                        ' Set tmp = r
                        ' Set r = r.Offset(1, 0)
                        For y = WizardMain.O_INDX To WizardMain.O_PUS_Number
                            psh.Cells(r.Row, y).Value = ""
                        Next y
                        'tmp.EntireRow.Delete
                        ' ThisWorkbook.Save
                        czy_zostalo_cos_usuniete = True
                    End If
                    
                    WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
                Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
                
            
                .ListBoxIndx.RemoveItem x
                .ListBoxQty.RemoveItem x
            End If
        Next x
        
    End With
    
    If czy_zostalo_cos_usuniete Then
        MsgBox "dane zostaly usuniete!"
    Else
        MsgBox "nie znaleziono danych do usuniecia"
    End If
    
    
    Set pickups_handler = New PickupsHandler
End Sub

Public Sub zmniejsz_liste_indx_poprzez_ptrn()

    pus_listbox_change
End Sub


Public Function edit_qty_key_pressed(ByVal KeyAscii As MSForms.ReturnInteger) As Boolean


    edit_qty_key_pressed = False


    If (((Int(KeyAscii) >= ASCII_0) And Int(KeyAscii) <= ASCII_9)) Or (Int(KeyAscii) = ASCII_ENTER) Then
    

        Dim te As Range
        Dim tmp_str As String
        tmp_str = ""
        ' TOP_EDIT_LIMIT
        Set te = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("top_edit").Offset(-1, 0)
        
        ' przesun teraz tyle ile trzeba
        If te.Offset(1, 0) <> "" Then
            Set te = te.End(xlDown).Offset(1, 0)
        Else
            Set te = te.Offset(1, 0)
        End If
    
        
        Me.connect_with_form_pickups E_EDIT, FormEditPuses
        
        With current_form
            For x = 0 To .ListBoxQty.ListCount - 1
            
                If .ListBoxQty.Selected(x) Then
                    ' MsgBox Chr(Int(KeyAscii))
                    If Int(KeyAscii) = ASCII_ENTER Then
                    
                        edit_qty_key_pressed = True
                        
                        te = ENTER_STR
                        Set te = te.Offset(-1, 0)
                        
                        If te = ENTER_STR Then
                            .ListBoxQty.List(x) = "Q: ERROR"
                            Exit For
                        End If
                        
                        Do
                            tmp_str = te & tmp_str
                            Set te = te.Offset(-1, 0)
                        Loop Until te = ENTER_STR
                        
                        If Len(tmp_str) > SIX Then
                            MsgBox "Qty for this PN is bigger than 999999 - this case is not possible"
                            .ListBoxQty.List(x) = "Q: ERROR"
                            Exit For
                        End If
                        
                        
                        If te.Row > TOP_EDIT_LIMIT Then
                            Set te = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("top_edit")
                            Set te = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range(te, te.End(xlDown))
                            te.Clear
                        End If
                        
                        .ListBoxQty.List(x) = "Q: " & tmp_str
                        Exit For
                    Else
                        
                        te = Chr(KeyAscii)
                        Exit For
                        
                    End If
                End If
            Next x
        End With
    Else
        MsgBox "nie wpisywac czegos co nie jest calkowita liczba dodatnia!"
    End If
End Function

Public Sub add_form_listbox_qty_sth_selected()

    Me.connect_with_form_pickups E_ADD, FormPickups
    With current_form
        

        For x = 0 To .ListBoxCurrPusQty.ListCount - 1
        
            If .ListBoxCurrPusQty.Selected(x) = True Then
                .ListBoxDone.Selected(x) = True
                .ListBoxSource.Selected(x) = True
                .ListBoxPNsFromThisDuns.Selected(x) = True
            Else
                .ListBoxDone.Selected(x) = False
                .ListBoxSource.Selected(x) = False
                .ListBoxPNsFromThisDuns.Selected(x) = False
            End If
            
        Next x
    End With
End Sub

Public Sub pus_listbox_change()
    Me.connect_with_form_pickups E_EDIT, FormEditPuses
    
    ' no tutaj jak wariat nie bede powtarzal warunkow
    ' --------------------------------------------------
    
    
    Dim psh As Worksheet, i  As Range, arr(1 To 256) As String, msh As Worksheet
    Dim ptrn As String
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    
    ptrn = ""
    

    With current_form
        ' z tych elementow to w sumie do uzupelnienia tylko ListBoxPuses:
        '.ListBoxPUSes
        '.ListBoxPickupDate
        '.ListBoxDelDate
        .ListBoxIndx.Clear
        .ListBoxQty.Clear
        
        
        If .TextBox1.Value <> "" Then
            ptrn = CStr(.TextBox1.Value)
        End If
        
        ' usun wszystkie selekcje na not enabled fields
        For x = 0 To .ListBoxPUSes.ListCount - 1
        
            .ListBoxPickupDate.Selected(x) = False
            .ListBoxDelDate.Selected(x) = False
            
        Next x
        
        pus_name = ""
        With .ListBoxPUSes
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    pus_name = .List(x)
                    current_form.ListBoxPickupDate.Selected(x) = True
                    current_form.ListBoxDelDate.Selected(x) = True
                    uzupelnij_indx_i_qty CStr(pus_name), current_form.ListBoxIndx, current_form.ListBoxQty, psh, ptrn
                    Exit For
                End If
            Next x
        End With
    End With
    
    
    ' --------------------------------------------------
End Sub


Private Sub uzupelnij_indx_i_qty(s As String, lb_indx As Variant, lb_qty As Variant, ByRef psh As Worksheet, prtn As String)


    Dim i  As Range
    Set i = psh.Range("A2")
        Do
            
            ' sprawdz czy indx jest matched z odpowiednimi danymi
            ' ----------------------------------------------------------
            If CStr(Trim(s)) = CStr(Trim(i.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX))) Then
                ' indx = CStr(i)
                If prtn = "" Then
                    innner_uzupelnij_indx_i_qty lb_indx, lb_qty, i
                ElseIf CStr(i) Like "*" & CStr(prtn) & "*" Then
                    innner_uzupelnij_indx_i_qty lb_indx, lb_qty, i
                End If
                ' Exit Do
            End If
            
            WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups i
        Loop Until i.Row = WizardMain.CAPACITY_ARKUSZA

End Sub

Private Sub innner_uzupelnij_indx_i_qty(lb_indx As Variant, lb_qty As Variant, ByRef i As Range)


    ' omijamy zera - bo i po co - ale to narazie koncepcja - moze jednak chce zeby tak sie dzialo
    ' zeby bylo widac slad po zerowaniu
    'If CStr(i.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX)) <> "0" Then
        lb_indx.AddItem CStr(i)
        lb_qty.AddItem "Q: " & CStr(i.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX))
    'End If
End Sub



Public Function edit_puses()



    ' ta funkcja zajmuje sie tylko edycja qty
    ' nazwa i daty leca zmiany tylko i wylacznie pod stronie mniejszych formularzy gdzie ma pomocniczej formuly
    ' w klasie pickup handler

    edit_puses = -1
    
    CzekajForm.show vbModeless
    
    Me.connect_with_form_pickups E_EDIT, FormEditPuses
    
    
    If ActiveCell.Row > 1 Then
        If Trim(Cells(ActiveCell.Row, Int(WizardMain.pn)).Value) = "" Then
            MsgBox "nie mozna edytowac nieistniejacego pusa"
        Else
        
            ' kilka warunkow zostalo spelnionych
            ' ------------------------------------------------
            
            Dim psh As Worksheet, i  As Range, arr(1 To 256) As String, msh As Worksheet
            Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
            Set msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
            
            
            
            With current_form
            
                ' znajdz nazwe pickup'a do update'u
                ' ------------------------------------
                pus_name = ""
                With .ListBoxPUSes
                    For x = 0 To .ListCount - 1
                        If .Selected(x) Then
                            pus_name = .List(x)
                            edit_puses = x
                            Exit For
                        End If
                    Next x
                End With
                ' ------------------------------------
                
                
                ' sciagamy dane z formularza i wrzucamy do arksuza
                ' PICKUPS
                ' -------------------------------------------------------
                ' warunek sprawdzajacy, czy w ogole cos znalezlismy
                ' i czy ma sens by sie bawic w pierdy kolejne loopowe
                If pus_name <> "" Then
                    Set i = psh.Range("a2") ' o_indx
                    
                    Do
                        ' body loop
                        ' ================================================================
                        
                        ' match na nazwie pusa (pus#)
                        If CStr(i.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX)) = pus_name Then
                            
                            ' teraz na tym wierszu znajdziemy co mamy updateowac
                            For x = 0 To .ListBoxIndx.ListCount - 1
                                If CStr(i) = .ListBoxIndx.List(x) Then
                                    i.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX) = Replace(.ListBoxQty.List(x), "Q: ", "")
                                End If
                            Next x
                        End If
                        
                        
                        
                        ' ================================================================
                        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups i
                    Loop Until i.Row = WizardMain.CAPACITY_ARKUSZA
                End If
                
                
                
                ' -------------------------------------------------------
            End With
            
            ' ------------------------------------------------
        End If
    End If
    
    
    CzekajForm.hide
End Function

Public Sub dodaj()
    czy_moge_dodac_nowy_pickup = True
    czy_chociaz_raz_pickup_zostal_dodany = False
    czy_pojawily_sie_zerowe_badz_ujemne_pusy = False
    
    
    CzekajForm.show vbModeless
    
    Me.connect_with_form_pickups E_ADD, FormPickups

    With current_form.MultiPage
        If .SelectedItem.Name = .Pages.item("PagePN").Name Then
            
            ' adding only one part number
            ' =========================================
            
            ' 1. check if there's a place for this order
            ' (duplicates) and optionally notify user that he did some duplicates
            
            If Trim(Cells(ActiveCell.Row, Int(WizardMain.pn)).Value) = "" Then
                MsgBox "nie mozna dodac zamowienia dla pustej wartosci part #"
            Else
                MsgBox "nie ma jeszcze gotowej implementacji"
                Exit Sub
                ' dodaj_pickup_gdy_zakladka_pn_jest_aktywna r
            End If
            
            ' =========================================
        ElseIf .SelectedItem.Name = .Pages.item("PageDUNS").Name Then
        
            ' adding many PN from one DUNS
            ' =========================================
            
            ' need to think how to do it
            ' but for sure there is a necessity to have notification about duplicates
            
            ' zapobiega wszelkim problemom
            dodaj_pickup_gdy_zakladka_duns_jest_aktywna r
            
            
        End If
    End With
    
    
    ' ThisWorkbook.Save
    
    adjust_content_if_selection_changed


    CzekajForm.hide
End Sub

Private Sub dodaj_pickup_gdy_zakladka_duns_jest_aktywna(ByRef r As Range)


    If ActiveSheet.Name = MASTER_SHEET_NAME Then
        
        If Trim(Cells(ActiveCell.Row, Int(WizardMain.pn)).Value) = "" Then
            MsgBox "nie mozna dodac zamowienia dla pustej wartosci part #"
        Else
        

                ' current_form.ListBoxPNsFromThisDuns
                czy_moge_dodac_nowy_pickup = True
                czy_chociaz_raz_pickup_zostal_dodany = False
                ' miejsce puste ustawia sie na referencji r
                czy_moge_dodac_nowy_pickup = zweryfikuj_arkusz_pickups_i_ustaw_na_pierwsze_puste_miejsce(r)
                czy_moge_dodac_nowy_pickup = zweryfikuj_poprawnosc_dat_w_formie()
                
                
                ' olewam ta logike - niech dziala tylko przy labelce del confa
                ' args: Variant & Range
                'sprawdz_del_conf_del_date_danego_pusa CStr(current_form.DTPickerDUNS_DeliveryDate.Value), _
                '    Cells(ActiveCell.Row, Int(WizardMain.Delivery_confirmation))
                
                silowy_duplikat_nazwy_pusa = False
                silowy_duplikat_nazwy_pusa = sprawdz_czy_nie_ma_takiego_pusa_juz()
                czy_moge_dodac_nowy_pickup = Not silowy_duplikat_nazwy_pusa
                    
                If czy_moge_dodac_nowy_pickup Then
                    For x = 0 To current_form.ListBoxPNsFromThisDuns.ListCount - 1
   
                            'r.Parent.Unprotect 123
                            ' tutaj logika share - zeby ustawic pozycje na arkuszu pickups
                            If ustaw_tylko_na_pierwsze_puste_miejsce(r) Then
                            
                                If IsNumeric(current_form.ListBoxCurrPusQty.List(x)) Then
                                    If CLng(current_form.ListBoxCurrPusQty.List(x)) > 0 Then
                                    
                                        czy_chociaz_raz_pickup_zostal_dodany = True
                                    
                                        r = CStr(current_form.ListBoxPNsFromThisDuns.List(x))
                                        ' ThisWorkbook.Save
                                        arr = Split(r, ",")
                                        
                                        ' dane z splitu
                                        r.Offset(0, WizardMain.O_PN - WizardMain.O_INDX) = arr(LBound(arr))
                                        r.Offset(0, WizardMain.O_DUNS - WizardMain.O_INDX) = arr(LBound(arr) + 1)
                                        r.Offset(0, WizardMain.O_FUP_code - WizardMain.O_INDX) = arr(UBound(arr))
                                        
                                        ' pickup details
                                        r.Offset(0, WizardMain.O_Pick_up_date - WizardMain.O_INDX) = current_form.DTPickerDUNS_PickUpDate.Value
                                        r.Offset(0, WizardMain.O_Delivery_Date - WizardMain.O_INDX) = current_form.DTPickerDUNS_DeliveryDate.Value
                                        If IsNumeric(current_form.ListBoxCurrPusQty.List(x)) Then
                                            r.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX) = CLng(current_form.ListBoxCurrPusQty.List(x))
                                        Else
                                            MsgBox "FATAL ERROR! Nigdy nie powinien przypisac zerowego pusa!" & Chr(10) & _
                                                "dodaj pickup gdy zakladka duns jest aktywna - klasa pickups handler!"
                                            
                                            ' r.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX) = CLng(0)
                                        End If
                                        
                                        
                                        If Trim(current_form.TextBoxMaskForPusNumber.Value) = "" Then
                                            r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX) = "NULL"
                                        Else
                                            r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX) = CStr(current_form.TextBoxMaskForPusNumber)
                                        End If
                                    Else
                                        czy_pojawily_sie_zerowe_badz_ujemne_pusy = True
                                    End If
                                End If
                            Else
                            
                                ' tutaj jesli pus jest duzy moze sie zdarzyc fluktuacji tylko czesci dodania danych
                                ' trzeba dotestowac!
                                MsgBox "synchronizacja uzytkownikow nie przebiegla pomyslnie - nie dodano nowego PUSa!"
                                Exit Sub
                            End If
                            
                            'r.Parent.Protect 123
                    Next x
                    
                    
                    If czy_chociaz_raz_pickup_zostal_dodany Then
                        If czy_pojawily_sie_zerowe_badz_ujemne_pusy Then
                            MsgBox "Dodane zostaly tylko pusy z ilosciami wiekszymi od zera!"
                        Else
                            MsgBox "Wszystko dodane!"
                        End If
                    Else
                        MsgBox "zaden pickup nie zostal dodany - warunki nie zostaly spelnione - upewnij sie ze lista ma sens"
                    End If
                Else
                    MsgBox "zaden pickup nie zostal dodany - duplikat nazwy pusa prawdopodobnie jest problemem"
                End If
            
        End If
    Else
        MsgBox "nie mozesz przejsc na ten arkusz, gdy aktywny jest formularz dodawania pickupow z poziomu numeru DUNS!"
        ThisWorkbook.Sheets(MASTER_SHEET_NAME).Activate
    End If
    
    ' =========================================
End Sub

Private Sub sprawdz_del_conf_del_date_danego_pusa(del_date As Variant, del_conf As Range)

    ' del_date ze niby w formacie yyyy-mm-dd
    ' del_conf std delivery confirmation
    
    ' podstawowe kryterium by moc rozpoczac prace weryfikacji i ewentualne wyswietlenia alertu
    If IsDate(del_date) And (CStr(del_conf) Like "*Y*CW*") Then
    
        ' no to mamy date i jest poprawnosc del confa ze posiada tam date
        
        prsd_del_conf_txt = CStr(wyciagnij_date_z_del_conf(CStr(del_conf)))
        
        If prsd_del_conf_txt <> "" Then
            
            ' jakas data zostala wyparsowana
            ' ================================
            
            Dim dd As Date
            dd = CDate(del_date)
            
            y = Year(dd)
            cw = CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(dd)))
            
            If Len(cw) = 1 Then
                cw = "0" & CStr(cw)
            End If
            
            ycw = CStr(y) & CStr(cw)
            
            wynik = CLng(ycw) - CLng(prsd_del_conf_txt)
            
            If wynik > 0 Then
                
                msgbox_answer = MsgBox("Del date (PUS): " & CStr(dd) & ", Del Conf: " & CStr(del_conf) & ", czy chcesz dostosowac Del Confa?", _
                    vbYesNo, "Twardy Alert")
                    
                    
                If msgbox_answer = vbYes Then
                    
                    ' del conf jest jako ref arg
                    nowa_wartosc_w_del_conf_range dd, del_conf
                        
                End If
                
            ElseIf wynik = 0 Then
                MsgBox "FYI - delivery date PUSa pokrywa sie DEL CONF"
            Else
            End If
            
            ' ================================
        End If
    End If
End Sub



Private Function wyciagnij_date_z_del_conf(dc As String) As String
    wyciagnij_date_z_del_conf = ""
    
    
    tmp = Right(dc, 9)
    
    ' czyli wszystko jest OK i lecimy dalej
    If Left(tmp, 1) = "Y" Then
    ' cos chyba CW jest za krotkie
    ElseIf Mid(tmp, 2, 1) = "Y" Then
        tmp = Right(dc, 8)
        
        tmp = Left(tmp, 7) & "0" & Right(tmp, 1)
    Else
        Exit Function
    End If
    
    tmp = Replace(Replace(tmp, "Y", ""), "CW", "")
    
    wyciagnij_date_z_del_conf = tmp
    
End Function

Private Function sprawdz_czy_nie_ma_takiego_pusa_juz() As Boolean
    sprawdz_czy_nie_ma_takiego_pusa_juz = False
    
    
    With ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)


        
        Set r = .Cells(1, WizardMain.O_INDX)
        Do
            If CStr(r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX)) = _
                current_form.TextBoxMaskForPusNumber.Value Then
                
                sprawdz_czy_nie_ma_takiego_pusa_juz = True
                MsgBox "taki pus zostal juz zdefiniowany - operacja przerwana!"
                Exit Function
            End If
            WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
        Loop Until r.Row = WizardMain.CAPACITY_ARKUSZA

    End With
    
End Function

Private Function zweryfikuj_poprawnosc_dat_w_formie() As Boolean

    zweryfikuj_poprawnosc_dat_w_formie = True
    
    With current_form
    
        Dim pdate As Date
        Dim ddate As Date
        
        pdate = CDate(.DTPickerDUNS_PickUpDate.Value)
        ddate = CDate(.DTPickerDUNS_DeliveryDate.Value)
        
        If ddate < pdate Then
            zweryfikuj_poprawnosc_dat_w_formie = False
            MsgBox "Daty nie sa chronologicznie!"
        End If
    End With
End Function

Private Function ustaw_tylko_na_pierwsze_puste_miejsce(ByRef r As Range) As Boolean


    ' na poczatku jestesmy optymistami
    ' potem atakujemy ta flage czy cos poszlo nie tak
    ustaw_tylko_na_pierwsze_puste_miejsce = True
    '    MsgBox Application.UserName
    '
    '    Users = ThisWorkbook.UserStatus
    '    With ThisWorkbook.ActiveSheet
    '        For Row = 1 To UBound(Users, 1)
    '            .Cells(Row, 1) = Users(Row, 1)
    '            .Cells(Row, 2) = Users(Row, 2)
    '            Select Case Users(Row, 3)
    '                Case 1
    '                    .Cells(Row, 3).Value = "Exclusive"
    '                Case 2
    '                    .Cells(Row, 3).Value = "Shared"
    '            End Select
    '        Next
    '    End With

    With ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
        
        a = vbYes
        
        ' gdzie zaczynamy
        ' G_STEP_BETWEEN_PARALELL_USERS
        gdzie_zaczynamy = 1
        Users = ThisWorkbook.UserStatus
        
        If USERS_LIMIT < UBound(ThisWorkbook.UserStatus, 1) Then
            ' users_status_usun_moje_stare_instancje CStr(Application.UserName)
            MsgBox "przekroczono limit uzytkownikow pliku - sprawdz liste w Review -> Share Workbook"
            ustaw_tylko_na_pierwsze_puste_miejsce = False
            Exit Function
        End If
        
        For x = 1 To UBound(ThisWorkbook.UserStatus, 1)
            If CStr(Application.UserName) = CStr(Users(x, 1)) Then
                gdzie_zaczynamy = (G_STEP_BETWEEN_PARALELL_USERS * (x - 1)) + 1
                ustaw_tylko_na_pierwsze_puste_miejsce = True
            End If
        Next x
        
        
        If gdzie_zaczynamy > WizardMain.POLOWA_CAPACITY_ARKUSZA Then
            MsgBox "przekroczylismy mozliwy zakres arkusza - dodanie PUSa jest niemozliwe!"
            ustaw_tylko_na_pierwsze_puste_miejsce = False
            Exit Function
        End If
        
        Set r = .Cells(gdzie_zaczynamy, WizardMain.O_INDX)
        If r = "" Then
            ' no operation
        ElseIf r.Offset(1, 0) = "" Then
        
            Set r = r.Offset(1, 0)
        Else
            
            Do
                Set r = r.Offset(1, 0)
                
                If r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA Then
                    MsgBox "przekroczylismy mozliwy zakres arkusza - dodanie PUSa jest niemozliwe!"
                    ustaw_tylko_na_pierwsze_puste_miejsce = False
                    Exit Function
                End If
                
            Loop Until Trim(r) = ""
            
            
        End If
        
    End With
End Function


' tutaj jeszcze i tak nie ma znaczenia ustawienie r wiec zostawie jak jest
Private Function zweryfikuj_arkusz_pickups_i_ustaw_na_pierwsze_puste_miejsce(ByRef r As Range) As Boolean


    With ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)


        zweryfikuj_arkusz_pickups_i_ustaw_na_pierwsze_puste_miejsce = True
        a = vbYes
        
        
        Set r = .Cells(1, WizardMain.O_INDX)
        Do
            If CStr(r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX)) = _
                current_form.TextBoxMaskForPusNumber.Value Then
                
                a = vbNo
                zweryfikuj_arkusz_pickups_i_ustaw_na_pierwsze_puste_miejsce = False
                Exit Function
            End If
            Set r = r.Offset(1, 0)
        Loop Until r = ""
        
        Set r = .Cells(1, WizardMain.O_INDX)
        If a = vbYes Then
            Do
    
                If a = vbYes Then
                    If r.Offset(1, 0).Value = "" Then
                        Set r = r.Offset(1, 0)
                    Else
                        Set r = r.End(xlDown).Offset(1, 0)
                    End If
                    Exit Do
                
                Else
                    MsgBox "msgbox yes / no - nigdy nie powinien sie pokazac ten msgbox"
                End If
                
                Set r = r.Offset(1, 0)
            Loop Until r = ""
        End If
    End With
End Function

Private Sub dodaj_pickup_gdy_zakladka_pn_jest_aktywna(ByRef r As Range)


    ' na dzien 201-10-09
    ' jest to tylko szablon pod przyszla implementacje

    

    With ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    
    

        ' current_form.ListBoxPNsFromThisDuns
        czy_moge_dodac_nowy_pickup = False
        ' miejsce puste ustawia sie na referencji r
        ' to jest logika, gdy arkusz master jest aktywny
        ' czy_moge_dodac_nowy_pickup = zweryfikuj_arkusz_pickups_i_ustaw_na_pierwsze_puste_miejsce(r)
            
        If czy_moge_dodac_nowy_pickup Then
    
    
            Set r = .Cells(1, WizardMain.O_INDX)
            Do
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
        End If
    End With
                
End Sub


Public Sub quick_layout_config()
    'MsgBox current_form.MultiPage.Pages.Item("PagePN").Name
    'MsgBox current_form.MultiPage.Pages.Item("PageDUNS").Name
    
    With current_form.MultiPage
    
        If .SelectedItem.Name = .Pages.item("PagePN").Name Then
            ' MsgBox .SelectedItem.Name
            ' wtedy form bedzie mniejszy
            
            current_form.Height = 200
            current_form.Width = 480
            current_form.BtnDodaj.Top = 138
            current_form.MultiPage.Height = 120
            
            
            ' dates to today
            If CStr(current_form.DTPickerPN_DeliveryDate.Value) = "1900-01-01" Then
                current_form.DTPickerPN_DeliveryDate.Value = Date
                current_form.DTPickerPN_PickUpDate.Value = Date
            End If
            
            
        ElseIf .SelectedItem.Name = .Pages.item("PageDUNS").Name Then
            ' MsgBox .SelectedItem.Name
            ' wtedy form bedzie wiekszy
            current_form.Height = 400
            current_form.Width = 480
            current_form.BtnDodaj.Top = 348
            current_form.MultiPage.Top = 12
            current_form.MultiPage.Height = 324
            
            
            ' dates to today
            If CStr(current_form.DTPickerDUNS_DeliveryDate.Value) = "1900-01-01" Then
                current_form.DTPickerDUNS_DeliveryDate.Value = Date
                current_form.DTPickerDUNS_PickUpDate.Value = Date
            End If
        End If
    End With
End Sub


Public Sub adjust_content_if_selection_changed()

    ' selection changed event od razu wszystkie reguly zostaly zdefiniowane, czy ma zostac uruchomiona dany sub
    ' thisworkbook is active
    ' master is active
    ' selection is one active cell
    
    
    
    If ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1 Then
    
    
        If ActiveSheet.Name = MASTER_SHEET_NAME Then
        
            If current_form.Name = "FormPickups" Then
        
                With current_form.MultiPage
                    If .SelectedItem.Name = .Pages.item("PagePN").Name Then
                    
                        inner_logic_for_pn_page_on_active_master_worksheet
                    ElseIf .SelectedItem.Name = .Pages.item("PageDUNS").Name Then
                    
                        inner_logic_for_duns_page_on_active_master_worksheet
                        
                    End If
                End With
            
            ElseIf current_form.Name = "FormEditPuses" Then
            
            
                ' 2015-10-05 draft implementacji
                ' ------------------------------------------------
                inner_logic_for_edit_puses_on_active_master_worksheet
                ' ------------------------------------------------
            End If
            
        ' to ma byc edycja!
        ElseIf ActiveSheet.Name = PICKUPS_SHEET_NAME Then
        
        
            ' tylko add - nazwa forma form pickups nie ma else
            ' narazie nie ma implementacji pod to
            If current_form.Name = "FormPickups" Then
            
                With current_form.MultiPage
                    If .SelectedItem.Name = .Pages.item("PagePN").Name Then
                        
                        inner_logic_for_pn_page_on_active_pickups_worksheet
                        
                    ElseIf .SelectedItem.Name = .Pages.item("PageDUNS").Name Then
                    
                        ' na dzien 20.ix.2015 no operation
                        ' =========================================
                    End If
                End With
            
            End If
            
        Else
            ' no operation
        End If
        
        
    End If
    

End Sub


Private Sub inner_logic_for_pn_page_on_active_master_worksheet()



    If current_form.ComboBoxSourceDUNS.Enabled = True Then

        ghost_activated = False
    
        current_form.BtnDodaj.Caption = "Dodaj"
        current_form.ComboBoxPN.Visible = True
        current_form.Label10.Visible = True
        
        current_form.ComboBoxSourceDUNS.Value = current_form.ComboBoxPN.Value
        ' MsgBox current_form.ComboBoxPN.Value
        If current_form.ComboBoxPN.Value = ALL_ORDERED_QTY Then
            column_source = MRD1_Ordered_QTY
        ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.MRD2_Ordered_QTY)) Then
            column_source = MRD2_Ordered_QTY
        ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.MRD1_Ordered_QTY)) Then
            column_source = MRD1_Ordered_QTY
        ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.Total_QTY)) Then
            column_source = Total_QTY
        ElseIf current_form.ComboBoxPN.Value = GHOST Then
            ghost_activated = True
        Else
            ' Exit Sub
            MsgBox "ten msgbox nigdy nie powinien sie pokazac adjust"
        End If
            
        
        ' adding only one part number
        ' =========================================
        
        ' 1. check if there's a place for this order
        ' (duplicates) and optionally notify user that he did some duplicates
        
        ' co by nie sciagac danych z labelek
        If ActiveCell.Row > 1 Then
        
            tmp_for_pn = Cells(ActiveCell.Row, Int(WizardMain.pn))
            tmp_for_duns = Cells(ActiveCell.Row, WizardMain.duns)
            tmp_for_fupcode = Cells(ActiveCell.Row, WizardMain.fup_code)
            
            tmp_for_pn = dodaj_zera_przed(CStr(tmp_for_pn), 8)
            tmp_for_duns = dodaj_zera_przed(CStr(tmp_for_duns), 9)
            
            current_form.TextBoxPN1.Value = tmp_for_pn & "," & tmp_for_duns & "," & tmp_for_fupcode
            
            ' nie zerujmy pozostawmy stare wartosci
            'current_form.TextBoxPusName1.Value = ""
            ' na podstawie checkboxa wybieramy wartosc
            If Not ghost_activated Then
                current_form.TextBoxQty1.Value = Cells(ActiveCell.Row, column_source)
            ElseIf ghost_activated Then
                current_form.TextBoxQty1.Value = 0
            Else
                MsgBox "to nie moze sie poakazac - pn text box qty1 na zero"
            End If
            
            ' juz wczesniej zostala opracowana logika
            ' current_form.DTPickerPN_DeliveryDate = d
        End If
        
        ' =========================================
    Else
        MsgBox "brak implementacji w przypadku synchro z jenny"
    End If
End Sub


Private Sub inner_logic_for_edit_puses_on_active_master_worksheet()

    ' aktywnym arkuszem jest MASTER
    ' plus do tego kliknelismy na edytowanie aktualnych PUSow
    
    ' najpierw sprawdzmy czy w ogole jest sens zaczynac:
    If ActiveCell.Row > 1 Then
        If Trim(ActiveCell.Parent.Cells(ActiveCell.Row, 1)) <> "" Then
            
            fill_edit_listboxes
            ' fill_labels
        End If
    End If
End Sub


Private Sub inner_logic_for_duns_page_on_active_master_worksheet()



    If current_form.ComboBoxSourceDUNS.Enabled = True Then
        
        normal_init_for_dodaj_pus
        fill_labels
    Else
        If Not jr Is Nothing Then jenny_init_for_dodaj_pus
    End If
End Sub

Private Sub jenny_init_for_dodaj_pus()
    If current_form.ComboBoxSourceDUNS.Enabled = False Then
        ' ================================================================================================
        ' ================================================================================================
        
        ' tutaj body bieremy dane z jr
        ' i wsadzamy w odpowiednie pole current forma
        ' ale oczywiscie w pierwszej kolejnosci sprawdz w arkuszu master,
        ' czy w ogole takie dane sa i czy mozna taki oto pus
        'Dim r As Range
        'Set r = Cells(1, 1)
        '
        'Set r = czy_mozna_dodac_tego_pusa_czy_jest_taki_pn()
        'r.Activate
            
        'If r.Row > 1 Then
        '
        '
        '    If sprawdz_czy_dane_w_wizardzie_sa_w_linii_z_danymi_w_jenny() Then
        
        
                tmp_for_duns = jr.item(2, 2)
                With current_form
                
                    .LabelDUNS.Caption = "DUNS: " & tmp_for_duns
                    .LabelSuppName.Caption = "SUPPLIER NAME: " & jr.item(1, 2)
                    .TextBoxMaskForPusNumber.Value = jr.item(7, 2)
                    .DTPickerDUNS_PickUpDate.Value = CDate(jr.item(12, 2))
                    .DTPickerDUNS_DeliveryDate.Value = CDate(jr.item(16, 2))
                    
                    
                    .ListBoxPNsFromThisDuns.Clear
                    .ListBoxSource.Clear
                    .ListBoxDone.Clear
                    .ListBoxCurrPusQty.Clear
                End With
                
                
                petla_wypelnienia_listboxow tmp_for_duns, E_JENNY_JENNY
        '    End If
        'End If
        
        jr.Parent.Activate
        ' jr.Activate
        
        ' ================================================================================================
    Else
        MsgBox "ten msgbox nie ma prawa sie pokazac - nieoczekiwane zlamanie logiki na subie jenny_init"
    End If
End Sub

Private Function sprawdz_czy_dane_w_wizardzie_sa_w_linii_z_danymi_w_jenny() As Boolean
    sprawdz_czy_dane_w_wizardzie_sa_w_linii_z_danymi_w_jenny = False
    
    If CStr(Cells(ActiveCell.Row, WizardMain.duns)) = CStr(jr.item(2, 2)) Then
        sprawdz_czy_dane_w_wizardzie_sa_w_linii_z_danymi_w_jenny = True
    End If
End Function

Private Function czy_mozna_dodac_tego_pusa_czy_jest_taki_pn() As Range
    
    Set czy_mozna_dodac_tego_pusa_czy_jest_taki_pn = Cells(1, 1)
    ' ta funkcja musi pracowac i tak na aktywnym arkuszu master wiec raczej bez obaw moge odnosic sie posrednio
    Dim r As Range
    Set r = Cells(2, 1)
    Do
        ' znajdz pierwszy lepszy pn
        ' znalazl sie
        If jr.item(2, 6) = r Then
            ' pn' ow nie bede sprawdzal
            Set czy_mozna_dodac_tego_pusa_czy_jest_taki_pn = r
            Exit Function
            ' rola tej logiki dalej bedzie lezala w pobieraniu danych tylko i wylacznie z wizarda
            ' informacjami z jenny tylko sie posilkujemy i nic wiecej - nic na sile
            ' jesli dane nie beda w lini najwyzej poleci jakis msgbox i nic wiecej
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function

Private Sub normal_init_for_dodaj_pus()

    ghost_activated = False

    current_form.ComboBoxPN.Value = current_form.ComboBoxSourceDUNS.Value
    
    If current_form.ComboBoxPN.Value = ALL_ORDERED_QTY Then
        column_source = MRD1_Ordered_QTY
    ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.MRD2_Ordered_QTY)) Then
        column_source = MRD2_Ordered_QTY
    ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.MRD1_Ordered_QTY)) Then
        column_source = MRD1_Ordered_QTY
    ElseIf current_form.ComboBoxPN.Value = Cells(1, Int(WizardMain.Total_QTY)) Then
        column_source = Total_QTY
    ElseIf current_form.ComboBoxPN.Value = GHOST Then
        ' no operation
        ghost_activated = True
    Else
        ' Exit Sub
        MsgBox "ten msgbox nigdy nie powinien sie pokazac adjust"
    End If
    
    ' adding many PN from one DUNS
    ' =========================================
    
    ' need to think how to do it
    ' but for sure there is a necessity to have notification about duplicates
    ' co by nie sciagac danych z labelek
    If ActiveCell.Row > 1 Then
    
        tmp_for_duns = Cells(ActiveCell.Row, WizardMain.duns)
        
        
        
        ' TextBoxMaskForPusNumber
        
        
        
        ' Me.ListBoxPNsFromThisDuns.Clear
        ' Me.ListBoxPNsFromThisDuns.AddItem
        With current_form
        
            .LabelDUNS.Caption = "DUNS: " & tmp_for_duns
            .LabelSuppName.Caption = "SUPPLIER NAME: " & Cells(ActiveCell.Row, WizardMain.Supplier_Name)
            .TextBoxMaskForPusNumber.Value = ""
            
            .ListBoxPNsFromThisDuns.Clear
            .ListBoxSource.Clear
            .ListBoxDone.Clear
            .ListBoxCurrPusQty.Clear
        End With
        
        petla_wypelnienia_listboxow tmp_for_duns, E_JENNY_NO_JENNY
        
        
        ' listboxy wypelnione
        ' zostalo teraz do uzupleniania tylko pus #
        stworz_nazwe_pusa
    End If
    
    ' =========================================
End Sub

Private Sub petla_wypelnienia_listboxow(tmp_for_duns, e As E_IS_JENNY)


    If e = E_JENNY_NO_JENNY Then

        Set r = Cells(2, WizardMain.pn)
        Do
        
            ' iterate list in master to check pn with same duns
            If ((Not r.EntireRow.Hidden) And (current_form.CheckBoxWorkOnlyOnVisibleRows.Value = True)) Or _
                (current_form.CheckBoxWorkOnlyOnVisibleRows.Value = False) Then
                
                    If CStr(tmp_for_duns) = CStr(r.Offset(0, WizardMain.duns - WizardMain.pn)) Then
                    
                    
                        If (current_form.CheckBoxOnlyFMAResp.Value = True) Then
                            If UCase(CStr(r.Offset(0, WizardMain.Responsibility - WizardMain.pn))) Like FMA_WITH_STARS Then
                            
                        
                        
                                spelnione_warunki_wypelnij_listboxy r, CStr(tmp_for_duns), ghost_activated
                            End If
                        ElseIf (current_form.CheckBoxOnlyFMAResp.Value = False) Then
                            
                            spelnione_warunki_wypelnij_listboxy r, CStr(tmp_for_duns), ghost_activated
                            
                        End If
                    End If
                
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
    ElseIf e = E_JENNY_JENNY Then
        wypelnij_listboxy_z_jenny tmp_for_duns
    Else
        MsgBox "critical - petla_wypelnienia_listboxow w PUS handler: jenny, no jenny"
    End If
End Sub

Private Sub wypelnij_listboxy_z_jenny(tmp_for_duns)


    Dim wr As Range
    Set wr = jr.item(2, 6)
    
    If Trim(wr) <> "" Then
        Do
        
            current_form.ListBoxPNsFromThisDuns.AddItem CStr(wr) & "," & CStr(tmp_for_duns) & _
                "," & CStr(jr.item(28, 2))
              
              
            liczba_orderu = -1
            liczba_juz_zrobionych_pusow = -1
            liczba_do_zrobienia = wr.Offset(0, 1)
           
            If True Then
                current_form.ListBoxSource.AddItem liczba_orderu
                current_form.ListBoxDone.AddItem liczba_juz_zrobionych_pusow
                current_form.ListBoxCurrPusQty.AddItem liczba_do_zrobienia
                
            Else
                MsgBox "nie powinien sie pokazac ten msgbox - ghost item place"
            End If
            
            Set wr = wr.Offset(1, 0)
        Loop Until Trim(wr) = ""
    
    End If
End Sub

Private Sub stworz_nazwe_pusa()
    ' zalozenie ze juz listboxy uzupelnione
    
    With current_form
        ' .TextBoxMaskForPusNumber.Value = ""
        'For x = 0 To ListBoxPNsFromThisDuns.ListCount - 1
        'Next x
        
        If .ListBoxPNsFromThisDuns.ListCount > 0 Then
            arr = Split(.ListBoxPNsFromThisDuns.List(0), COMMA)
            deck = arr(UBound(arr))
            
            .TextBoxMaskForPusNumber.Value = "FMA" & CStr(deck) & ktory_kolejny_numer_dla_pusa(CStr("FMA" & CStr(deck)), CLng(1))
            
        End If
    End With
End Sub

Private Function ktory_kolejny_numer_dla_pusa(ptrn As String, i As Long) As Long
    ktory_kolejny_numer_dla_pusa = i
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME).Cells(2, WizardMain.O_PUS_Number)
    Do
        If r = ptrn & CStr(ktory_kolejny_numer_dla_pusa) Then
            ktory_kolejny_numer_dla_pusa = ktory_kolejny_numer_dla_pusa(ptrn, i + 1)
            Exit Do
        End If
        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
    Loop Until r.Row = WizardMain.CAPACITY_ARKUSZA
End Function


Private Sub inner_logic_for_pn_page_on_active_pickups_worksheet()
    ' adding only one part number
    ' =========================================
    
    ' 1. check if there's a place for this order
    ' (duplicates) and optionally notify user that he did some duplicates
    
    ghost_activated = False

    current_form.BtnDodaj.Caption = "Edytuj"
    current_form.ComboBoxPN.Visible = False
    current_form.Label10.Visible = False
    
    ' ta logika wlasciwie jest taka sama jak przy pracy gdy arkusz master jest aktywny z ta mala roznica ze tutaj dane
    ' edytujemy a tam za kazdym razem dodajemy
    
    ' dodatkowa blokada przed wejsciem na labelki
    If ActiveCell.Row > 1 Then
    
        tmp_for_pn = Cells(ActiveCell.Row, Int(WizardMain.O_PN))
        tmp_for_duns = Cells(ActiveCell.Row, WizardMain.O_DUNS)
        tmp_for_fupcode = Cells(ActiveCell.Row, WizardMain.O_FUP_code)
        
        tmp_for_pn = dodaj_zera_przed(CStr(tmp_for_pn), 8)
        tmp_for_duns = dodaj_zera_przed(CStr(tmp_for_duns), 9)
        
        
        ' current_form.BtnDodaj.Caption = "Edytuj"
        
        current_form.TextBoxPN1.Value = tmp_for_pn & "," & tmp_for_duns & "," & tmp_for_fupcode
        
        
        If Not ghost_activated Then
            current_form.TextBoxQty1.Value = Cells(ActiveCell.Row, WizardMain.O_Pick_up_Qty)
        ElseIf ghost_activated Then
            current_form.TextBoxQty1.Value = 0
        Else
            MsgBox "to nie moze sie poakazac - pn text box qty1 na zero"
        End If
        
        
        
        current_form.DTPickerPN_PickUpDate = Cells(ActiveCell.Row, WizardMain.O_Pick_up_date)
        current_form.DTPickerPN_DeliveryDate = Cells(ActiveCell.Row, WizardMain.O_Delivery_Date)
        current_form.TextBoxPusName1 = Cells(ActiveCell.Row, WizardMain.O_PUS_Number)
    End If
    ' =========================================
End Sub



Private Sub spelnione_warunki_wypelnij_listboxy(ByRef r As Range, tmp_for_duns As String, Optional ghost_item As Boolean)

    current_form.ListBoxPNsFromThisDuns.AddItem CStr(r) & "," & CStr(tmp_for_duns) & _
        "," & CStr(r.Offset(0, WizardMain.fup_code - WizardMain.pn))
      
      
    liczba_orderu = 0
    liczba_juz_zrobionych_pusow = 0
    liczba_do_zrobienia = 0
    

    If current_form.ComboBoxPN.Value = ALL_ORDERED_QTY Then
        If CStr(r.Offset(0, column_source - WizardMain.pn)) <> "" Then
            liczba_orderu = r.Offset(0, WizardMain.MRD2_Ordered_QTY - WizardMain.pn) + r.Offset(0, WizardMain.MRD1_Ordered_QTY - WizardMain.pn)
        End If
    Else
        
    
        If CStr(r.Offset(0, column_source - WizardMain.pn)) <> "" Then
            liczba_orderu = r.Offset(0, column_source - WizardMain.pn)
        End If
    End If
        
        
    If r.Offset(0, WizardMain.Total_PUS_QTY - WizardMain.pn) <> "" Then
        liczba_juz_zrobionych_pusow = CLng(r.Offset(0, WizardMain.Total_PUS_QTY - WizardMain.pn))
    End If
    
    liczba_do_zrobienia = liczba_orderu - liczba_juz_zrobionych_pusow
    
    
    
    
    ' If Not (liczba_do_zrobienia < 0) Then
       
        If Not ghost_item Then
            current_form.ListBoxSource.AddItem CStr(liczba_orderu)
            current_form.ListBoxDone.AddItem CStr(liczba_juz_zrobionych_pusow)
            current_form.ListBoxCurrPusQty.AddItem CStr(liczba_do_zrobienia)
        ElseIf ghost_item Then
            current_form.ListBoxSource.AddItem 0
            current_form.ListBoxDone.AddItem 0
            current_form.ListBoxCurrPusQty.AddItem 0
            ' ListBoxDone
            ' ListBoxCurrPusQty
        Else
            MsgBox "nie powinien sie pokazac ten msgbox - ghost item place"
        End If
    'End If
End Sub








Private Sub Class_Initialize()
    co_poszlo_zle = ""
    Set jr = Nothing
End Sub


Public Function dodaj_zera_przed(ByRef jaki_text As String, ile As Integer)

    ile_zer_dodajemy_przed = ile - Len(jaki_text)
    zeros = ""
    If ile_zer_dodajemy_przed > 0 Then
        For x = 1 To ile_zer_dodajemy_przed
            zeros = zeros & "0"
        Next x
    End If
    dodaj_zera_przed = zeros & jaki_text
    

End Function
