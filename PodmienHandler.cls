VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PodmienHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa nie pracuje jeszcze z subem
' work_around_custom_copy_worksheet
' nie wiem czemu ale wygodniej mi jest gdy ona jest niezalezna

Public out_msh As Worksheet
Public ost_r As Range
Public this_msh As Worksheet
Public to_cp As Range

Public source_workbook As Workbook
Public this_workbook As Workbook

Private dsh As Worksheet
Private psh As Worksheet

Private cs As Worksheet

Private sdsh As Worksheet
Private spsh As Worksheet

Public Sub just_clear()
    Me.clear_master_sheet
    Me.clear_pickups_sheet
    Me.clear_details_sheet
    Me.clear_comment_source
    
    usun_customowe_sheety
    
    MsgBox "gotowe!"
End Sub

Private Sub usun_customowe_sheety()

    Dim sh As Worksheet
    
    For Each sh In Me.this_workbook.Sheets
    
        If sh.Name <> WizardMain.PICKUPS_SHEET_NAME Then
            If sh.Name <> WizardMain.MASTER_SHEET_NAME Then
                If sh.Name <> WizardMain.CONFIG_SHEET_NAME Then
                    If sh.Name <> WizardMain.REGISTER_SHEET_NAME Then
                        If sh.Name <> WizardMain.DETAILS_SHEET_NAME Then
                            If sh.Name <> WizardMain.CUSTOM_COPY_SHEET_NAME Then
                                If sh.Name <> WizardMain.COMMENT_SOURCE_SHEET_NAME Then
                                    If sh.Name <> WizardMain.DELIVERY_CONFIRMATION_SPECIAL_SHEET_NAME Then
                                        If sh.Name <> WizardMain.CACHE Then
                                            ' body to delete
                                        
                                            answer = MsgBox("Czy chcesz usunac arkusz: " & sh.Name & "?", vbYesNo, "Usuniecie customowego arkusza")
                                        
                                            If answer = vbYes Then
                                                sh.Delete
                                                usun_customowe_sheety
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next sh
End Sub


Public Sub copy_from_raw_td()
    
    If sprawdz_czy_przypadkiem_dany_td_nie_ma_juz_std_wizard() Then
        MsgBox "plik ktory wybrales okazal sie zgodny ze std wizarda, dalej lecimy z logika zgodna dla wizarda!"
        go_on_with_coping_data_in_master_sheet_in_std_only
    Else
    
        ' nie ma std wizardowego wiec opor
        go_on_with_coping_data_in_master_sheet_in_non_std_only
        ' MsgBox "plik ktory wybrales nie jest w standardzie wizarda"
    End If
End Sub

Private Function sprawdz_czy_przypadkiem_dany_td_nie_ma_juz_std_wizard() As Boolean
    sprawdz_czy_przypadkiem_dany_td_nie_ma_juz_std_wizard = False
    
    Dim sh As Worksheet
    For Each sh In source_workbook.Sheets
        
        If sh.Name = WizardMain.MASTER_SHEET_NAME Then
            Set out_msh = sh
            sprawdz_czy_przypadkiem_dany_td_nie_ma_juz_std_wizard = sprawdz_czy_ten_arkusz_master_jest_wg_std(out_msh)
            Exit For
        End If
    Next sh
    
    
    
End Function

Private Function sprawdz_czy_ten_arkusz_master_jest_wg_std(master_sh As Worksheet)
    sprawdz_czy_ten_arkusz_master_jest_wg_std = False
    
    sprawdz_czy_ten_arkusz_master_jest_wg_std = sprawdz_czy_kolejnosc_std_kolumn_jest_wlasciwa(master_sh, this_msh)
    
    
End Function

Public Sub go_on_with_coping_data()

    
    Me.clear_master_sheet
    Me.clear_pickups_sheet
    Me.clear_details_sheet
    Me.clear_comment_source
    
    usun_customowe_sheety
    
    
    ' klejmy teraz dopiero
    master_worksheet_cp
    
    ' tutaj sprawdzimy czy po prostu jest taki arkusz w ogole
    details_worksheet_cp
    
    ' pus sheet sprawdz nazwy i kolejnosc kolumn
    pikcups_worksheet_cp
    
    ' comment source
    comment_source_worksheet_cp
    
    
    ' kopiowanie calych arkuszy
    ' --------------------------------------------
    
    copy_custom_worksheets
    ' --------------------------------------------
        
    MsgBox "gotowe!"
    

End Sub

Public Sub go_on_with_coping_data_in_master_sheet_in_non_std_only()



    Me.clear_master_sheet
    Me.clear_pickups_sheet
    Me.clear_details_sheet
    
    usun_customowe_sheety
    
    
    ' klejmy teraz dopiero
    custom_master_worksheet_cp
    ' tutaj sprawdzimy czy po prostu jest taki arkusz w ogole
    'details_worksheet_cp
    
    ' pus sheet sprawdz nazwy i kolejnosc kolumn
    'pikcups_worksheet_cp
    
    
    ' kopiowanie calych arkuszy
    ' --------------------------------------------
    
    'copy_custom_worksheets
    ' --------------------------------------------
        
    ' MsgBox "gotowe!"

End Sub

Public Sub go_on_with_coping_data_in_master_sheet_in_std_only()

    
    Me.clear_master_sheet
    Me.clear_pickups_sheet
    Me.clear_details_sheet
    
    usun_customowe_sheety
    
    
    ' klejmy teraz dopiero
    master_worksheet_cp
    ' tutaj sprawdzimy czy po prostu jest taki arkusz w ogole
    'details_worksheet_cp
    
    ' pus sheet sprawdz nazwy i kolejnosc kolumn
    'pikcups_worksheet_cp
    
    
    ' kopiowanie calych arkuszy
    ' --------------------------------------------
    
    'copy_custom_worksheets
    ' --------------------------------------------
        
    MsgBox "gotowe!"
    

End Sub

Private Sub copy_custom_worksheets()

    'Set Me.source_workbook = Workbooks(CStr(nap))
    'Set Me.this_workbook = ThisWorkbook
    '
    'Set out_msh = Workbooks(CStr(nap)).Sheets(MASTER_SHEET_NAME)
    'Set this_msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    
    Dim wrsh As Worksheet
    For Each wrsh In Me.source_workbook.Sheets
    
        ' sprawdz czy arkusz jest customowy
        If is_wrsh_is_custom(wrsh) Then
            answer = MsgBox("chcesz skopiowac arkusz " & wrsh.Name & "?", vbYesNo, "Copy")
            
            If answer = vbYes Then
                wrsh.Copy Me.this_workbook.Sheets(1)
            Else
                MsgBox "nie skopiowano arkusza: " & wrsh.Name
                
            End If
        End If
    Next wrsh
    
End Sub

Private Function is_wrsh_is_custom(sh As Worksheet) As Boolean
    is_wrsh_is_custom = False
    
    If sh.Name <> WizardMain.PICKUPS_SHEET_NAME Then
        If sh.Name <> WizardMain.MASTER_SHEET_NAME Then
            If sh.Name <> WizardMain.CONFIG_SHEET_NAME Then
                If sh.Name <> WizardMain.REGISTER_SHEET_NAME Then
                    If sh.Name <> WizardMain.DETAILS_SHEET_NAME Then
                        If sh.Name <> WizardMain.CUSTOM_COPY_SHEET_NAME Then
                            If sh.Name <> WizardMain.COMMENT_SOURCE_SHEET_NAME Then
                                If sh.Name <> WizardMain.DELIVERY_CONFIRMATION_SPECIAL_SHEET_NAME Then
                                    If sh.Name <> WizardMain.CACHE Then
                                        is_wrsh_is_custom = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        is_wrsh_is_custom = False
    End If
End Function

Private Sub pikcups_worksheet_cp()

    ' wczesniej jest juz uruchomione to przypisanie
    'Set psh = Nothing
    'On Error Resume Next
    'Set psh = Me.source_workbook.Sheets(PICKUPS_SHEET_NAME)
    
    clear_pickups_sheet
    
    If Not spsh Is Nothing Then
        
        te_same_arkusze_pickups = True
        
        For x = 1 To WizardMain.O_PUS_Number
            
            If spsh.Cells(1, x) = Me.this_workbook.Sheets(PICKUPS_SHEET_NAME).Cells(1, x) Then
            Else
                te_same_arkusze_pickups = False
                MsgBox "nie ma std na arkuszu pickups"
                Exit For
            End If
        Next x
        
        ' wszystkie nazwy te same mozna kopiowac!
        If te_same_arkusze_pickups Then
        
            Dim r As Range, sr As Range
            With Me.this_workbook.Sheets(PICKUPS_SHEET_NAME)
                
                Set r = .Range(.Cells(2, WizardMain.O_INDX), _
                    .Cells(WizardMain.CAPACITY_ARKUSZA, WizardMain.O_PUS_Number))
                r.Clear
                r.Value = ""
                
                
            End With
            
            With spsh
                Set sr = .Range(.Cells(2, WizardMain.O_INDX), _
                    .Cells(WizardMain.CAPACITY_ARKUSZA, WizardMain.O_PUS_Number))
                sr.Copy r
            End With
        End If
    Else
        MsgBox "arkusz: " & PICKUPS_SHEET_NAME & " nie zostal odnaleziony"
    End If

End Sub

Private Sub details_worksheet_cp()
    
    ' wczesniej juz zostalo wykonaneprzypisanie na sdsh
    
    If Not sdsh Is Nothing Then
    
        Dim r As Range
        Set r = Me.this_workbook.Sheets(DETAILS_SHEET_NAME).Range("b1")
        For Each ir In sdsh.Range("b1:b100")
            r = ir
            Set r = r.Offset(1, 0)
        Next ir
    Else
        MsgBox "arkusz: " & DETAILS_SHEET_NAME & " nie zostal odnaleziony"
    End If
End Sub


Public Sub clear_master_sheet()
    ' teraz lecimy juz z logika, zatem moge usunac stare dane
    this_msh.Range(this_msh.Cells(2, 1), this_msh.Cells(CLng(WizardMain.POLOWA_CAPACITY_ARKUSZA), WizardMain.BLANK4)).Clear
    ' this_msh.Range(this_msh.Cells(2, 1), this_msh.Cells(CLng(WizardMain.POLOWA_CAPACITY_ARKUSZA), WizardMain.BLANK4)).EntireRow.Delete
    
    
    ' trzeba jeszcze clear zrobic na nie std kolumnach
    Dim r As Range
    x = WizardMain.BLANK4 + 1
    Set r = this_msh.Range(this_msh.Cells(1, x), this_msh.Cells(CLng(WizardMain.POLOWA_CAPACITY_ARKUSZA), x))
    Do
        r.Clear
        Set r = r.Offset(0, 1)
    Loop Until Trim(r.item(1, 1)) = ""
End Sub

Public Sub clear_pickups_sheet()
    psh.Range(psh.Cells(2, 1), psh.Cells(CLng(WizardMain.POLOWA_CAPACITY_ARKUSZA), WizardMain.O_PUS_Number)).Clear
    ' psh.Range(psh.Cells(2, 1), psh.Cells(CLng(WizardMain.POLOWA_CAPACITY_ARKUSZA), WizardMain.O_PUS_Number)).EntireRow.Delete
End Sub

Public Sub clear_details_sheet()

    ' juz wczesniej
    'Set dsh = Nothing
    'On Error Resume Next
    'Set dsh = Me.source_workbook.Sheets(DETAILS_SHEET_NAME)
    
    dsh.Range("b1:c100").Value = ""
End Sub


Public Sub clear_comment_source()
        
    Set cs = ThisWorkbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME)
    
    ' tak losowo te BB :P
    cs.Range("A1:BB" & CStr(WizardMain.POLOWA_CAPACITY_ARKUSZA)).Clear
    
End Sub
        
Private Sub custom_master_worksheet_cp()


    If True Then
    
        ' to jest OK
        this_workbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2") = 1
        CustomMasterCpForm.BtnPrev.Enabled = False
        CustomMasterCpForm.Label1.Caption = _
            this_workbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Cells(this_workbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME).Range("d2"), 2)
        CustomMasterCpForm.Label2.Caption = WizardMain.NIC_NIE_WYBRANO_TXT
        Set CustomMasterCpForm.podmien_handler = Me
        
        ' LISTBOX_CUSTOM_COLUMN_NAMES_LIMIT
        With CustomMasterCpForm
            .ListBoxRawData.Clear
            
            Dim iter_sh As Worksheet
            i = 1
            For Each iter_sh In source_workbook.Sheets
            
                .ListBoxRawData.AddItem iter_sh.Name
                
                i = i + 1
            Next iter_sh
        End With
        
        CustomMasterCpForm.ListBoxRawData.MultiSelect = fmMultiSelectSingle
        CustomMasterCpForm.show vbModeless
        
    Else
        ' tu cos poszlo nie tak
        MsgBox "out_msh teraz nie moze nie byc pusty, cos nie tak z logika sub custom master worksheet cp!"
        
        
    End If
End Sub

Private Sub master_worksheet_cp()
    If sprawdz_czy_kolejnosc_std_kolumn_jest_wlasciwa(out_msh, this_msh) Then
        
        ' wczesniej juz bylo
        ' clear_master_sheet
        
        
        out_msh.Parent.Activate
        out_msh.Activate
        
        ' yellow columns
        ' ===========================================================================
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.PN
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Alternative_PN
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.PN_Name
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.GPDS_PN_Name
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.DUNS
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Supplier_Name
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Country_code
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MGO_code
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Responsibility
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.fup_code
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.SQ
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.PPAP_Status
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.SQ_Comments
        'kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.ADD_to_T_slash_D
        
        ' teortycznie caly ten kod mozna byc zamknac w jeden petli
        ' ----------------------------------------------------------------------
        For x = 1 To WizardMain.ADD_to_T_slash_D
            kopiowanie_ out_msh, this_msh, ost_r.Row, Int(x)
        Next x
        ' ===========================================================================
        
        'blue columns
        ' ===========================================================================
        ' mrd 1
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD1_Ordered_date
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD1_Ordered_QTY
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD1_confirmed_qty
        
        ' mrd 2
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD2_Ordered_date
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD2_Ordered_QTY
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.MRD2_confirmed_qty
        
        
        ' confirmations
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Delivery_confirmation
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.First_Confirmed_PUS_Date
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Delivery_reconfirmation
        
        'rest
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Comments
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Bottleneck
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.Future_Osea
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.DRE
        kopiowanie_ out_msh, this_msh, ost_r.Row, WizardMain.EDI_Received
        
        ' blanki (gdzie blank1 stal sie capacity
        For x = WizardMain.Capacity To WizardMain.BLANK4
            kopiowanie_ out_msh, this_msh, CLng(ost_r.Row), Int(x)
        Next x
        
        
        ' kolumny spoza standardu
        
        kopiowanie_customowych_kolumn_spoza_standardu out_msh, this_msh, CLng(ost_r.Row)
        
        ' ===========================================================================
        ' ----------------------------------------------------------------------
        
        ThisWorkbook.Activate
        ' MsgBox "arkusz mastergotowe"
    Else
        MsgBox "uklad kolumn nie jest w standardzie Wizard - kopiowanie danych nie odbedzie sie"
    End If

End Sub


Private Sub comment_source_worksheet_cp()

    ' nie bede nic sprawdzal - po prostu przekopiuje ten arkusz
    ' ==================================================
    'Public source_workbook As Workbook
    'Public this_workbook As Workbook
    
    
    Dim podwiersz As Range
    Dim podkolumne As Range
    
    
    'out_msh.Parent.Activate
    'out_msh.Activate
    source_workbook.Activate
    source_workbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME).Activate
    
    With source_workbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME)
    
        Set podwiersz = .Range("A1")
        Set podkolumne = .Range("A1")
        
        Do
            Set podwiersz = podwiersz.Offset(1, 0)
        Loop Until podwiersz = ""
        
        Do
            Set podkolumne = podkolumne.Offset(0, 1)
        Loop Until podkolumne = ""
    End With
    
    
    
    kopiowanie_ source_workbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME), this_workbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME), _
        podwiersz.Row, _
        podkolumne.Column, 1
    
    ' ==================================================

End Sub


Public Sub ref_first(nap As String, Optional lecimy_poza_std As Boolean)



    If lecimy_poza_std = True Then
        If nap <> "" Then
            ' txt nie jest pusty zatem zostalo cos wybrane do przekopiowania
            Set Me.source_workbook = Workbooks(CStr(nap))
            
            Set out_msh = Nothing
            
            Set Me.this_workbook = ThisWorkbook
            Set this_msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
            Set psh = Nothing
            On Error Resume Next
            Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
            Set dsh = Nothing
            On Error Resume Next
            Set dsh = ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
            
            
        End If
    Else
    
        ' quick hack
        If nap <> "" Then
            Set Me.source_workbook = Workbooks(CStr(nap))
            Set out_msh = Workbooks(CStr(nap)).Sheets(MASTER_SHEET_NAME)
            Set ost_r = out_msh.Range("a2")
            If ost_r.Offset(1, 0) <> "" Then
                Set ost_r = ost_r.End(xlDown)
            End If
            
            Set Me.source_workbook = Workbooks(CStr(nap))
            Set spsh = Nothing
            On Error Resume Next
            Set spsh = Me.source_workbook.Sheets(PICKUPS_SHEET_NAME)
            
            
            Set sdsh = Nothing
            On Error Resume Next
            Set sdsh = Me.source_workbook.Sheets(DETAILS_SHEET_NAME)
        End If
        
        
        Set Me.this_workbook = ThisWorkbook
        Set this_msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
        Set psh = Nothing
        On Error Resume Next
        Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
        Set dsh = Nothing
        On Error Resume Next
        Set dsh = ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
    End If
    
    
End Sub




Private Function sprawdz_czy_kolejnosc_std_kolumn_jest_wlasciwa(ByRef out_msh As Worksheet, ByRef this_msh As Worksheet) As Boolean

    sprawdz_czy_kolejnosc_std_kolumn_jest_wlasciwa = True
    For x = 1 To WizardMain.EDI_Received
    
        If out_msh.Cells(1, x).Value <> this_msh.Cells(1, x).Value Then
            sprawdz_czy_kolejnosc_std_kolumn_jest_wlasciwa = False
            MsgBox "kolumny nie sa w std!"
            Exit Function
        End If
        
    Next x
    
End Function

Private Sub kopiowanie_customowych_kolumn_spoza_standardu(ByRef out_msh As Worksheet, ByRef this_msh As Worksheet, w As Long)

    x = WizardMain.BLANK4 + 1
    Do
    
        ' kopiowanie labelki jest jednoczesnie warunkiem sprawdzajacym, czy jest co kopiowac w kolejnej formule
        If kopiowanie_labelki(out_msh, this_msh, CLng(ost_r.Row), Int(x)) Then
            kopiowanie_ out_msh, this_msh, CLng(ost_r.Row), Int(x)
        Else
            Exit Sub
        End If
        x = x + 1
    Loop

End Sub

Public Sub kopiowanie_(ByRef out_msh As Worksheet, ByRef this_msh As Worksheet, w As Long, x As Integer, Optional od_ktorego_wiersza As Integer)


    ' tutaj tylko jedna konkretna kolumna
    If od_ktorego_wiersza = 0 Then
    
        out_msh.Range(out_msh.Cells(2, x), out_msh.Cells(w, x)).Copy this_msh.Range(this_msh.Cells(2, x), this_msh.Cells(w, x))
    ElseIf od_ktorego_wiersza > 0 Then
    
    
        ' ten else ma w ogole inna zasade dzialania
        ' ten kod moze w przyszlosci przysporzyc mase problemow z powodu rozchwiania logiki dzialania znacznej
        out_msh.Range(out_msh.Cells(od_ktorego_wiersza, 1), out_msh.Cells(w, x)).Copy _
            this_msh.Range(this_msh.Cells(od_ktorego_wiersza, 1), this_msh.Cells(w, x))
    End If

End Sub

Private Function kopiowanie_labelki(ByRef out_msh As Worksheet, ByRef this_msh As Worksheet, w As Long, x As Integer) As Boolean
    If Trim(out_msh.Cells(1, x).Value) <> "" Then
        out_msh.Cells(1, x).Copy this_msh.Cells(1, x)
        kopiowanie_labelki = True
    Else
        kopiowanie_labelki = False
    End If
End Function

