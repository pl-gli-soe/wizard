VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private czy_walidacja_przebiegla_pomyslnie As Boolean

Private master_worksheet As Worksheet
Private mw As Worksheet

Private reg As Worksheet

Private error_wk As Workbook
Private error_sh As Worksheet

' ref
Private r As Range, ir As Range
Private last_row_from_pn_column As Long
Private lr As Long


Private co_poszlo_zle As String



Public Sub waliduj_arkusz_master()

    ' na poczatek sprawdzenie czy jest w ogole co walidowac
    
    Set r = mw.Range("a2")
    If Trim(r) <> "" Then
    
        Dim tmp_lng As E_MASTER_MANDATORY_COLUMNS
        tmp_lng = 1
        
        
        Do
            ' blad zanczyna od pierwsszego wierssza zamiast od drugiego
            Set r = mw.Range(mw.Cells(2, tmp_lng), mw.Cells(CLng(lr), tmp_lng))
            czy_walidacja_przebiegla_pomyslnie = czy_walidacja_przebiegla_pomyslnie And waliduj_inner(tmp_lng, r)
            tmp_lng = tmp_lng + 1
        Loop Until tmp_lng = BLANK4 + 1
        
        
        
    Else
        MsgBox "ten master jest pusty", vbInformation
    End If

    

End Sub

Private Function waliduj_inner(emmc As E_MASTER_MANDATORY_COLUMNS, ByRef r As Range) As Boolean


    ' narazie niech wszystko przechodzi dopoki nie mam gotowej logiki
    waliduj_inner = True
    
    
    
    If emmc = pn Then
        petla_po_pn r, waliduj_inner
    ElseIf emmc = duns Then
        petla_po_duns r, waliduj_inner
    ElseIf emmc = MGO_code Then
        petla_po_mgo_code r, waliduj_inner
    ElseIf emmc = fup_code Then
        petla_po_fup_code r, waliduj_inner
    ElseIf emmc = Total_QTY Then
        petla_po_total_qty r, waliduj_inner
    End If
        
    
    
    
End Function


Private Sub petla_po_fup_code(ByRef r As Range, ByRef wi As Boolean)
    ' validate Part #
    Dim wi_tmp As Boolean
    For Each ir In r
        wi_tmp = (Len(CStr(ir)) = 2) Or (CStr(ir) = "")
        
        If wi_tmp = False Then
            co_poszlo_zle = co_poszlo_zle & "fup code, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            wi = False
            'Exit Sub
        End If
    Next ir
End Sub


Private Sub petla_po_pn(ByRef r As Range, ByRef wi As Boolean)
    ' validate Part #
    
    Dim wi_tmp As Boolean
    For Each ir In r
        wi_tmp = IsNumeric(ir) Or (Len(CStr(ir)) > reg.Cells(WizardMain.pn + 1, 4))
        
        If wi_tmp = False Then
            co_poszlo_zle = co_poszlo_zle & "pn, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            wi = False
            'Exit Sub
        End If
        
        wi_tmp = (CBool(Left(ir, 1) <> "0") Or CBool(CStr(ir) = "0"))
        
        If wi_tmp = False Then
            wi = False
            co_poszlo_zle = co_poszlo_zle & "konwertuj pn na liczbe, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            'Exit Sub
        End If
    Next ir
End Sub

Private Sub petla_po_duns(ByRef r As Range, ByRef wi As Boolean)
    ' validate Part #
    Dim wi_tmp As Boolean
    For Each ir In r
        wi_tmp = IsNumeric(ir) Or (Len(CStr(ir)) > reg.Cells(WizardMain.duns + 1, 4))
        
        If wi_tmp = False Then
            co_poszlo_zle = co_poszlo_zle & "duns, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            wi = False
            'Exit Sub
        End If
        
        ' Debug.Assert CStr(ir) <> "000242693"
        
        wi_tmp = (CBool(Left(ir, 1) <> "0") Or CBool(CStr(ir) = "0"))
        
        If wi_tmp = False Then
            wi = False
            co_poszlo_zle = co_poszlo_zle & "konwertuj duns na liczbe, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            'Exit Sub
        End If
    Next ir
End Sub

Private Sub petla_po_mgo_code(ByRef r As Range, ByRef wi As Boolean)
    ' validate Part #
    Dim wi_tmp As Boolean
    For Each ir In r
        ' Debug.Print reg.Cells(WizardMain.MGO_code + 1, 4)
        wi_tmp = (Len(CStr(ir)) = reg.Cells(WizardMain.MGO_code + 1, 4)) Or (CStr(ir) = "")
        
        If wi_tmp = False Then
            co_poszlo_zle = co_poszlo_zle & "mgo code, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz, " & ir.Row & Chr(10)
            wi = False
            'Exit Sub
        End If
    Next ir
End Sub

Private Sub petla_po_total_qty(ByRef r As Range, ByRef wi As Boolean)
    ' validate Part #
    Dim wi_tmp As Boolean
    For Each ir In r
        wi_tmp = (Trim(ir) = "") Or (IsNumeric(ir))
        
        If wi_tmp = False Then
        
            co_poszlo_zle = co_poszlo_zle & "total qty, PN: " & ir.Parent.Cells(ir.Row, WizardMain.pn) & ", wiersz: " & ir.Row & Chr(10)
            wi = False
            'Exit Sub
        End If
    Next ir
End Sub


Public Sub pokaz_wynik_walidacji()
    
    If czy_walidacja_przebiegla_pomyslnie Then
        MsgBox "walidacja przebiegla pomyslnie", vbInformation
    Else
        MsgBox "walidacja nie przebiegla pomyslnie", vbCritical
        Set error_wk = Workbooks.Add
        Set error_sh = error_wk.Sheets.Add
        On Error Resume Next
        error_sh.Name = error_sh.Name & "_LOG"
        arr = Split(co_poszlo_zle, Chr(10))
        
        For x = LBound(arr) To UBound(arr)
            error_sh.Cells(x + 1, 1) = arr(x)
        Next x
    End If
End Sub

Private Sub Class_Initialize()
    czy_walidacja_przebiegla_pomyslnie = True
    
    Set master_worksheet = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    Set mw = master_worksheet
    
    Set reg = ThisWorkbook.Sheets(REGISTER_SHEET_NAME)
    
    ' pierwszy potencjalny PN
    Set r = mw.Cells(1, WizardMain.pn).Offset(1, 0)
    If Trim(r.Offset(1, 0)) = "" Then
        last_row_from_pn_column = mw.Cells(1, WizardMain.pn).Row
    Else
        last_row_from_pn_column = mw.Cells(1, WizardMain.pn).End(xlDown).Row
    End If
    
    co_poszlo_zle = ""
    
    lr = last_row_from_pn_column
    
End Sub

Public Function make_validation_on_all_pickups_in_system(ByRef sh As Worksheet) As Boolean
    make_validation_on_all_pickups_in_system = False
    
    
    ' if everything went ok then
    ' make_validation_on_all_pickups_in_system = true
    
    ' sprawdz czy to ten arkusz
    If sh.Name = PICKUPS_SHEET_NAME Then
    
        If sprawdz_czy_arkusz_nie_jest_przypadkiem_pusty(sh) Then

            ' jesli arkusz jest pusty validacje pomiajmy i dajemy true
            make_validation_on_all_pickups_in_system = True
            
        ElseIf sprawdz_prostokatnsc_danych_w_tym_arkuszu(sh) Then
            If sprawdz_uklad_kolumny_indx(sh) Then
                make_validation_on_all_pickups_in_system = True
            End If
        End If

    End If
    
End Function

Private Function sprawdz_uklad_kolumny_indx(sh As Worksheet) As Boolean

    ' tutaj troche inaczej poniewaz
    ' w petli wystarczy ze raz sie cos sypnie
    ' i operator and to wylapie
    sprawdz_uklad_kolumny_indx = True
    
    Dim r As Range
    Set r = sh.Cells(2, WizardMain.O_INDX)
    Do
        sprawdz_uklad_kolumny_indx = sprawdz_uklad_kolumny_indx And sprawdz_ten_cell_indx(r)
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    ' jesli do tej pory uklad jest na true - to znaczy ze wszystko ladnie przeszlo przez walidacje
    ' jesli nie - no to klops - wyrzyc dodatkowy text do msgboxa
    If sprawdz_uklad_kolumny_indx = False Then
        co_poszlo_zle = co_poszlo_zle & "kolumna pierwsza nie zostala poprawnie zwalidowana - nie zostal zachowany pattern zapisu" & Chr(10)
    End If
    
    
    
End Function

Private Function sprawdz_ten_cell_indx(r As Range) As Boolean
    sprawdz_ten_cell_indx = False
    
    arr = Split(r.Value, ",")
    
    If UBound(arr) = 2 Then
    
        ' pierwszy element PN - to jest liczba
        If IsNumeric(arr(LBound(arr))) Then
        
            ' duns to liczba
            If IsNumeric(arr(LBound(arr) + 1)) Then
            
                ' deck to txt z dwoch liter
                If Len(arr(UBound(arr))) = 2 Then
                    sprawdz_ten_cell_indx = True
                End If
            End If
        End If
    End If
    
    
End Function

Private Function sprawdz_czy_arkusz_nie_jest_przypadkiem_pusty(sh As Worksheet) As Boolean
    sprawdz_czy_arkusz_nie_jest_przypadkiem_pusty = False
    
    If sh.Range("A2").Value = "" Then
        sprawdz_czy_arkusz_nie_jest_przypadkiem_pusty = True
    End If
End Function

Private Function sprawdz_prostokatnsc_danych_w_tym_arkuszu(ByRef sh As Worksheet) As Boolean
    sprawdz_prostokatnsc_danych_w_tym_arkuszu = False
    
    
    ile_wierszy = sh.Range("A1").End(xlDown).Row
    ile_kolumn = sh.Range("A1").End(xlToRight).Column
    
    Dim ur As Range
    Set ur = sh.UsedRange()
    my_count = Application.WorksheetFunction.CountA(ur)
    
    If CLng(my_count) = (ile_wierszy * ile_kolumn) Then
        sprawdz_prostokatnsc_danych_w_tym_arkuszu = True
    Else
        sprawdz_prostokatnsc_danych_w_tym_arkuszu = False
        co_poszlo_zle = co_poszlo_zle & "dane w arkuszu PICKUPS sa niekompletne - wystepuja blanki" & Chr(10)
    End If
    
End Function
