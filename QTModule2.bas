Attribute VB_Name = "QTModule2"
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

' QT2 ma sie roznic architektura od QT1 - ma byc szybszy
' zatem musze zmienic podejscie std gdzie zbieram wszystko do obietkow, a skupiam sie bardziej
' na konkretnym zarzadzaniu danymi
' bedzie to moje mam nadzieje pierwsze makro, w ktorym nie bede podwajal pobierania danych ale
' bezposrednio zajme sie liczeniem na samyc komorkach tak jak sa
' zatem na dzien dobry zadnej klasy ani zadnych kolekcji ktore zawsze u mnie w logice dzialaly
' jako proxy dla dalszy przeliczen ktore w moim malym mozdzku uznawane byly za lepsze do ogarniecia



Public Sub QT2(ictrl As IRibbonControl)
    inner_qt2
End Sub

Private Sub inner_qt2()
    
    ' sekcja bez pivotow
    ' dopasowanie do 6p
    ' ======================================================
    ''
    '
    
    ' aby w ogole rozpoczac liczenie musze zrozumiec podstawowe definicje jakimi rzadzi sie poprzedni Quarter i w jakim cely
    ' mam w ogole zaciagac dane
    
    ' najpierw zrobmy nowy arkusz do ktorego tak jak w pierwszej generacji QT bedzie wsadzac kolejne dane
    ' jednak tym razem zrobimy to lepiej poniewaz z gory narzuce uklad kolumn taki jaki bedzie dostepny w nowym makrze 6p (nastepca Q)
    Dim w As Workbook, wrksh As Worksheet, m As Worksheet
    Set w = dodaj_nowy_arkusz()
    ' dodajemy nowy arkusz - nie wazne, czy sa tam jakies inne arkusze
    Set wrksh = wyodrebnij_arkusz_na_ktorym_bede_pracowal(w)
    Set m = ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME)
    
    
    
    ' najwygodniej zaczac od tego co wiem napewno
    ' wez z arkusza details wartosc mrd poniewaz na jej bazie bede decydowal jakie del confy sa cacy a jakie nie
    date_mrd = wez_date_mrd_z_details(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME), _
        sprawdz_czy_jest_sens_brac_date_mrd(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME)))
            
    
    
    'e_5p_total = 5
    'e_5p_na
    'e_5p_itdc
    'e_5p_pnoc
    'e_5p_fma_eur
    'e_5p_fma_osea
    'e_5p_ordered
    'e_5p_arrived
    'e_5p_in_transit
    'e_5p_ppap_status
    'e_5p_no_ppap_status
    ' piaty pieces
    wrksh.Cells(1, 1) = "5P"
    ' total
    wrksh.Cells(2, 1) = "TOTAL FMA"
    wrksh.Cells(3, 1) = _
        iteruj_recur(0, _
            przelicz_zasieg(m, WizardMain.pn, _
                WizardMain.Responsibility), _
            "FMA", _
            E_LIKE)
    
    
    Dim rng As Range
    Set rng = wrksh.Cells(2, 2)
    
    rng.Offset(-1, 0) = "RESP"
    Set rng = zrob_recursy_dla(m, rng, WizardMain.Responsibility)
    
    
    rng.Offset(-1, 0) = "PPAP STATUS"
    Set rng = zrob_recursy_dla(m, rng, WizardMain.ppap_status)
    
    
    ' 5
    wrksh.Cells(5, 1) = "6P"
    wrksh.Cells(6, 1) = "DEL CONF, WHICH IS NOT CONNECTED WITH MRD PARAM."
    Set rng = wrksh.Cells(7, 1)
    Set rng = zrob_recursy_dla(m, rng, WizardMain.Delivery_confirmation, E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD)
    
    
    ' 10
    Set rng = wrksh.Cells(11, 1)
    rng.Offset(-1, 0) = "COUNTRY CODE"
    Set rng = zrob_recursy_dla(m, rng, WizardMain.country_code)
    
    
    
    
    '
    ''
    ' ======================================================
End Sub

Private Function zrob_recursy_dla(m As Worksheet, rng As Range, resp_col, Optional e As E_SPECIAL_CASE_FOR_DEL_CONF) As Range
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Set dic = wypelnij_slownik(dic, przelicz_zasieg(m, WizardMain.pn, resp_col))
    
    For Each ki In dic.Keys
    
        If e <> E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD Then
        
            If CStr(ki) <> "" Then
                rng = ki
                rng.Offset(1, 0) = iteruj_recur(0, przelicz_zasieg(m, WizardMain.pn, resp_col), ki, E_EQUAL)
                
                Set rng = rng.Offset(0, 1)
            End If
            
            
        ElseIf e = E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD Then
            
            If CStr(ki) <> "" And Not (CStr(ki) Like "*Y*CW*") Then
                rng = ki
                rng.Offset(1, 0) = iteruj_recur(0, przelicz_zasieg(m, WizardMain.pn, resp_col), ki, E_EQUAL)
                
                Set rng = rng.Offset(0, 1)
            End If
            
            
        End If
    Next
    
    Set zrob_recursy_dla = rng
    
End Function

Private Function wypelnij_slownik(ByRef d As Dictionary, r As Range) As Dictionary
    
    Dim fst As Range, tail As Range
    
    Set fst = r.item(1)
    
    If Not d.Exists(CStr(fst)) Then
        d.Add CStr(fst), Nothing
    End If
    
    
    If r.Count > 1 Then
        Set tail = r.Parent.Range(r.item(2), r.item(r.Count))
        Set d = wypelnij_slownik(d, tail)
    End If
    
    Set wypelnij_slownik = d
    
End Function

Private Function przelicz_zasieg(m As Worksheet, col1, docelowa_kolumna) As Range

    If Trim(m.Cells(2, docelowa_kolumna)) <> "" Then
        Set przelicz_zasieg = m.Range(m.Cells(2, docelowa_kolumna), m.Cells(m.Cells(1, col1).End(xlDown).Row, docelowa_kolumna))
    Else
        Set przelicz_zasieg = m.Cells(2, docelowa_kolumna)
    End If
    

End Function

Private Function iteruj_recur(start, r As Range, filter, e As E_MATCH) As Long
    
    ' robimy rekurencje - pobierz pierwszy element zasiegu
    ' i zostaw reszte dla kolejnej rekurencji
    Dim fst As Range, tail As Range
    Set fst = r.item(1)
    
    If e = E_LIKE Then
        If fst Like "*" & CStr(filter) & "*" Then
            start = start + 1
        End If
    ElseIf e = E_EQUAL Then
        If CStr(fst) = CStr(filter) Then
            start = start + 1
        End If
    End If
    
    If r.Count > 1 Then
        Set tail = r.Parent.Range(r.item(2), r.item(r.Count))
        start = iteruj_recur(start, tail, filter, e)
    End If
    
    iteruj_recur = start
    
    
End Function



Private Function wez_date_mrd_z_details(details_sh As Worksheet, directly_date_or_parse_from_str_mrd As Boolean) As Date
    
    If directly_date_or_parse_from_str_mrd Then
        wez_date_mrd_z_details = CDate(Format(details_sh.Cells(WizardMain.E_MRD_DATE, 2), "yyyy-mm-dd"))
    Else
        wez_date_mrd_z_details = CDate(parsuj_y_cw_do_daty_poniedzialkowej(details_sh.Cells(WizardMain.mrd, 2)))
    End If
    
    
End Function

Private Function parsuj_y_cw_do_daty_poniedzialkowej(r As Range) As Date
    ' sekcja parsu - r to komorka zawierajaca text y cw
    
    If CStr(r) Like "Y*CW*" Then
        
        y = Mid(CStr(r), 2, 4)
        d_str = y & "-01-01"
        Dim d As Date
        d = CDate(d_str)
        
        Do
            cw = CLng(Right(CStr(r), Len(CStr(r)) - 7))
            
            If CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(d))) = CLng(cw) Then
                parsuj_y_cw_do_daty_poniedzialkowej = d
                Exit Do
            End If
            d = d + 1
        Loop While CLng(Year(CDate(d_str))) = CLng(y)
    Else
        MsgBox "parametr MRD jest zle zdefiniowany"
        End
    End If
End Function

Private Function sprawdz_czy_jest_sens_brac_date_mrd(details_sh As Worksheet) As Boolean
    If IsDate(details_sh.Cells(WizardMain.E_MRD_DATE, 2)) Then
        sprawdz_czy_jest_sens_brac_date_mrd = True
    Else
        sprawdz_czy_jest_sens_brac_date_mrd = False
    End If
    
End Function

Private Function dodaj_nowy_arkusz() As Workbook
    Set dodaj_nowy_arkusz = Workbooks.Add()
End Function

Private Function wyodrebnij_arkusz_na_ktorym_bede_pracowal(ByRef mw As Workbook) As Worksheet
    ' tak jest najbezpieczniej
    Set wyodrebnij_arkusz_na_ktorym_bede_pracowal = mw.Sheets.Add
    wyodrebnij_arkusz_na_ktorym_bede_pracowal.Name = "workbench"
End Function
