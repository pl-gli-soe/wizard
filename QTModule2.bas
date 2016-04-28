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
    Dim w As Workbook, wrksh As Worksheet
    Set w = dodaj_nowy_arkusz()
    ' dodajemy nowy arkusz - nie wazne, czy sa tam jakies inne arkusze
    Set wrksh = wyodrebnij_arkusz_na_ktorym_bede_pracowal(w)
    
    
    
    ' najwygodniej zaczac od tego co wiem napewno
    ' wez z arkusza details wartosc mrd poniewaz na jej bazie bede decydowal jakie del confy sa cacy a jakie nie
    date_mrd = wez_date_mrd_z_details(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME), _
            sprawdz_czy_jest_sens_brac_date_mrd(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME)))
    
    '
    ''
    ' ======================================================
End Sub

Private Function wez_date_mrd_z_details(details_sh As Worksheet, directly_date_or_parse_from_str_mrd As Boolean) As Date
    
    If directly_date_or_parse_from_str_mrd Then
        wez_date_mrd_z_details = CDate(Format(details_sh.Cells(WizardMain.E_MRD_DATE, 2), "yyyy-mm-dd"))
    Else
        wez_date_mrd_z_details = CDate(parsuj_y_cw_do_daty_poniedzialkowej(details_sh.Cells(WizardMain.mrd, 2)))
    End If
    
    
End Function

Private Function parsuj_y_cw_do_daty_poniedzialkowej(r As Range)
    ' sekcja parsu - r to komorka zawierajaca text y cw
    
    If CStr(r) Like "Y*CW*" Then
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
    enidf
    
End Function

Private Function dodaj_nowy_arkusz() As Workbook
    Set dodaj_nowy_arkusz = Workbooks.Add()
End Function

Private Function wyodrebnij_arkusz_na_ktorym_bede_pracowal(ByRef mw As Workbook) As Worksheet
    ' tak jest najbezpieczniej
    Set wyodrebnij_arkusz_na_ktorym_bede_pracowal = mw.Sheets.Add
    wyodrebnij_arkusz_na_ktorym_bede_pracowal.Name = "workbench"
End Function
