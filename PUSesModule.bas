Attribute VB_Name = "PUSesModule"
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

Public Sub dodaj_pickup_dla_jednego_pna(ictrl As IRibbonControl)
    add_pickup
End Sub

Public Sub edytuj_pickup(ictrl As IRibbonControl)
    edit_pickup
End Sub


Public Sub add_pickup()

    With FormPickups
        .show vbModeless
        
        ' z racji pojawienia sie funkcji show mozemy zmienic wartosc w rejestrze
        ' ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        
        ' uzupelnienie de novo checkboxa
        .get_pickups_handler.fill_source_checkbox
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        .get_pickups_handler.adjust_content_if_selection_changed
        .am_i_visible = True
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
    End With
End Sub

Public Sub del_pickup_data_from_jenny(adr)


    Dim jr As Range
    Set jr = myReference(CStr(adr))

    If usun_wybrany_pickup_bez_potwierdzenia_i_info_zwrotnego(CStr(jr.item(7, 2))) Then
        MsgBox "PUS zostal usuniety z danego Wizarda"
    Else
        MsgBox "nie bylo czego usuwac"
    End If
End Sub

Public Sub add_pickup_data_from_jenny(adr)

    Dim jr As Range
    Set jr = myReference(CStr(adr))
    
    

    ' pamietaj ze jr to duzy range od a1 do p200

    With FormPickups
        .show vbModeless
        
        ' z racji pojawienia sie funkcji show mozemy zmienic wartosc w rejestrze
        ' ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        
        ' uzupelnienie de novo checkboxa
        .get_pickups_handler.disable_checkbox
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        .get_pickups_handler.setJr jr
        .get_pickups_handler.adjust_content_if_selection_changed
        .am_i_visible = True
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
    End With
End Sub

Public Sub edit_pickup_with_jenny(adr)

    Dim jr As Range
    Set jr = myReference(CStr(adr))
    
    
    If usun_wybrany_pickup_bez_potwierdzenia_i_info_zwrotnego(CStr(jr.item(7, 2))) Then

        ' pamietaj ze jr to duzy range od a1 do p200
        
        ' jakis pus zostal usuniety zatem teraz trzeba go na nowo dodac!
    
        With FormPickups
            .show vbModeless
            
            ' z racji pojawienia sie funkcji show mozemy zmienic wartosc w rejestrze
            ' ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
            
            ' uzupelnienie de novo checkboxa
            .get_pickups_handler.disable_checkbox
            ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
            .get_pickups_handler.setJr jr
            .get_pickups_handler.adjust_content_if_selection_changed
            .am_i_visible = True
            ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        End With
    Else
        ' MsgBox "nie ma czego edytowac!"
        add_pickup_data_from_jenny adr
    End If
End Sub

' http://stackoverflow.com/questions/13950816/convert-external-cell-address-into-range-in-vba
Private Function myReference(strAddress As String) As Range

    Dim intPos As Integer, intPos2 As Integer
    Dim strWB As String, strWS As String, strCell As String

    intPos = InStr(strAddress, "]")
    strWB = Mid(strAddress, 3, intPos - 3)

    intPos2 = InStr(strAddress, "!")
    strWS = Mid(strAddress, intPos + 1, intPos2 - intPos - 2)

    strCell = Mid(strAddress, intPos2 + 1)

    ' w orginale bylo bez set'a
    Set myReference = Workbooks(strWB).Worksheets(strWS).Range(strCell)

End Function


Public Sub edit_pickup()

    ' tutaj niby logika ulopadu z external listy
    ' MsgBox "jeszcze nie zaimplementowane!"
    
    With FormEditPuses
        .show vbModeless
        
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        .am_i_visible = True
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
        
        .get_pickups_handler.adjust_content_if_selection_changed
        
    End With
End Sub

Public Sub usun_zapisane_pickupy(ictrl As IRibbonControl)

    
    a = MsgBox("Uzytkowniku! Czy jestes absolutnie pewien tego co robisz?", vbCritical + vbYesNo)
    
    
    
    If a = vbYes Then
    
        haslo = InputBox("Wpisz klucz dostepu", "klucz dostepu", "0000-00-00")
        
        If CStr(haslo) = CStr(G_PASS) Then
    
        
            Dim r As Range, psh As Worksheet
            Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
            'psh.Unprotect 123
            Set r = psh.Range(psh.Cells(2, WizardMain.O_INDX), psh.Cells(WizardMain.CAPACITY_ARKUSZA, WizardMain.O_PUS_Number))
            r.Clear
            r.Value = ""
            
            'psh.Protect 123
        Else
            MsgBox "pass nie trafiony - Twoj komputer za chwile wybuchnie :)"
        End If
    Else
        MsgBox "dane nie zostana usuniete", vbInformation
    End If
End Sub

Public Sub usun_wybrany_pickup(ictrl As IRibbonControl)
    
    
    If ThisWorkbook.Sheets(PICKUPS_SHEET_NAME).Name = ActiveSheet.Name Then
        If ActiveCell.Row = 1 Then
            MsgBox "nie mozesz usunac nazwy kolumn baranku :)", vbInformation
        Else
        
            If CStr(Trim(Cells(ActiveCell.Row, WizardMain.O_PUS_Number))) = "" Then
                MsgBox "nie mozna usunac czegos, co nie istnieje - wybierz wiersz z konkretnym pusem :)", vbInformation
            Else
                a = MsgBox("Czy chcesz usunac PUS #: " & Cells(ActiveCell.Row, WizardMain.O_PUS_Number) & "?", vbCritical + vbYesNo)
                
                If a = vbYes Then
                    ' przejrzyj caly arkusz PICKUPS i usun wybrany pus
                    
                    usun_wybrany_pickup_po_potwierdzeniu CStr(Trim(Cells(ActiveCell.Row, WizardMain.O_PUS_Number)))
                ElseIf a = vbNo Then
                    ' no operation at last
                Else
                    MsgBox "ten msgbox niegdy nie powinien sie pokazac - usun wybrany pickup", vbCritical
                End If
            End If
        End If
    ElseIf ThisWorkbook.Sheets(MASTER_SHEET_NAME).Name = ActiveSheet.Name Then
    
        ' tutaj wazna dodatkowa logika zwiazana z mozliwoscia
        ' usuwania jak i edytowania wybranych pusow
        ' kwestia okreslenia w jaki sposob to zrobimy jeszcze
        
    Else
        MsgBox "na tym arkuszu nie mozesz usuwac pusow :) - przejdz do arkusza PICKUPS", vbInformation
    End If
End Sub

Public Sub usun_wybrany_pickup_po_potwierdzeniu(pus_name As String)

    Dim r As Range, psh As Worksheet
    Dim czy_cos_zostalo_usuniete As Boolean
    czy_cos_zostalo_usuniete = False
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    'psh.Unprotect 123
    
    Set r = psh.Cells(2, WizardMain.O_PUS_Number)
    Do
        If CStr(r) = pus_name Then
            For x = WizardMain.O_INDX To WizardMain.O_PUS_Number
                psh.Cells(r.Row, x).Value = ""
            Next x
            
            'Set r = r.Offset(1, 0)
            czy_cos_zostalo_usuniete = True
        Else
            ' Set r = r.Offset(1, 0)
        End If
        
        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
    Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    
    If czy_cos_zostalo_usuniete Then
        MsgBox "dane zostaly usuniete!"
    Else
        MsgBox "nie ma czego usuwac!"
    End If
    
    'psh.Protect 123
End Sub

Public Function usun_wybrany_pickup_bez_potwierdzenia_i_info_zwrotnego(pus_name As String) As Boolean

    Dim r As Range, psh As Worksheet
    Dim czy_cos_zostalo_usuniete As Boolean
    czy_cos_zostalo_usuniete = False
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    'psh.Unprotect 123
    
    Set r = psh.Cells(2, WizardMain.O_PUS_Number)
    Do
        If CStr(r) = pus_name Then
            For x = WizardMain.O_INDX To WizardMain.O_PUS_Number
                psh.Cells(r.Row, x).Value = ""
            Next x
            
            'Set r = r.Offset(1, 0)
            czy_cos_zostalo_usuniete = True
        Else
            ' Set r = r.Offset(1, 0)
        End If
        
        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
    Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    
    If czy_cos_zostalo_usuniete Then
    '    MsgBox "dane zostaly usuniete!"
        usun_wybrany_pickup_bez_potwierdzenia_i_info_zwrotnego = True
    Else
    '    MsgBox "nie ma czego usuwac!"
        usun_wybrany_pickup_bez_potwierdzenia_i_info_zwrotnego = False
    End If
    
    
    'psh.Protect 123
End Function
