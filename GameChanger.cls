VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

' ta klasa odpowiadac bedzie za zmiany na calym arkuszu
' jednak przede wszytkim mam na mysli synchro zmian na arkuszu PUS i MASTER

Private m As Worksheet
Private p As Worksheet


' selected range
Private sr As Range



Private Sub Class_Initialize()
    Set m = ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME)
    Set p = ThisWorkbook.Sheets(WizardMain.PICKUPS_SHEET_NAME)
End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set p = Nothing
End Sub


Public Sub change_deck_on_selection(deck As String)

    If sprawdz_czy_napewno_ten_string_to_deck(CStr(deck)) Then
        
        podmien_na_wybranej_selekcji_w_arkuszu_master_wszystkie_decki CStr(deck)
    End If
End Sub


Private Function sprawdz_czy_napewno_ten_string_to_deck(deck As String) As Boolean

    If Len(deck) = 2 Then
        sprawdz_czy_napewno_ten_string_to_deck = True
    Else
        sprawdz_czy_napewno_ten_string_to_deck = False
    End If
End Function


Private Sub podmien_na_wybranej_selekcji_w_arkuszu_master_wszystkie_decki(deck As String)
    ' no to lecimy jesli activesheet to master
    
    Set sr = Selection
    ' drugi check tego samego
    ' wczesniej byl w module glownym tej klasy
    ' (game changer)
    If sr.Parent.Name = WizardMain.MASTER_SHEET_NAME Then
        
        Dim r As Range
        For Each r In sr
            
            If Not r.EntireRow.Hidden Then
                If Not r.EntireColumn.Hidden Then
                    iteracja_pojedynczego_elementu_dla_zmiany_decku r, CStr(deck)
                End If
            End If
        Next r
    End If
End Sub

Private Sub iteracja_pojedynczego_elementu_dla_zmiany_decku(r As Range, deck As String)
    
    With m
        
    
        mypn = remove_zeros_from_left(CStr(.Cells(r.Row, WizardMain.pn)))
        myduns = remove_zeros_from_left(CStr(.Cells(r.Row, WizardMain.duns)))
        
        myindx = mypn & "," & myduns & "," & CStr(.Cells(r.Row, WizardMain.fup_code).Value)
        
        przejrzyj_teraz_arkusz_pickups_i_podmien_rowniez_decki CStr(myindx), CStr(deck)
        
        .Cells(r.Row, WizardMain.fup_code).Value = deck
    End With
        
End Sub


Private Function remove_zeros_from_left(s As String)
    
    If Left(s, 1) = "0" Then
        s = Mid(s, 2, Len(s) - 1)
    Else
        remove_zeros_from_left = s
        Exit Function
    End If
    
    remove_zeros_from_left = remove_zeros_from_left(s)
End Function


Private Sub przejrzyj_teraz_arkusz_pickups_i_podmien_rowniez_decki(myindx As String, mydeck As String)
    
    Dim i As Range
    Set i = p.Range("a2")
    Do
        If Trim(CStr(i)) = myindx Then
            podmien_w_tej_linii_arkusza_pickups_deck i, CStr(mydeck)
        End If
        nowy_schemat_offsetu_w_arkuszu_pickups i
    Loop Until i.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
End Sub

Private Sub podmien_w_tej_linii_arkusza_pickups_deck(i As Range, deck As String)
    
    p.Cells(i.Row, WizardMain.O_FUP_code) = deck
    p.Cells(i.Row, WizardMain.O_INDX) = p.Cells(i.Row, WizardMain.O_PN) & "," & p.Cells(i.Row, WizardMain.O_DUNS) & "," & deck
End Sub
