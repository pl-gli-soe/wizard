VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZmienPUSaForm 
   Caption         =   "Zmien"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "ZmienPUSaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZmienPUSaForm"
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


Private stara_nazwa_pusa As String
Private pus_date As Date
Private del_date As Date


Private psh As Worksheet


Private Sub BtnSubmit_Click()

    ' submit
    ' ================================
    
    If waliduj_pola_arkusza() Then
    
        Me.hide
        podmien_w_arkuszu_na_nowe_dane
        edit_pickup
    Else
        MsgBox "dane wpisane w ten formularz nie sa wlasciwie!"
    End If
    
    ' ================================
End Sub

Private Function waliduj_pola_arkusza() As Boolean

    waliduj_pola_arkusza = True
    
    If Me.TextBoxPUSName = "" Then
        waliduj_pola_arkusza = False
    End If
    
    If sprawdz_czy_juz_jest_taki_pus() Then
        waliduj_pola_arkusza = False
    End If
    
    If CDate(Me.DTPickerDelDate) < CDate(Me.DTPickerPUSDate) Then
        waliduj_pola_arkusza = False
    End If
End Function

Private Function sprawdz_czy_juz_jest_taki_pus() As Boolean
    sprawdz_czy_juz_jest_taki_pus = False
    
    Dim psh As Worksheet, r As Range
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set r = psh.Range("A2")
    Do
        
        If Me.TextBoxPUSName <> Me.TextBoxPUSName2 Then
            If r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX).Value = Me.TextBoxPUSName Then
                sprawdz_czy_juz_jest_taki_pus = True
                Exit Function
            End If
        End If
        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
    Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    
End Function

Private Sub podmien_w_arkuszu_na_nowe_dane()

    Dim r As Range
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set r = psh.Cells(2, WizardMain.O_PUS_Number)
    
    ' lista nie jest pusta
    If r <> "" Then
    
        Do
            If CStr(r) = Me.TextBoxPUSName2 Then
                r = Me.TextBoxPUSName.Value
                r.Offset(0, WizardMain.O_Pick_up_date - WizardMain.O_PUS_Number).Value = Me.DTPickerPUSDate.Value
                r.Offset(0, WizardMain.O_Delivery_Date - WizardMain.O_PUS_Number).Value = Me.DTPickerDelDate.Value
            End If
            WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
        Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    End If
End Sub
