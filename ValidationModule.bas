Attribute VB_Name = "ValidationModule"
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

Public Sub waliduj(ictrl As IRibbonControl)
    my_validation ThisWorkbook.Sheets(MASTER_SHEET_NAME)
End Sub


Public Sub my_validation(sh As Worksheet)


    dodatkowa_walidacja_arkusza_pusow False
    dodatkowa_walidacja_arkusza_master

End Sub

Public Sub dodatkowa_walidacja_arkusza_pusow(Optional s As Boolean)

    If s = True Then
        Dim v As MasterValidation
        Set v = New MasterValidation
        bool = v.make_validation_on_all_pickups_in_system(ThisWorkbook.Sheets(PICKUPS_SHEET_NAME))
        v.pokaz_wynik_walidacji
        
        Set v = Nothing
    End If
End Sub

Private Sub dodatkowa_walidacja_arkusza_master()
    Dim v As MasterValidation
    Set v = New MasterValidation
    v.waliduj_arkusz_master
    v.pokaz_wynik_walidacji
    
    Set v = Nothing
End Sub
