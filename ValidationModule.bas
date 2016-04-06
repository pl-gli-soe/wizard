Attribute VB_Name = "ValidationModule"
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
