Attribute VB_Name = "NewEditProjectModule"
Public Sub nowy_projekt(ictrl As IRibbonControl)
    With FormStart
        .Caption = "NEW"
        .BtnDetails.Caption = "Dodaj detale projektu"
        .show
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
    End With
End Sub


Public Sub edytuj_projekt(ictrl As IRibbonControl)
    With FormStart
        .Caption = "EDIT"
        .BtnDetails.Caption = "Edytuj detale projektu"
        .show
        ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 1
    End With
End Sub
