Attribute VB_Name = "ZmienDeckModule"
Public Sub zmien_deck(ictrl As IRibbonControl)
    If Selection.Parent.Name = WizardMain.MASTER_SHEET_NAME Then
        FUPForm.show
    Else
        MsgBox "selekcja nie jest zrobiona w arkuszu MASTER"
    End If
End Sub
