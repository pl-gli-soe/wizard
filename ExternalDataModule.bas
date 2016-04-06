Attribute VB_Name = "ExternalDataModule"
Public Sub dopisz_dane(ictrl As IRibbonControl)
    inner_dopisz
End Sub


Public Sub nadpisz_dane(ictrl As IRibbonControl)
    add_data E_NADPISZ
End Sub



Public Sub add_data(how As E_ADD_DATA)

    ' MsgBox "logika tej akcji nie zostala jeszcze stworzona"
    
    If how = E_NADPISZ Then
        PodmienForm.show
    End If
    
End Sub


Public Sub inner_dopisz()
    DopiszForm.show
End Sub
