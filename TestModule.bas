Attribute VB_Name = "TestModule"
Public Sub test_datepicker_form()
    WizardDatePicker.show
End Sub



Public Sub more_pro_details_fill()

    Dim dh As DetailsHandler
    Set dh = New DetailsHandler
    
    
    ' arg is for is it new definition of keywords in project details
    dh.init_wizard_for_details False
    
    Set dh = Nothing
End Sub
