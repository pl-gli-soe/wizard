Attribute VB_Name = "OnOffFormModule"
Public Sub on_off_form()
    ' Debug.Print "on off"
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets("config").Range("quick_forms")
    
    
    If r = 0 Then
        r = 1
        MsgBox "Formularze podpowiadajace sa teraz aktywne"
    ElseIf r = 1 Then
        r = 0
        MsgBox "Formularze podpowiadajace sa teraz nieaktywne"
    Else
        MsgBox "cfg quick forms - ten msgbox nigdy nie powinien sie pokazac!"
    End If

End Sub


Public Sub on_off_quick_form(ictrl As IRibbonControl)
    on_off_form
End Sub
