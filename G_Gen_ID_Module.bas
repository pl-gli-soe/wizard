Attribute VB_Name = "G_Gen_ID_Module"
Public Function generate_unique_id(r As Range) As String
    generate_unique_id = ""
    
    Do
        If CStr(r.Offset(0, -1).Value) <> CStr(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Cells(19, 1)) Then
            generate_unique_id = generate_unique_id & parse_to_num(CStr(r))
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function


Public Function parse_to_num(s As String) As String
    parse_to_num = ""
    For x = 1 To Len(s)
        parse_to_num = parse_to_num & CStr(Asc(Mid(s, x, 1)))
    Next x
    
End Function


Public Sub generuj_id(ictrl As IRibbonControl)
    Dim r As Range
    Set r = ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Range("b19")
    
    r = generate_unique_id(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Range("b1"))
    
    
End Sub
