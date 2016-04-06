Attribute VB_Name = "FindNewModule"
Public Sub znajdz_te_co_maja_byc_dopisane()


    Dim w As Workbook

    With FindForm
    
        .ListBoxW.Clear
        .ListBoxS.Clear
    
        For Each w In Workbooks
            If Trim(w.Name) <> "" Then
                .ListBoxW.AddItem w.Name
            End If
        Next w
        
        
        .TextBoxPNCol.Value = 1
        Set .ref_find_new_handler = New FindNewHandler
        .show
    End With

End Sub
