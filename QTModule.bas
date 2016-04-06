Attribute VB_Name = "QTModule"
Public Sub quarter_time_inner(current_form As Variant)

    Dim wb As Workbook, mrd_date As Date, custom_date As Date, today_date As Date, ktory_dzien As Integer
    
    
    Select Case CStr(current_form.ComboBox1.Value)
        Case "Poniedzialek"
            ktory_dzien = 1
        Case "Wtorek"
            ktory_dzien = 2
        Case "Sroda"
            ktory_dzien = 3
        Case "Czwartek"
            ktory_dzien = 4
        Case "Piatek"
            ktory_dzien = 5
        Case "Sobota"
            ktory_dzien = 6
        Case "Niedziela"
            ktory_dzien = 7
    End Select
    
    mrd_date = 2
    custom_date = 2
    today_date = 2
    
    
    With current_form
        
        If .CheckBoxPivotInTransitCustomDate.Value = True Then
            custom_date = CDate(.DTPickerDataPodzialuInTransit.Value)
        End If
        
        
        If .CheckBoxPivotInTransitTODAY.Value = True Then
            today_date = Date
        End If
        
        ' mrd_date - narazie nie mam dostepu do danych - pamietaj ze to sie wpisuje we flakach
        If .CheckBoxPivotInTransitMRD.Value = True Then
            ' taki maly trick
            mrd_date = 3
        End If
    End With
    
    Set wb = Workbooks.Add
    
    Dim qh As QTHandler
    Set qh = New QTHandler
    
    With qh
        .init wb, mrd_date, custom_date, today_date, Int(ktory_dzien)
        .count_pns_and_fill_qi_collection
        
        ' .test_qi_collection
        
        .fill_new_workbook wb
    End With
    
    
    MsgBox "gotowe!"
    Set qh = Nothing

End Sub


Public Sub quarter_time(ictrl As IRibbonControl)
    QTForm.ComboBox1.Clear
    
    With QTForm
        With .ComboBox1
            
            .AddItem "Poniedzialek"
            .AddItem "Wtorek"
            .AddItem "Sroda"
            .AddItem "Czwartek"
            .AddItem "Piatek"
            .AddItem "Sobota"
            .AddItem "Niedziela"
            
            
            .Value = "Poniedzialek"
        End With
        
        .CheckBoxPivotInTransitMRD.Value = 1
        .CheckBoxPivotInTransitTODAY.Value = 1
        .CheckBoxPivotInTransitCustomDate.Value = 1
        
        .DTPickerDataPodzialuInTransit = Date
    End With
    
    QTForm.show
End Sub
