Attribute VB_Name = "QTModule"
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
