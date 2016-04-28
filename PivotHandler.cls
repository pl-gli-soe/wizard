VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private pivot_table_cache As PivotCache
Private pivot_table As PivotTable

Private wb As Workbook
Private pivot_sh As Worksheet

Private mrd_date As Date
Private today_date As Date
Private custom_date As Date


Private adres As String


Public Sub init_(ByRef w As Workbook, sh As Worksheet, e As E_PIVOTS, md As Date, td As Date, cd As Date)


    mrd_date = md
    today_date = td
    custom_date = cd

    Set wb = w
    Set pivot_sh = sh
    
    ' adres jest juz przygotowany w poprzedniej procedurze
    
    Set pivot_table_cache = Nothing
    Set pivot_table_cache = wb.PivotCaches.Create(xlDatabase, SourceData:=adres)
    ' to nie dziala
    ' pivot_sh.PivotTables.Clear
    
    'For Each pt In pivot_sh.PivotTables
    '    pt.TableRange2.Clear
    'Next pt
    
    Dim r As Range
    
    
    If e = E_PIVOT_COUNTRY_CODE Then
        pivot_sh.Range("A3") = "COUNTRY CODE"
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, pivot_sh.Range("A4"))
        
    ElseIf e = E_PIVOT_FUP_CODE Then
        pivot_sh.Range("F3") = "FOLLOW UP CODE"
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, pivot_sh.Range("F4"))
        
        
    ' =============================================================================================
    ' =============================================================================================
    ElseIf e = E_PIVOT_INTRANSIT_TODAY Then
        pivot_sh.Range("K3") = "INTRANSIT TODAY: " & CStr(today_date)
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, pivot_sh.Range("K4"))
        
    ' =============================================================================================
        
    ElseIf e = E_PIVOT_INTRANSIT_MRD Then
    
        If pivot_sh.Range("K3") = "" Then
            Set r = pivot_sh.Range("K3")
            r = "INTRANSIT MRD: " & CStr(mrd_date)
        ElseIf pivot_sh.Range("K3").End(xlDown).Offset(4, 0) = "" Then
            Set r = pivot_sh.Range("K3").End(xlDown).Offset(4, 0)
            r = "INTRANSIT MRD: " & CStr(mrd_date)
        Else
            Set r = pivot_sh.Range("K3").End(xlDown).End(xlDown).End(xlDown).Offset(4, 0)
            r = "INTRANSIT MRD: " & CStr(mrd_date)
        End If
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, r.Offset(1, 0))
        
        
        
    ElseIf e = E_PIVOT_INTRANSIT_CUSTOM_DATE Then
        If pivot_sh.Range("K3") = "" Then
            Set r = pivot_sh.Range("K3")
            r = "INTRANSIT CUSTOM DATE: " & CStr(custom_date)
            
        ElseIf pivot_sh.Range("K3").End(xlDown).Offset(4, 0) = "" Then
            Set r = pivot_sh.Range("K3").End(xlDown).Offset(4, 0)
            r = "INTRANSIT CUSTOM DATE: " & CStr(custom_date)
        Else
            Set r = pivot_sh.Range("K3").End(xlDown).End(xlDown).End(xlDown).Offset(4, 0)
            r = "INTRANSIT CUSTOM DATE: " & CStr(custom_date)
        End If
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, r.Offset(1, 0))
        
        
    ' =============================================================================================
    ' =============================================================================================
        
    ElseIf e = E_PIVOT_PPAP_STATUS Then
        pivot_sh.Range("P3") = "PPAP_STATUS"
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, pivot_sh.Range("P4"))
        
    ElseIf e = E_PIVOT_RESP Then
        pivot_sh.Range("U3") = "RESP"
        Set pivot_table = pivot_sh.PivotTables.Add(pivot_table_cache, pivot_sh.Range("U4"))
    End If
    
    
End Sub

Public Sub no_to_rysuj(e As E_PIVOTS)


    'r.Offset(-1, 8) = "getInTransitMRD_" & CStr(mrd_date)
    'r.Offset(-1, 8) = "getInTransitToday_" & CStr(today_date)
    'r.Offset(-1, 8) = "getInTransitCustom_" & CStr(custom_date)


    Dim tmp_txt As String
    If e = E_PIVOT_RESP Then
        tmp_txt = "RESP"
    ElseIf e = E_PIVOT_COUNTRY_CODE Then
        tmp_txt = "COUNTRY_CODE"
    ElseIf e = E_PIVOT_FUP_CODE Then
        tmp_txt = "FUP_CODE"
    ElseIf e = E_PIVOT_INTRANSIT_MRD Then
        tmp_txt = "getInTransitMRD_" & CStr(mrd_date)
    ElseIf e = E_PIVOT_INTRANSIT_TODAY Then
        tmp_txt = "getInTransitToday_" & CStr(today_date)
    ElseIf e = E_PIVOT_INTRANSIT_CUSTOM_DATE Then
        tmp_txt = "getInTransitCustom_" & CStr(custom_date)
    ElseIf e = E_PIVOT_PPAP_STATUS Then
        tmp_txt = "PPAP_STATUS"
    Else
        MsgBox "cos poszlo nie tak przy: no to rysuj"
    End If
    
    With pivot_table
    
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
    
        .PivotFields("RESP").Orientation = xlPageField
        .PivotFields("RESP").Position = 1
        
    
        If tmp_txt <> "RESP" Then
        
            .PivotFields(tmp_txt).Orientation = xlRowField
            .PivotFields(tmp_txt).Position = 1
            .PivotFields("DEL_CONF").Orientation = xlRowField
            .PivotFields("DEL_CONF").Position = 2
            
        Else
            
            .PivotFields("DEL_CONF").Orientation = xlRowField
            .PivotFields("DEL_CONF").Position = 1
            
        End If
        
        With .PivotFields("PN")
            .Orientation = xlDataField
            .Function = xlCount
        End With
    End With
End Sub

Public Sub przygotuj_adres(s As Worksheet)


    adr = "A1"
    
    Dim lr As Long
    lr = CLng(s.Range("a1").End(xlDown).Row)
    ' tutaj super statycznie - pamietaj ze to tez trzeba zmienic, gdy cos wyszlo do poprawy
    adr = adr & ":M" & CStr(lr)
    adres = s.Name & "!" & adr
End Sub


Private Sub fake_sub()
    'With pivot_table
    '    .PivotFields("plant").Orientation = xlRowField
    '    .PivotFields("plant").Position = 1
    '
    '    .PivotFields("part number").Orientation = xlRowField
    '    .PivotFields("part number").Position = 2
    '
    '    .PivotFields("name").Orientation = xlRowField
    '    .PivotFields("name").Position = 3
    '
    '    '.PivotFields("regular transport").Orientation = xlRowField
    '    '.PivotFields("regular transport").Position = 4
    '
    '    .PivotFields("pickup date").Orientation = xlColumnField
    '    .PivotFields("pickup date").Position = 1
    '
    '    .PivotFields("qty for this transport").Orientation = xlDataField
    '
    '    .PivotFields("regular transport").Orientation = xlPageField
    '    .PivotFields("regular transport").Position = 1
    '
    '    .ColumnGrand = False
    '    .RowGrand = False
    '
    '    .TableStyle2 = "PivotStyleMedium6"
    '
    '
    '    .PivotFields("plant").LayoutBlankLine = _
    '        True
    '    .PivotFields("part number"). _
    '        LayoutBlankLine = True
    '    .PivotFields("name").LayoutBlankLine = _
    '        True
    '    .PivotFields("delivery date"). _
    '        LayoutBlankLine = True
    '    .PivotFields("pickup date"). _
    '        LayoutBlankLine = True
    '    .PivotFields("qty for this transport"). _
    '        LayoutBlankLine = True
    '    .PivotFields("value in cell"). _
    '        LayoutBlankLine = True
    '    .PivotFields("difference"). _
    '        LayoutBlankLine = True
    '    .PivotFields("valid change"). _
    '        LayoutBlankLine = True
    '    .PivotFields("regular transport"). _
    '        LayoutBlankLine = True
    '    .PivotFields("qty for this transport"). _
    '        LayoutBlankLine = True
    'End With
End Sub
