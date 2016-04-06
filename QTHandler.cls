VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QTHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private w As Workbook
Private how_many_pns_you_have As Long

Private qi As QTItem
Private qi_collection As Collection

Private mrd_date As Date
Private custom_date As Date
Private today_date As Date


Private pod_mrd_ktory_dzien As Integer


Private Sub Class_Initialize()
    Set w = Nothing
    Set qi_collection = New Collection
End Sub

Private Sub Class_Terminate()
    Set w = Nothing
    Set qi_collection = Nothing
End Sub


Public Sub init(ByRef wb As Workbook, m_mrd_date As Date, m_custom_date As Date, m_today_date As Date, m_ktory_dzien As Integer)
    Set w = wb
    
    mrd_date = m_mrd_date
    custom_date = m_custom_date
    today_date = m_today_date
    
    pod_mrd_ktory_dzien = m_ktory_dzien
End Sub

Public Sub count_pns_and_fill_qi_collection()
    
    Dim r As Range, m As Worksheet, p As Worksheet
    Set p = ThisWorkbook.Sheets(WizardMain.PICKUPS_SHEET_NAME)
    Set m = ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME)
    Set r = m.Range("a2")
    
    ile_numerow_okolo = r.End(xlDown).Offset(1, 0).Row - 1
    
    Dim gp As StatusHandler
    Set gp = New StatusHandler
    gp.init_statusbar ile_numerow_okolo
    
    gp.show
    
    ' Dim fst_mrd As Date
    If mrd_date = 3 Then
        ' kilka rzeczy
        ' 2 jesli w formularzu wyrazilem chec zobaczenia tego pivota
        ' firstMRDDate_prepare_from_details() + (m_ktory_dzien - 1)
        ' dodatek po prawej stronie jest po to by nie zmieniac zbytnio pierwotnej funkcji
        ' dodaje tylko tyle dni ile sam wybralem :)
        mrd_date = firstMRDDate_prepare_from_details() + (pod_mrd_ktory_dzien - 1)
    End If
    
    
    With m
        Do
        
        
            Set qi = Nothing
            Set qi = New QTItem
            
            
            
            On Error Resume Next
            qi.pn = m.Cells(r.Row, WizardMain.pn)
            On Error Resume Next
            qi.duns = m.Cells(r.Row, WizardMain.duns)
            On Error Resume Next
            qi.supp_nm = m.Cells(r.Row, WizardMain.Supplier_Name)
            On Error Resume Next
            qi.del_conf = m.Cells(r.Row, WizardMain.Delivery_confirmation)
            On Error Resume Next
            qi.fup_code = m.Cells(r.Row, WizardMain.fup_code)
            On Error Resume Next
            qi.country_code = m.Cells(r.Row, WizardMain.country_code)
            On Error Resume Next
            qi.ppap_status = m.Cells(r.Row, WizardMain.ppap_status)
            On Error Resume Next
            qi.resp = m.Cells(r.Row, WizardMain.Responsibility)
            
            If m.Cells(r.Row, WizardMain.Total_PUS_QTY) > 0 Then
            
            
                Set qi.pus_bucket = Nothing
                Set qi.pus_bucket = New PUSBUcket
                qi.pus_bucket.init mrd_date, custom_date, today_date
                
                qi.pus_bucket.setIndx CStr(m.Cells(r.Row, WizardMain.pn)) & "," & _
                    CStr(m.Cells(r.Row, WizardMain.duns)) & "," & _
                    CStr(m.Cells(r.Row, WizardMain.fup_code))
                qi.pus_bucket.search_pus_sheet_and_prepare_bucket p
            End If
            
            
            qi_collection.Add qi
            Set r = r.Offset(1, 0)
            gp.progress_increase
        Loop Until Trim(r) = ""
    End With
    
    gp.hide
    Set gp = Nothing
    
    how_many_pns_you_have = r.Row - 1
End Sub


Public Sub test_qi_collection()
    
    Set qi = Nothing
    For Each qi In qi_collection
        work_with_this_qi qi
    Next qi
End Sub

Private Sub work_with_this_qi(ByRef qi As QTItem)


    Dim pi As PUSItemForPUSBucket
    
    If Not qi.pus_bucket Is Nothing Then
        If Not qi.pus_bucket.pus_collection Is Nothing Then
            If qi.pus_bucket.pus_collection.Count > 0 Then
            
                For Each v In qi.pus_bucket.pus_collection
                    Set pi = v
                    
                    ' Debug.Print pi.duns & " " & pi.eda_date & " " & pi.fup_code & " " & pi.nm
                Next v
            End If
        End If
    End If
End Sub


Private Function firstMRDDate_prepare_from_details() As Date


    Dim tmp_date As Date

    Dim tmp As Worksheet
    Set tmp = ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME)
    Dim mrd_range As Range
    Set mrd_range = tmp.Range("mrd")
    
    Dim txt As String
    txt = Replace(Replace(mrd_range.Value, "Y", ""), "CW", "")
    
    If Len(txt) < 6 Then
    
        If Len(txt) = 5 Then
            'MsgBox "nie powinno sie pokazac - skontaktuj sie z autorem " & Chr(10) & _
            '    "firstMRDDate_prepare_from_details"
            txt = Left(txt, 4) & "0" & Right(txt, 1)
        Else
            MsgBox "cos poszlo nie tak w QT proces"
            End
        End If
    End If
    
    If Len(txt) = 6 Then
        If IsNumeric(txt) Then
            
            
            ' mozemy narszcie przejsc do sedna!
            ' -------------------------------------------------------------
            yyyy = Left(txt, 4)
            cw = Right(txt, 2)
            
            tmp_date = CDate(CStr(yyyy) & "-" & CStr(Month(Date)) & "-" & CStr(Day(Date)))
            jaki_dzien_tygodnia = Weekday(tmp_date, vbMonday)
            tmp_date = CDate(tmp_date - jaki_dzien_tygodnia + 1)
            
            tmp_date = CDate(get_month_and_day_from(tmp_date, CStr(yyyy), CLng(cw)))
            
            
            firstMRDDate_prepare_from_details = tmp_date
            ' -------------------------------------------------------------
        Else
            MsgBox "cos poszlo nie tak w QT proces"
            End
        End If
    Else
        MsgBox "cos poszlo nie tak w QT proces"
        End
    End If
    
    
End Function

Private Function get_month_and_day_from(data As Date, yyyy As String, cw As Long) As Date

    
    
    If CLng(Application.WorksheetFunction.IsoWeekNum(data)) < cw Then
    
        If CLng(Year(data)) < CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data + 7, yyyy, cw)
        ElseIf CLng(Year(data)) > CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data - 7, yyyy, cw)
        ElseIf CLng(Year(data)) = CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data + 7, yyyy, cw)
        Else
            MsgBox "cos poszlo nie tak w QT proces"
        End If
    ElseIf CLng(Application.WorksheetFunction.IsoWeekNum(data)) > cw Then
    
        If CLng(Year(data)) > CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data - 7, yyyy, cw)
        ElseIf CLng(Year(data)) < CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data + 7, yyyy, cw)
        ElseIf CLng(Year(data)) = CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data - 7, yyyy, cw)
        Else
            MsgBox "cos poszlo nie tak w QT proces"
        End If
        
    ElseIf CLng(Application.WorksheetFunction.IsoWeekNum(data)) = CLng(cw) Then
    
        If CLng(Year(data)) = CLng(yyyy) Then
            get_month_and_day_from = CDate(data)
            Exit Function
        ElseIf CLng(Year(data)) > CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data - 7, yyyy, cw)
        ElseIf CLng(Year(data)) < CLng(yyyy) Then
            get_month_and_day_from = get_month_and_day_from(data + 7, yyyy, cw)
        End If
    Else
        MsgBox "cos poszlo nie tak w QT proces - get_month_and_day_from"
    End If
    
End Function

' subroutine for new workbook and to fill content from qi_collection
Public Sub fill_new_workbook(ByRef wb As Workbook)
    Set w = wb
    
    Dim gp As StatusHandler
    Set gp = New StatusHandler
    gp.init_statusbar 4
    
    gp.show
    
    gp.progress_increase
    
    ' psh - pickups sheet
    Dim sh_source As Worksheet, sh_pivot As Worksheet, psh As Worksheet
    Set sh_source = w.ActiveSheet
    sh_source.Name = "FlatTable"
    
    Dim r As Range
    Set r = sh_source.Range("a2")
    
    przygotuj_labelki r
    
    iteracja_pierwsza_zrob_flat_table w, r
    gp.progress_increase
    
    Set sh_pivot = w.Sheets.Add
    sh_pivot.Name = "Pivots"
    iteracja_przygotuj_tabele_przestawne sh_source, sh_pivot
    gp.progress_increase
    
    Set psh = w.Sheets.Add
    psh.Name = "PUSes"
    wyrzuc_pusy psh
    gp.progress_increase
    
    gp.hide
    Set gp = Nothing
End Sub

Private Sub wyrzuc_pusy(psh As Worksheet)


    'Dim pi As PUSItemForPUSBucket
   '
   ' If Not qi.pus_bucket Is Nothing Then
   '     If Not qi.pus_bucket.pus_collection Is Nothing Then
   '         If qi.pus_bucket.pus_collection.Count > 0 Then
   '
   '             For Each v In qi.pus_bucket.pus_collection
   '                 Set pi = v
   '
   '                 Debug.Print pi.duns & " " & pi.eda_date & " " & pi.fup_code & " " & pi.nm
   '             Next v
   '         End If
   '     End If
   ' End If


    Dim r As Range
    Dim pi As PUSItemForPUSBucket
    
    Set r = psh.Cells(1, 1)
    zrob_labelki_dla_wyrzuconych_pusow r
    Set r = r.Offset(1, 0)

    Set qi = Nothing
    For Each qi In qi_collection
    
        If Not qi.pus_bucket Is Nothing Then
            If Not qi.pus_bucket.pus_collection Is Nothing Then
                If qi.pus_bucket.pus_collection.Count > 0 Then
                    For Each v In qi.pus_bucket.pus_collection
                    
                        Set pi = v
                        
                        r = pi.pn
                        r.Offset(0, 1) = pi.duns
                        r.Offset(0, 2) = qi.supp_nm
                        r.Offset(0, 3) = pi.fup_code
                        r.Offset(0, 4) = pi.pus_date
                        r.Offset(0, 5) = pi.eda_date
                        r.Offset(0, 6) = pi.qty
                        r.Offset(0, 7) = pi.nm
                        r.Offset(0, 8) = qi.del_conf
                        Set r = r.Offset(1, 0)
                    Next v
                End If
            End If
        End If
    Next qi

End Sub

Private Sub zrob_labelki_dla_wyrzuconych_pusow(r As Range)

    r.Offset(0, 0) = "PN"
    r.Offset(0, 1) = "DUNS"
    r.Offset(0, 2) = "SUPP_NM"
    r.Offset(0, 3) = "FUP_CODE"
    r.Offset(0, 4) = "PUS Date"
    r.Offset(0, 5) = "EDA Date"
    r.Offset(0, 6) = "Qty"
    r.Offset(0, 7) = "PUS#"
    r.Offset(0, 8) = "DEL_CONF"
End Sub

Private Sub iteracja_przygotuj_tabele_przestawne(ByRef sh_source As Worksheet, ByRef sh_pivot As Worksheet)

    
    tabela_przestawna sh_source, sh_pivot, E_PIVOT_COUNTRY_CODE
    tabela_przestawna sh_source, sh_pivot, E_PIVOT_FUP_CODE
    tabela_przestawna sh_source, sh_pivot, E_PIVOT_PPAP_STATUS
    tabela_przestawna sh_source, sh_pivot, E_PIVOT_RESP
    
    If today_date > CDate(2) Then
        tabela_przestawna sh_source, sh_pivot, E_PIVOT_INTRANSIT_TODAY
    End If
    
    If custom_date > CDate(2) Then
        tabela_przestawna sh_source, sh_pivot, E_PIVOT_INTRANSIT_CUSTOM_DATE
    End If
    
    If mrd_date > CDate(2) Then
        tabela_przestawna sh_source, sh_pivot, E_PIVOT_INTRANSIT_MRD
    End If
    
End Sub


Private Sub tabela_przestawna(ByRef sh_source As Worksheet, ByRef sh_pivot As Worksheet, e As E_PIVOTS)


    Dim ph As PivotHandler
    Set ph = New PivotHandler
    
    
    
    With ph
        .przygotuj_adres sh_source
        .init_ sh_source.Parent, sh_pivot, e, mrd_date, today_date, custom_date
        .no_to_rysuj e
    End With

End Sub

Private Sub iteracja_pierwsza_zrob_flat_table(ByRef w As Workbook, ByRef r As Range)
    Set qi = Nothing
    For Each qi In qi_collection
    
        r = qi.pn
        r.Offset(0, 1) = qi.duns
        r.Offset(0, 2) = qi.resp
        r.Offset(0, 3) = qi.supp_nm
        r.Offset(0, 4) = qi.del_conf
        r.Offset(0, 5) = qi.fup_code
        r.Offset(0, 6) = qi.country_code
        r.Offset(0, 7) = qi.ppap_status
        If Not qi.pus_bucket Is Nothing Then
            With qi.pus_bucket
                r.Offset(0, 8) = .getEarliestSDate
                r.Offset(0, 9) = .getLatestEDA
                r.Offset(0, 10) = .getInTransitMrd
                r.Offset(0, 11) = .getInTransitToday
                r.Offset(0, 12) = .getInTransitCustom
            End With
        End If
        
        Set r = r.Offset(1, 0)
    Next qi
End Sub


Private Sub przygotuj_labelki(ByRef r As Range)
    ' pn    duns    resp    fup_code    country_code    ppap_status getEarliestSDate    getLatestEDA    getInTransit

    r.Offset(-1, 0) = "PN"
    r.Offset(-1, 1) = "DUNS"
    r.Offset(-1, 2) = "RESP"
    r.Offset(-1, 3) = "SUPP_NM"
    r.Offset(-1, 4) = "DEL_CONF"
    r.Offset(-1, 5) = "FUP_CODE"
    r.Offset(-1, 6) = "COUNTRY_CODE"
    r.Offset(-1, 7) = "PPAP_STATUS"
    r.Offset(-1, 8) = "getEarliestSDate"
    r.Offset(-1, 9) = "getLatestEDA"
    r.Offset(-1, 10) = "getInTransitMRD_" & CStr(mrd_date)
    r.Offset(-1, 11) = "getInTransitToday_" & CStr(today_date)
    r.Offset(-1, 12) = "getInTransitCustom_" & CStr(custom_date)
End Sub

