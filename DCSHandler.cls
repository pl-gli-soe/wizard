VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DCSHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public r As Range

Private dsc_sh As Worksheet
Private details_sh As Worksheet

' Private current_form As Variant

Private dscr As Range


Private Sub Class_Initialize()
    ' przygotowanie listy pod dcsa
    ' Set current_form = DCS_DropdownList
    ' current_form.Show vbModeless
    
    Set dsc_sh = ThisWorkbook.Sheets(DCS_SHEET_NAME)
    Set details_sh = ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
    Set dscr = dsc_sh.Range("B2")
    
    If CStr(dscr) <> "" Then
        przygotuj_listboxa1
    End If
    
    Set DCS_DropdownList.dcs_handler = dcs_handler
End Sub

Private Sub przygotuj_listboxa1()
    ' DCS_DropdownList
    With DCS_DropdownList
        .ListBox1.Clear
        
        Do
            .ListBox1.AddItem CStr(dscr.Offset(0, -1))
            Set dscr = dscr.Offset(1, 0)
        Loop Until Trim(dscr) = ""
    End With
    
End Sub

Public Sub work_around_listbox1_and_listbox2()

    
    o_ile_przesuwamy_MRD = 0
    

    Set dscr = ThisWorkbook.Sheets(DCS_SHEET_NAME).Range("B2")
    Do
        With DCS_DropdownList
        
        
        
            If .ListBox1.Value = CStr(dscr.Offset(0, -1)) Then
            
                ' ============================================
                
                
                If dscr = 1 Then
                
                
                    If .TextBox1.Value <> "" Then
                        If IsNumeric(.TextBox1.Value) Then
                        
                            If CLng(.TextBox1.Value) > CLng(dscr.Offset(0, 1)) Then
                                MsgBox "Twoj offset przekracza limit okreslony na: " & CStr(dscr.Offset(0, 1)) & Chr(10) & _
                                    "Przesuniecie zostalo zdefiniowane na 0!"
                                .TextBox1.Value = 0
                                o_ile_przesuwamy_MRD = 0
                            Else
                                o_ile_przesuwamy_MRD = CLng(.TextBox1.Value)
                            End If
                        Else
                            o_ile_przesuwamy_MRD = 0
                        End If
                    ElseIf .TextBox1.Value = "" Then
                        o_ile_przesuwamy_MRD = 0
                    End If
                
                
                    .ListBox2.Clear
                    txt_pod_move = CStr(przelicz_ten_y_cw(CStr(details_sh.Cells(WizardMain.mrd, 2).Value), CLng(o_ile_przesuwamy_MRD)))
                    txt_pod_move = Replace(CStr(.ListBox1.Value), MRD_KLUCZ_DO_PODMIANY, txt_pod_move)
                    
                    .ListBox2.AddItem txt_pod_move
                    
                Else
                    .ListBox2.Clear
                    .ListBox2.AddItem CStr(.ListBox1.Value)
                    
                End If
                
                ' ============================================
            End If
        End With
            
            Set dscr = dscr.Offset(1, 0)
    Loop Until Trim(dscr) = ""
    
End Sub

Private Function przelicz_ten_y_cw(txt As String, ile As Long)

    Dim d As Date

    If ile = 0 Then
        przelicz_ten_y_cw = txt
    Else
        tmp = Split(txt, "CW")
        
        rok = Replace(tmp(LBound(tmp)), "Y", "")
        cw = tmp(UBound(tmp))
        
        
        d = fnDateFromWeek(Int(rok), Int(cw), 2)
        d = d + (ile * 7)
         
        If Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 2 Then
            przelicz_ten_y_cw = "Y" & CStr(Year(d)) & "CW" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
        ElseIf Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 1 Then
            przelicz_ten_y_cw = "Y" & CStr(Year(d)) & "CW0" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
        Else
            MsgBox "przelicz_ten_y_cw - nie powinien sie pojawic!"
        End If
    End If

End Function

Private Sub Class_Terminate()
    ' Set current_form = Nothing
End Sub
