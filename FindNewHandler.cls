VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FindNewHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private this_w As Workbook
Private source_w As Workbook

Private this_sh As Worksheet
Private source_sh As Worksheet

Private tr As Range
Private sr As Range


Public Sub create_connections(sw As String, ss As String)

    Set this_w = ThisWorkbook
    Set this_sh = this_w.Sheets(WizardMain.MASTER_SHEET_NAME)
    
    Set source_w = Workbooks(sw)
    Set source_sh = source_w.Sheets(ss)
End Sub


Public Sub pomaluj_te_ktorych_nie_ma_w_aktualnym_td(ktora_kolumna_to_pn As Integer)
    
    Set tr = this_sh.Range("A2")
    Set sr = source_sh.Range("A" & CStr(ktora_kolumna_to_pn))
    Set sr = source_sh.Range(sr, sr.End(xlDown))
    
    
    Dim wynik As Range
    firstAdr = ""
    Do
    
        With sr
            Set wynik = .Find(CStr(Trim(tr)))
            
            If Not wynik Is Nothing Then
                firstAdr = wynik.Address
                Do
                    wynik.Interior.Color = RGB(0, 190, 230)
                    Set wynik = .FindNext(wynik)
                Loop While Not wynik Is Nothing And wynik.Address <> firstAdr
            End If
        End With
        Set tr = tr.Offset(1, 0)
    Loop Until Trim(tr) = ""
    
End Sub


