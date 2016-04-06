VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



' niniejsza klasa zajmuje sie ustawieniem w sposob automatyczny
' dodawaniem do nowych wierszy ustandaryzowanego schematu formul

Private row_handler As Long
' columns with formulas
Private cwf As E_COLUMNS_WITH_FORMULAS
' all columns
' ten enum przechowuje wszystkie
Private ac As E_MASTER_MANDATORY_COLUMNS

Private czy_trzeba_dopisac As Boolean

Private m As Worksheet
Private ostatni_pn As Range
Private r As Range

Private referencja_z_config As Range
Private rzc As Range

Private czy_walidacja_przeszla As Boolean

Private sh As StatusHandler


Private oknoklist(1) As String
Dim openclosedlist(1) As String


Private Sub Class_Initialize()

    oknoklist(0) = "OK"
    oknoklist(1) = "NOK"
    
    openclosedlist(0) = "OPEN"
    openclosedlist(1) = "CLOSED"


    Set m = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    If m.Range("A1").Offset(1, 0) = "" Then
        ' nie ma danych - nic nie robie
        MsgBox "w wierszu drugim (kolumna pierwsza) nie ma danych - logika nie moze zostac wykonana"
    Else
        Set ostatni_pn = m.Range("A1").End(xlDown)
        Set referencja_z_config = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("j1:j9")
        Set rzc = referencja_z_config
        
        row_handler = -1
        czy_trzeba_dopisac = False
        czy_walidacja_przeszla = True
        
        
        ' tutaj oprocz "czy trzeba dopisac" bedzie jeszcze wewnetrzenie potencjalnie zmieniana wartosc
        ' walidacji za pomoca side effectu
        czy_trzeba_dopisac = sprawdz_i_ewentualnie_dopisz()
        
        
        If czy_trzeba_dopisac Then
            MsgBox "brakowalo formul - zostaly dopisane!"
            
        End If
        
        If czy_walidacja_przeszla Then
            MsgBox "cos nie tak bylo z formulami wejsciowymi - zostaly naprawione!"
        End If
    End If
    
    
    
    ' teraz opcja validation
    add_validations
    
    
    

End Sub

Private Sub add_validations()
    
    Set sh = New StatusHandler
    
    
    
    ' teraz niestety wszystkie warunki po kolei
    wierszyk = ostatni_pn.Row
    sh.init_statusbar wierszyk
    Application.ScreenUpdating = False
    sh.show
    
    For x = 2 To wierszyk
        
        Set r = m.Cells(x, WizardMain.Delivery_reconfirmation)
        On Error Resume Next
        r.Validation.Delete
        r.Validation.Add xlValidateList, xlValidAlertStop, , Join(oknoklist, ",")
        
        
        Set r = m.Cells(x, WizardMain.Oncost_confirmation)
        On Error Resume Next
        r.Validation.Delete
        r.Validation.Add xlValidateList, xlValidAlertStop, , Join(openclosedlist, ",")
        
        sh.progress_increase
    Next x
    
    
    sh.hide
    Application.ScreenUpdating = True
End Sub



Private Function sprawdz_i_ewentualnie_dopisz() As Boolean
    sprawdz_i_ewentualnie_dopisz = False
    czy_walidacja_przeszla = True
    
    
    Set sh = New StatusHandler
    
    
    
    ' teraz niestety wszystkie warunki po kolei
    wierszyk = ostatni_pn.Row
    sh.init_statusbar rzc.Count
    Application.ScreenUpdating = False
    sh.show
    
    
    ' losowo jedna formula do weryfikacji
    If wierszyk > m.Cells(1, WizardMain.E_F_MRD1_ORDERED_STATUS).End(xlDown).Row Then
        sprawdz_i_ewentualnie_dopisz = True
    End If
    
    ' jesli pozostaje false i tak zrobimy skromna walidacje i sprawdzimy sobie poprawnosc wpisanych formul
    
    Dim ir As Range
    For Each ir In rzc
        
        ktora_kolumna = ir.Offset(0, 1)
        
        Set r = m.Cells(2, ktora_kolumna)
        Application.EnableEvents = False
        czy_walidacja_przeszla = petla_po_tej_kolumnie_i_uzupelnij_braki(r, ir)
        Application.EnableEvents = True
        sh.progress_increase
    Next ir
    sh.hide
    Application.ScreenUpdating = True
End Function


' ten argument ref nie jest potrzebny jednak zeby nie zrobic
' znow jakies mgly na kodzie oficjalnie wpisuje r i ir jako argumenty
Private Function petla_po_tej_kolumnie_i_uzupelnij_braki(ByRef r As Range, ByRef ir As Range) As Boolean

    petla_po_tej_kolumnie_i_uzupelnij_braki = True
    
    
    If r.Formula = ir.Formula Then
        ' wszsytko jest ok wystarczy przecieagnac formule z wiersza drugiego i ciachnac
        ' ja do reszty
        petla_po_tej_kolumnie_i_uzupelnij_braki = True
    Else
        ' powazna sprawa - juz na poczatku cos jest nie tak z formula
        petla_po_tej_kolumnie_i_uzupelnij_braki = False
        r.Formula = ir.Formula
    End If

    Dim tmp As Range
    'Set tmp = r.Offset(1, 0)
    'Do
    '    ' nie wazne co z walidacja  i tak dla spokoju skopiuje wszystko raz jeszcze
    '    r.Copy tmp
    '    Set tmp = tmp.Offset(1, 0)
    'Loop Until tmp.Row > ostatni_pn.Row
    
    Set tmp = r.Parent.Range(r.Offset(1, 0), r.Offset(ostatni_pn.Row - r.Row, 0))
    ' MsgBox tmp.Count
    ' AutoFill Destination:=Range("P2:P19")
    ' r.Copy tmp
    'r.AutoFill Destination:=Range("P2:P19"), Type:=xlFillDefault
    
    'Selection.Copy
    'Range("P3").Select
    'Application.CutCopyMode = False
    'Selection.Copy
    'Range("P8:P15").Select
    'ActiveSheet.Paste
    'Range("P9").Select
    'Application.CutCopyMode = False
    
    r.Parent.Activate
    r.Select
    Selection.Copy
    tmp.Select
    r.Parent.Paste
    ' Application.CutCopyMode = False
    
    
End Function





