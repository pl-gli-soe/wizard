VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HideColumnsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m As Worksheet

Private Sub Class_Initialize()
    Set m = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
End Sub

Private Sub inner_show_poka(b As Boolean)

    Dim r As Range
    
    Set r = m.Cells(1, WizardMain.MRD2_QTY)
    r.EntireColumn.Hidden = b
    
    Set r = m.Range(m.Cells(1, WizardMain.MRD2_Ordered_date), m.Cells(1, WizardMain.MRD2_Total_PUS_STATUS))
    r.EntireColumn.Hidden = b
    
End Sub

Public Sub pokaz_kolumny_dla_mrd2()
    inner_show_poka False
End Sub

Public Sub showaj_kolumny_dla_mrd2()
    inner_show_poka True
End Sub

