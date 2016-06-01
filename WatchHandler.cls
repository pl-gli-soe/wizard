VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WatchHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ww As String
Private rng_line As Range



Public Sub createWatchForThisLine(r As Range)
    
    ' Application.Watches.Add ThisWorkbook.Sheets("MASTER").Range("A3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
        Range("B3")
    ''Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("C3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("D3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("E3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("F3")
    'a 'pplication.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
     '   Range("G3")
    'a 'pplication.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
     '   Range("H3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("I3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("J3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("K3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("L3")
    'Application.Watches.Add Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER"). _
    '    Range("M3")
    'Application.Watches(Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER").Range( _
    '    "F3")).Delete
    'Application.Watches(Workbooks("M_wizard_3.98 dev.xlsm").Sheets("MASTER").Range( _
    '    "A3")).Delete
    'Application.CommandBars("Watch Window").Visible = False
    'Application.CommandBars("Watch Window").Visible = False
    
    If ThisWorkbook.FullName = ActiveWorkbook.FullName Then
    
        If r.Parent.Name = WizardMain.MASTER_SHEET_NAME Then
        
                deleteWatchForThisLine r
                
                ' okno watch nie jest aktywne
                ' ==================================
                Application.CommandBars(CStr(ww)).Visible = True
                
                'Set zasieg =
                '    r.Parent.Range( _
                '        r.Parent.Cells(r.Row, WizardMain.pn), _
                '        r.Parent.Cells(r.Row, WizardMain.SQ_Comments))
                    
                
                
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("A" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("B" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("C" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("D" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("E" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("F" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("G" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("H" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("I" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("J" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("K" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("L" & CStr(r.Row))
                Application.Watches.Add ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("M" & CStr(r.Row))
                
                
                Application.CommandBars(CStr(ww)).Visible = True
                ' ==================================
        End If
    End If
End Sub

Public Sub deleteWatchForThisLine(r As Range)

    Application.CommandBars(CStr(ww)).Visible = False


    If Application.Watches.Count > 0 Then
                Application.Watches.Delete

    End If


End Sub

Private Sub Class_Initialize()
    ww = "Watch Window"
End Sub

