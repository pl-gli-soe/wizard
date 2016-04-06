VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DodajPNForm 
   Caption         =   "Dodaj"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3210
   OleObjectBlob   =   "DodajPNForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DodajPNForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBoxINDX_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    hide
    Dim psh As Worksheet, r As Range
    
    Set psh = ThisWorkbook.Sheets(PICKUPS_SHEET_NAME)
    Set r = psh.Range("a2")
    
    linia_txtu = ""
    For x = 0 To Me.ListBoxIndx.ListCount - 1
        If Me.ListBoxIndx.Selected(x) = True Then
            linia_txtu = Me.ListBoxIndx.List(x)
            Exit For
        End If
    Next x
    
    arr = Split(linia_txtu, ",")
    str_pn = arr(LBound(arr))
    str_duns = arr(LBound(arr) + 1)
    str_deck = arr(LBound(arr) + 2)
    
    ' sprawdz, czy dany part number zostal juz wczesniej dodany dla wybranego PUSa
    ' Set r = psh.Range("a2") ' <- to jest u gory jeszcze nie zmod zatem to jest w rem
    Do
    
        If r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX) = Me.TextBoxPUSName Then
            If CStr(r.Offset(0, WizardMain.O_PN - WizardMain.O_INDX)) = str_pn Then
                MsgBox "taki PN zostal juz wybrany dla tego PUSa - zadna akcja dodawania nie zostanie podjeta"
                Exit Sub
            End If
        End If
        
        WizardMain.nowy_schemat_offsetu_w_arkuszu_pickups r
    Loop Until r.Row > WizardMain.POLOWA_CAPACITY_ARKUSZA
    
    
    
    If USERS_LIMIT < UBound(ThisWorkbook.UserStatus, 1) Then
        ' users_status_usun_moje_stare_instancje CStr(Application.UserName)
        MsgBox "przekroczono limit uzytkownikow pliku - sprawdz liste w Review -> Share Workbook"
        End
    End If
    
    ' gdzie zaczynamy
    ' G_STEP_BETWEEN_PARALELL_USERS
    gdzie_zaczynamy = 1
    Users = ThisWorkbook.UserStatus
    For x = 1 To UBound(ThisWorkbook.UserStatus, 1)
        If CStr(Application.UserName) = CStr(Users(x, 1)) Then
            gdzie_zaczynamy = (G_STEP_BETWEEN_PARALELL_USERS * (x - 1)) + 1
            Exit For
        End If
    Next x
    
    Set r = psh.Range("a1").Offset(gdzie_zaczynamy, 0)
    
    If r = "" Then
        
    ElseIf r.Offset(1, 0) = "" Then
    
        Set r = r.Offset(1, 0)
    Else
        Set r = r.End(xlDown).Offset(1, 0)
    End If
    
    r.Value = linia_txtu
    ' ThisWorkbook.Save
    r.Offset(0, WizardMain.O_PN - WizardMain.O_INDX).Value = str_pn
    r.Offset(0, WizardMain.O_DUNS - WizardMain.O_INDX).Value = str_duns
    r.Offset(0, WizardMain.O_FUP_code - WizardMain.O_INDX).Value = str_deck
    
    r.Offset(0, WizardMain.O_Pick_up_date - WizardMain.O_INDX).Value = Me.DTPickerPUSDate
    r.Offset(0, WizardMain.O_Delivery_Date - WizardMain.O_INDX).Value = Me.DTPickerDelDate
    
    r.Offset(0, WizardMain.O_Pick_up_Qty - WizardMain.O_INDX).Value = 0
    
    r.Offset(0, WizardMain.O_PUS_Number - WizardMain.O_INDX).Value = Me.TextBoxPUSName
    
    ' ThisWorkbook.Save
    edit_pickup
    
End Sub

Private Sub TextBoxPtrn_Change()



    Dim msh As Worksheet
    Dim r As Range

    Set msh = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    Set r = msh.Cells(2, WizardMain.pn)
    
    Me.ListBoxIndx.Clear
    
    Do
        'If Me.ListBoxIndx.ListCount > 0 Then
            If CStr(Me.TextBoxBufferForIndx0.Value) Like _
                "*" & CStr(msh.Cells(r.Row, WizardMain.duns)) & "," & CStr(msh.Cells(r.Row, WizardMain.fup_code)) Then
                
                    tmp_txt = CStr(msh.Cells(r.Row, WizardMain.pn)) & _
                        "," & CStr(msh.Cells(r.Row, WizardMain.duns)) & _
                        "," & CStr(msh.Cells(r.Row, WizardMain.fup_code))
                        
                    If (tmp_txt Like "*" & CStr(Me.TextBoxPtrn) & "*") Or CStr(Me.TextBoxPtrn) = "" Then
                
                        Me.ListBoxIndx.AddItem tmp_txt
                    End If
            End If
        ' End If
        Set r = r.Offset(1, 0)
    Loop Until r = ""
End Sub
