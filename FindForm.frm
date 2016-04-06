VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindForm 
   Caption         =   "Znajdz"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   OleObjectBlob   =   "FindForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ref_find_new_handler As FindNewHandler

Private Sub BtnSubmit_Click()
    
    ' Dim w As String, s As String
    
    w = ""
    s = ""
    
    On Error Resume Next
    w = Me.ListBoxW.Value
    On Error Resume Next
    s = Me.ListBoxS.Value
    
    
    If s <> "" And w <> "" Then
    
        hide
        With Me.ref_find_new_handler
            ' params
            On Error Resume Next
            .create_connections CStr(w), CStr(s)
            
            If IsNumeric(Me.TextBoxPNCol.Value) And Int(Me.TextBoxPNCol.Value) > 0 Then
                .pomaluj_te_ktorych_nie_ma_w_aktualnym_td Int(Me.TextBoxPNCol.Value)
            Else
                MsgBox "zle dobrana wartosc kolumny - zostala wybrana automatycznie kolumna pierwsza jako PN!"
                .pomaluj_te_ktorych_nie_ma_w_aktualnym_td Int(1)
            End If
        End With
    Else
        MsgBox "poszlo cos nie tak z proba"
    End If
End Sub

Private Sub ListBoxW_Click()
    
    inner_refresh_sheet_listbox
End Sub

Private Sub inner_refresh_sheet_listbox()
    
    Me.ListBoxS.Clear
    
    w_txt = ""

    On Error Resume Next
    w_txt = CStr(Me.ListBoxW.Value)
    
    
    If w_txt <> "" Then
        
        Dim sh As Worksheet
        
        For Each sh In Workbooks(CStr(w_txt)).Sheets
            Me.ListBoxS.AddItem sh.Name
        Next sh
        
    End If
End Sub
