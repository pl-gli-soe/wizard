VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WizardToggle 
   Caption         =   "Wizard"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   OleObjectBlob   =   "WizardToggle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WizardToggle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public e As E_NEW_PROJECT_ITEM
Private details_handler As DetailsHandler

Public Sub connect_with_details_handler(arg As DetailsHandler)
    Set details_handler = arg
End Sub

Private Sub BtnNext_Click()
    Me.TextBoxBufor.Value = Me.ToggleButtonInput.Caption
    details_handler.dalej Me
    Set details_handler = Nothing
End Sub

Private Sub BtnPrev_Click()
    Me.TextBoxBufor.Value = Me.ToggleButtonInput.Caption
    details_handler.cofnij Me
    Set details_handler = Nothing
End Sub

Private Sub ToggleButtonInput_Click()


    If CStr(Me.LabelDesc.Caption) = "" Then
        Me.LabelDesc.Caption = CStr(BIW_GA)
    End If
    
    e = CLng(Me.LabelDesc.Caption)
    
    If e = BIW_GA Then
        If Me.ToggleButtonInput.Value = True Then
            Me.ToggleButtonInput.Caption = "GA"
        
        ElseIf Me.ToggleButtonInput.Value = False Then
            Me.ToggleButtonInput.Caption = "BIW"
        Else
            MsgBox "ten msgbox nigdy nie powinien sie pokazac - toggle button click ga biw"
        
        End If
    ElseIf e = E_ACTIVE Then
        If Me.ToggleButtonInput.Value = True Then
            Me.ToggleButtonInput.Caption = "YES"
        
        ElseIf Me.ToggleButtonInput.Value = False Then
            Me.ToggleButtonInput.Caption = "NO"
        Else
            MsgBox "ten msgbox nigdy nie powinien sie pokazac - toggle button click active"
        
        End If
    Else
            MsgBox "ten msgbox nigdy nie powinien sie pokazac - toggle button click ogolnie"
    End If
        
End Sub

