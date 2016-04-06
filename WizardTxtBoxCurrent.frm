VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WizardTxtBoxCurrent 
   Caption         =   "Wizard"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   OleObjectBlob   =   "WizardTxtBoxCurrent.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WizardTxtBoxCurrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private details_handler As DetailsHandler

Public Sub connect_with_details_handler(arg As DetailsHandler)
    Set details_handler = arg
End Sub

Private Sub BtnNext_Click()
    Me.TextBoxBufor.Value = Me.TextBoxInput.Value
    details_handler.dalej Me
    Set details_handler = Nothing
End Sub

Private Sub BtnPrev_Click()
    Me.TextBoxBufor.Value = Me.TextBoxInput.Value
    details_handler.cofnij Me
    Set details_handler = Nothing
End Sub

Private Sub UserForm_Initialize()
    With Me.TextBoxInput
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        
End Sub
