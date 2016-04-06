VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FUPForm 
   Caption         =   "FUP"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3255
   OleObjectBlob   =   "FUPForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FUPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnSubmit_Click()
    hide
    
    Dim gc As GameChanger
    Set gc = New GameChanger
    
    gc.change_deck_on_selection CStr(Me.TextBoxFUP)
    
    Set gc = Nothing
End Sub
