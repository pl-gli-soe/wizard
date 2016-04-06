VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WybierzPlikForm 
   Caption         =   "Wybierz Plik: "
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3045
   OleObjectBlob   =   "WybierzPlikForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WybierzPlikForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    hide
    If Me.ListBox1.ListCount > 0 Then
        PodmienForm.nazwa_aktywnego_pliku = Me.ListBox1.Value
        
    Else
        MsgBox "nie ma czego wybrac!"
    End If

End Sub
