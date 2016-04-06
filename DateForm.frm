VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateForm 
   Caption         =   "Date"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1905
   OleObjectBlob   =   "DateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnSubmit_Click()
    hide
    Selection.Value = Me.DTPicker1.Value
End Sub

Private Sub DTPicker1_AfterUpdate()

    Selection.Value = Me.DTPicker1.Value
End Sub


Private Sub DTPicker1_DblClick()
    hide
    Selection.Value = Me.DTPicker1.Value
End Sub
