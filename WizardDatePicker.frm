VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WizardDatePicker 
   Caption         =   "Wizard"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "WizardDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WizardDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Private details_handler As DetailsHandler

Public Sub connect_with_details_handler(arg As DetailsHandler)
    Set details_handler = arg
End Sub

Private Sub BtnNext_Click()


    ' jedyny wyjatek na dwoch polach
    ' nie daje cw tylko normlanie date
    If Me.CheckBoxDateAvail.Value = True Then
        Me.TextBoxBufor.Value = TBD
    ElseIf Me.CheckBoxDateAvail.Value = False Then
        
        If (details_handler.get_e = PICKUP_DATE) Or _
            (details_handler.get_e = PPAP_GATE) Or _
            (details_handler.get_e = E_MRD_REG_ROUTES) Or _
            (details_handler.get_e = E_MRD_DATE) Then
            
                Me.TextBoxBufor.Value = CStr(DTPickerInput.Value)
        Else
                Me.TextBoxBufor.Value = Me.LabelCW.Caption
        End If
        
    End If
    
    details_handler.dalej Me
    Set details_handler = Nothing
End Sub

Private Sub BtnPrev_Click()

    
    If Me.CheckBoxDateAvail.Value = True Then
        Me.TextBoxBufor.Value = TBD
    ElseIf Me.CheckBoxDateAvail.Value = False Then
        If (details_handler.get_e = PICKUP_DATE) Or (details_handler.get_e = PPAP_GATE) Then
            Me.TextBoxBufor.Value = CStr(DTPickerInput.Value)
        Else
            Me.TextBoxBufor.Value = Me.LabelCW.Caption
        End If
    End If
    
    details_handler.cofnij Me
    Set details_handler = Nothing
End Sub


Private Sub CheckBoxDateAvail_Change()
    
    If Me.CheckBoxDateAvail.Value = True Then
        Me.DTPickerInput.Enabled = False
        'Me.DTPickerInput.Visible = False
        'Me.LabelCW.Visible = False
    ElseIf Me.CheckBoxDateAvail.Value = False Then
        Me.DTPickerInput.Enabled = True
        'Me.DTPickerInput.Visible = True
        'Me.LabelCW.Visible = True
    End If
End Sub

Private Sub DTPickerInput_Change()
    redefine_cw
End Sub

Private Sub UserForm_Initialize()
    DTPickerInput.Value = Now
    Me.DTPickerInput.SetFocus
    redefine_cw
    Me.CheckBoxDateAvail.Value = False
End Sub


Private Sub redefine_cw()
    Dim d As Date
    d = CDate(DTPickerInput.Value)
    If Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 2 Then
        LabelCW.Caption = "Y" & CStr(Year(d)) & "CW" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
    ElseIf Len(CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))) = 1 Then
        LabelCW.Caption = "Y" & CStr(Year(d)) & "CW0" & CStr(Application.WorksheetFunction.IsoWeekNum(CDate(d)))
    Else
        MsgBox "redefine_cw - nie powinno sie pokazac"
    End If
End Sub
