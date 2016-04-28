VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CfgForm 
   Caption         =   "Config"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "CfgForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CfgForm"
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


Private Sub BtnToggleBackupSys_Click()

    If Me.BtnToggleBackupSys.Value = True Then
        ThisWorkbook.Sheets("config").Range("backup_system").Value = 1
        Me.LabelBackup.Caption = "Backup system is activated"
    ElseIf Me.BtnToggleBackupSys.Value = False Then
    
        'a = MsgBox("Wlasnie zamierzasz wylaczyc system zabezpieczenia tresci pliku, napewno tego chcesz?", vbYesNo)
        'If a = vbYes Then
            ThisWorkbook.Sheets("config").Range("backup_system").Value = 0
            Me.LabelBackup.Caption = "Backup system is unactive"
        'Else
        '    Me.BtnToggleBackupSys.Value = True
        '    ' MsgBox "System kontroli tresci pliku pozosal aktywny"
        'end If
    Else
        MsgBox "ten msgbox nie moze nigdy sie pokazac jest to stan niezdefiniowany dla toggle buttona"
    End If
End Sub



Private Sub ToggleButtonProdDesign_Click()

    If Me.ToggleButtonProdDesign.Value = True Then
        ThisWorkbook.Sheets("config").Range("design_mode") = 1
        Me.ToggleButtonProdDesign.Caption = "Design Mode"
        Me.LabelDesignMode.Caption = "Ustawiono widok projektu"
        ' me.LabelDesignMode.Caption = "Faza produkcyjna"
    ElseIf Me.ToggleButtonProdDesign.Value = False Then
        ThisWorkbook.Sheets("config").Range("design_mode") = 0
        Me.ToggleButtonProdDesign.Caption = "Production Mode"
        Me.LabelDesignMode.Caption = "Faza produkcyjna"
    End If
End Sub

Private Sub ToggleButtonProjectType_Click()


    Dim hch As HideColumnsHandler
    Set hch = New HideColumnsHandler

    If Me.ToggleButtonProjectType.Value = True Then
    
        ThisWorkbook.Sheets("config").Range("project_type").Value = 1
        Me.LabelProjectType.Caption = "Ustawiony 2nd project type"
        
        hch.pokaz_kolumny_dla_mrd2
        'MsgBox "Ustawiony 2nd project type"
    ElseIf Me.ToggleButtonProjectType.Value = False Then
    
        ThisWorkbook.Sheets("config").Range("project_type").Value = 0
        Me.LabelProjectType.Caption = "Ustawiono std project type"
        
        hch.showaj_kolumny_dla_mrd2
        'MsgBox "Ustawiono std project type"
    End If
    
End Sub

Private Sub UserForm_Initialize()


    ' backup_system
    If ThisWorkbook.Sheets("config").Range("backup_system").Value = 0 Then
    
        Me.BtnToggleBackupSys.Value = False
    ElseIf ThisWorkbook.Sheets("config").Range("backup_system").Value = 1 Then
    
        Me.BtnToggleBackupSys.Value = True
    End If
    
    
    If ThisWorkbook.Sheets("config").Range("project_type").Value = 1 Then
        Me.ToggleButtonProjectType.Value = True
        Me.LabelProjectType.Caption = "Ustawiony 2nd project type"
    ElseIf ThisWorkbook.Sheets("config").Range("project_type").Value = 0 Then
        Me.ToggleButtonProjectType.Value = False
        Me.LabelProjectType.Caption = "Ustawiono std project type"
    End If
    
    
    If ThisWorkbook.Sheets("config").Range("design_mode").Value = 1 Then
        Me.ToggleButtonProdDesign.Value = True
        Me.ToggleButtonProdDesign.Caption = "Design Mode"
        Me.LabelDesignMode.Caption = "Ustawiono widok projektu"
        ' me.LabelDesignMode.Caption = "Faza produkcyjna"
    ElseIf ThisWorkbook.Sheets("config").Range("design_mode").Value = 0 Then
        Me.ToggleButtonProdDesign.Value = False
        Me.ToggleButtonProdDesign.Caption = "Production Mode"
        ' Me.LabelDesignMode.Caption = "Ustawiono widok projektu"
        Me.LabelDesignMode.Caption = "Faza produkcyjna"
    End If
End Sub
