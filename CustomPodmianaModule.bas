Attribute VB_Name = "CustomPodmianaModule"
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

' baw sie
' ------------------------------------------------------------------------------
' target = range
' target.parent = sheet
' target.parent.parent = workbook


Public Sub work_around_custom_copy_worksheet()
    
    ' sprawdz ostatni nie pusty wiersz dla zewnetrznego pliku
    Dim external_sh As Worksheet, ostatni As Range, r As Range, mr As Range
    Dim ttmp As Range, coll As Collection
    
    Set external_sh = _
        CustomMasterCpForm.podmien_handler.source_workbook.Sheets(WizardMain.RAW_TXT)
    Set ostatni = external_sh.Range("A1")
    ' MsgBox external_sh.Name
    
    Do
        ' troche szybciej (jak przekopiujemy kilka pustych pol to nic zlego sie nie stanie)
        Set ostatni = ostatni.Offset(100, 0)
    Loop Until Trim(ostatni) = ""
    
    
    ' no to ile mamy do przekopiowania?
    
    Set coll = New Collection
    Set ttmp = external_sh.Range("A1")
    i = 0
    Do
    
    
        'If CStr(Trim(CustomMasterCpForm.ListBoxRawData.Value)) = ttmp Then
        '    Set r = ttmp
        '    Exit Do
        'End If
        
        If CustomMasterCpForm.ListBoxRawData.Selected(i) Then
            Set r = ttmp
            
            coll.Add r
        End If
        
        
        i = i + 1
        Set ttmp = ttmp.Offset(0, 1)
    Loop Until CStr(Trim(ttmp)) = ""
    
    
    If coll.Count > 0 Then
    
        With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
        
            num = .Range("d2")
            
            
            Select Case num
                Case 1
                
                    If coll.Count = 1 Then
                
                        CustomMasterCpForm.Label2.Caption = r.Parent.Parent.Name & "," & r.Parent.Name
                        .Range("d3") = r.Parent.Parent.Name & "," & r.Parent.Name
                    Else
                        MsgBox "powazny blad w listbox wyboru pliku dla niestandardowej podmiany danych!"
                        MsgBox "mateo sprawdz custom podmiana module"
                        End
                    End If
                Case WizardMain.G_WYBIERZ_KOLUMNE_PN To WizardMain.G_WYBIERZ_KOLUMNE_CMNTS
                    ' MsgBox "meh"
                    If coll.Count = 1 Then
                        special_kopiowanie_ r.Parent, ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME), ostatni.Row, r.Column, num - 1
                    Else
                        MsgBox "powazny blad w listbox wyboru pliku dla niestandardowej podmiany danych!"
                        MsgBox "mateo sprawdz custom podmiana module"
                        End
                    End If
                Case WizardMain.G_WYBIERZ_KOLUMNE_CUSTOMOWA
                    
                    ' tutaj customowe kolumny
                    For Each r In coll
                        Set mr = ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("A1").End(xlToRight).Offset(0, 1)
                        special_kopiowanie_ r.Parent, ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME), ostatni.Row, r.Column, mr.Column, True
                    Next r
                Case Else
                    MsgBox "cos poszlo nie tak"
            End Select
        End With
    End If
End Sub


' ten sub powinien byc w podmien handler chyba
Public Sub work_around_custom_copy_worksheet_osbolete()
    
    ' sprawdz ostatni nie pusty wiersz dla zewnetrznego pliku
    Dim external_sh As Worksheet, ostatni As Range, r As Range, mr As Range
    
    
    Set external_sh = _
        CustomMasterCpForm.podmien_handler.source_workbook.Sheets(WizardMain.RAW_TXT)
    Set ostatni = external_sh.Range("A1")
    ' MsgBox external_sh.Name
    
    Do
        ' troche szybciej (jak przekopiujemy kilka pustych pol to nic zlego sie nie stanie)
        Set ostatni = ostatni.Offset(100, 0)
    Loop Until Trim(ostatni) = ""
    
    
    
    
    Dim ttmp As Range
    Set ttmp = external_sh.Range("A1")
    i = 1
    Do
    
    
    
        On Error Resume Next
        If CStr(Trim(CustomMasterCpForm.ListBoxRawData.Value)) = ttmp Then
            Set r = ttmp
            Exit Do
        End If
        Set ttmp = ttmp.Offset(0, 1)
    Loop Until CStr(Trim(ttmp)) = ""
    
    
    
    

    With ThisWorkbook.Sheets(WizardMain.CUSTOM_COPY_SHEET_NAME)
    
        num = .Range("d2")
        
        
        Select Case num
            Case 1
                CustomMasterCpForm.Label2.Caption = r.Parent.Parent.Name & "," & r.Parent.Name
                .Range("d3") = r.Parent.Parent.Name & "," & r.Parent.Name
            Case WizardMain.G_WYBIERZ_KOLUMNE_PN To WizardMain.G_WYBIERZ_KOLUMNE_CMNTS
                ' MsgBox "meh"

                special_kopiowanie_ r.Parent, ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME), ostatni.Row, r.Column, num - 1
            Case WizardMain.G_WYBIERZ_KOLUMNE_CUSTOMOWA
                
                ' tutaj customowe kolumny
                Set mr = ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME).Range("A1").End(xlToRight).Offset(0, 1)
                special_kopiowanie_ r.Parent, ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME), ostatni.Row, r.Column, mr.Column, True
            Case Else
                MsgBox "cos poszlo nie tak"
        End Select
    End With
End Sub


Private Sub special_kopiowanie_(ByRef out_msh As Worksheet, ByRef this_msh As Worksheet, w As Long, sx As Integer, dx As Integer, Optional labelka As Boolean)


    ' sx - source x
    ' dx - destination x
    If labelka Then
        out_msh.Range(out_msh.Cells(1, sx), out_msh.Cells(w, sx)).Copy this_msh.Range(this_msh.Cells(1, dx), this_msh.Cells(w, dx))
    Else
        out_msh.Range(out_msh.Cells(2, sx), out_msh.Cells(w, sx)).Copy this_msh.Range(this_msh.Cells(2, dx), this_msh.Cells(w, dx))
    End If

End Sub
