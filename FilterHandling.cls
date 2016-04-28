VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

Private fup_code As String

Private fup_code_ok As Boolean

Dim r As Range, ir As Range

Public fma_resp_cfg As Range
Public my_fup_code_cfg As Range
Public selection_cfg As Range


Public Sub ustaw_swiezy_filtr()

    If ActiveSheet.AutoFilterMode = True Then
    Else
        r.AutoFilter
    End If
    On Error Resume Next
    r.Parent.ShowAllData
End Sub


Public Function filtruj_po_selekcji() As Boolean

    If Selection.Count = 1 Then
        filtruj_po_selekcji = filtruj_po_aktywnej_komorce()
    Else
        filtruj_po_selekcji = filtruj_po_aktywnych_komorkach()
    End If

End Function


Private Function filtruj_po_aktywnych_komorkach() As Boolean
    filtruj_po_aktywnych_komorkach = False
    
    Dim tmp As Range
    Dim v(1 To SELECTION_LIMIT) As String
    Set tmp = Selection
        
    If Me.selection_cfg.Value = 1 Then
    
        If tmp.Columns.Count = 1 Then
        
            Dim ir As Range
            i = 1
            For Each ir In tmp
                If Not ir.EntireRow.Hidden Then
                    v(i) = ir
                    i = i + 1
                End If
            Next ir
            
        
            ' v = Application.Transpose(Selection)
            Selection.AutoFilter Field:=tmp(1, 1).Column, Criteria1:=v, Operator:=xlFilterValues
        Else
            MsgBox "nie mozesz nalozyc dwoch filtrow na raz!"
        End If
    ElseIf Me.selection_cfg.Value = 0 Then
        Selection.AutoFilter Field:=tmp(1, 1).Column
    End If
End Function

Private Function filtruj_po_aktywnej_komorce() As Boolean
    filtruj_po_aktywnej_komorce = False
    
    If Me.selection_cfg.Value = 1 Then
        r.AutoFilter ActiveCell.Column, ActiveCell.Value
    ElseIf Me.selection_cfg.Value = 0 Then
        r.AutoFilter Field:=ActiveCell.Column
    End If
    
    
        
End Function


Public Function filtruj_po_resp() As Boolean
    filtruj_po_resp = False
    
    ' fma_resp_fliter
    If Me.fma_resp_cfg.Value = 1 Then
        r.AutoFilter WizardMain.Responsibility, FMA_WITH_STARS
    Else
        r.AutoFilter WizardMain.Responsibility
    End If
        
End Function

Public Function filtruj_po_deck() As Boolean

    filtruj_po_deck = False
    
   
    
    If Me.my_fup_code_cfg.Value = 1 Then
         ' my_fup_code_filter
        okresl_deck
        r.AutoFilter WizardMain.fup_code, fup_code
    Else
        r.AutoFilter WizardMain.fup_code
    End If
End Function



Public Sub okresl_deck()


    If fup_code = "" Then
        If Len(Selection.Value) = 2 Then
            If UCase(Left(Selection.Value, 1)) = "W" Then
                fup_code = Selection.Value
            End If
        End If
        
        If fup_code = "" Then
        
            fup_code_ok = False
        
            wpisz_fup_code 1
            
            
        End If
    End If
End Sub



Private Sub Class_Initialize()

    fup_code = ""
    fup_code_ok = False

    Set r = ThisWorkbook.Sheets(MASTER_SHEET_NAME).Range("A1")
    
    Set fma_resp_cfg = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("fma_resp_fliter")
    Set my_fup_code_cfg = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("my_fup_code_filter")
    Set selection_cfg = ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("selection_filter")
    
    
End Sub

Private Sub wpisz_fup_code(ile_razy As Integer)


    If ile_razy < 3 Then

        fup_code = InputBox("Wpisz swoj FUP code: ")
        
        If Len(fup_code) = 2 Then
            If UCase(Left(fup_code, 1)) = "W" Then
                fup_code_ok = True
            End If
        End If
        
        If fup_code_ok = False Then
            MsgBox "niepoprawny fup code!"
            wpisz_fup_code ile_razy + 1
        End If
    Else
        MsgBox "Baranku! 3 razy sie duples podczas proby wpisania FUP coda, nie mam slow!"
        End
    End If
End Sub
