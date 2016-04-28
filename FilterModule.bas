Attribute VB_Name = "FilterModule"
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

Public Sub quick_resp_filter_with_deck()
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    fh.ustaw_swiezy_filtr
    fh.filtruj_po_resp
    fh.filtruj_po_deck
    Set fh = Nothing
End Sub


Public Sub quick_selection_filter(ictrl As IRibbonControl)
    quick_selection_filter_inner
End Sub


Private Sub quick_selection_filter_inner()

    ' inner_quick_clear

    Dim fh As FilterHandling
    Set fh = New FilterHandling
    
    With fh
    
        If .selection_cfg.Value = 0 Then
            .selection_cfg.Value = 1
            fh.filtruj_po_selekcji
        ElseIf .selection_cfg.Value = 1 Then
            .selection_cfg.Value = 0
            fh.filtruj_po_selekcji
        End If
    
    End With

    Set fh = Nothing
End Sub

Public Sub quick_resp_filter(ictrl As IRibbonControl)

    ' inner_quick_clear

    Dim fh As FilterHandling
    Set fh = New FilterHandling
    
    ' fma_resp_fliter
    ' fh.ustaw_swiezy_filtr
    ' fh.filtruj_po_resp
    ' fh.filtruj_po_deck
    
    With fh
    
        If .fma_resp_cfg.Value = 0 Then
            .fma_resp_cfg.Value = 1
            fh.filtruj_po_resp
        ElseIf .fma_resp_cfg.Value = 1 Then
            .fma_resp_cfg.Value = 0
            fh.filtruj_po_resp
        End If
    
    End With
    
    Set fh = Nothing
End Sub

Public Sub quick_deck_filter(ictrl As IRibbonControl)

    'inner_quick_clear

    Dim fh As FilterHandling
    Set fh = New FilterHandling
    ' my_fup_code_filter
    
    
    'fh.ustaw_swiezy_filtr
    ' fh.filtruj_po_resp
    'fh.filtruj_po_deck
    
    With fh
    
        If .my_fup_code_cfg.Value = 0 Then
            .my_fup_code_cfg.Value = 1
            fh.filtruj_po_deck
        ElseIf .my_fup_code_cfg.Value = 1 Then
            .my_fup_code_cfg.Value = 0
            fh.filtruj_po_deck
        End If
    
    End With
    
    
    Set fh = Nothing
End Sub


Public Sub quick_clear(ictrl As IRibbonControl)
    inner_quick_clear
End Sub

Private Sub inner_quick_clear()
    Dim fh As FilterHandling
    Set fh = New FilterHandling
    
    With fh
        .my_fup_code_cfg.Value = 0
        .fma_resp_cfg.Value = 0
        .selection_cfg.Value = 0
        .ustaw_swiezy_filtr
    End With
    
    Set fh = Nothing
End Sub
