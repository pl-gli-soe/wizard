Attribute VB_Name = "FilterModule"
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
