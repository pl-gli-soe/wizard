Attribute VB_Name = "CmntSourceModule"
Public Sub comment_synchro(ictrl As IRibbonControl)
    If ActiveSheet.Name = WizardMain.MASTER_SHEET_NAME Then
        zsynchronizuj_komentarz_dla_tej_labelki Selection, ThisWorkbook.Sheets(WizardMain.COMMENT_SOURCE_SHEET_NAME)
    Else
        MsgBox "niewlasciwy arkusz jest aktywny - tylko arkusz master moze korzystac z tej fukcjonalnosci"
    End If
End Sub


Public Sub zsynchronizuj_komentarz_dla_tej_labelki(ByRef r As Range, ByRef sh As Worksheet)


    Dim ch As CommentHandler
    

    Dim f As Range, ir As Range
    
    
    For Each ir In r
    
        If ir.Row = 1 Then
        
        
            Set f = sh.Range("a1")
        
            Do
                If f.Value = ir.Value Then
                
                    ir.ClearComments
                    
                    ' od odnalzalem my precious - teraz lecimy z wpisami
                    Set ch = New CommentHandler
                    ch.initWithTxt ir, CStr(f) & Chr(10)
                    Set f = f.Offset(1, 0)
                    
                    Do
                        ch.appendTxt CStr(f) & Chr(10)
                        Set f = f.Offset(1, 0)
                    Loop Until Trim(f) = ""
                    
                    ch.adjustSizeOfThisCmnt
                    
                    Exit Do
                End If
                
                Set f = f.Offset(0, 1)
            Loop Until Trim(f) = ""
        End If
    Next ir
End Sub
