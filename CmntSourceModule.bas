Attribute VB_Name = "CmntSourceModule"
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
