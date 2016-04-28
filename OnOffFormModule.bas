Attribute VB_Name = "OnOffFormModule"
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

Public Sub on_off_form()
    ' Debug.Print "on off"
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets("config").Range("quick_forms")
    
    
    If r = 0 Then
        r = 1
        MsgBox "Formularze podpowiadajace sa teraz aktywne"
    ElseIf r = 1 Then
        r = 0
        MsgBox "Formularze podpowiadajace sa teraz nieaktywne"
    Else
        MsgBox "cfg quick forms - ten msgbox nigdy nie powinien sie pokazac!"
    End If

End Sub


Public Sub on_off_quick_form(ictrl As IRibbonControl)
    on_off_form
End Sub
