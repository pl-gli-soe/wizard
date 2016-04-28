Attribute VB_Name = "G_Gen_ID_Module"
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

Public Function generate_unique_id(r As Range) As String
    generate_unique_id = ""
    
    Do
        If CStr(r.Offset(0, -1).Value) <> CStr(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Cells(19, 1)) Then
            generate_unique_id = generate_unique_id & parse_to_num(CStr(r))
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function


Public Function parse_to_num(s As String) As String
    parse_to_num = ""
    For x = 1 To Len(s)
        parse_to_num = parse_to_num & CStr(Asc(Mid(s, x, 1)))
    Next x
    
End Function


Public Sub generuj_id(ictrl As IRibbonControl)
    Dim r As Range
    Set r = ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Range("b19")
    
    r = generate_unique_id(ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME).Range("b1"))
    
    
End Sub
