VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FindNewHandler"
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

Private this_w As Workbook
Private source_w As Workbook

Private this_sh As Worksheet
Private source_sh As Worksheet

Private tr As Range
Private sr As Range


Public Sub create_connections(sw As String, ss As String)

    Set this_w = ThisWorkbook
    Set this_sh = this_w.Sheets(WizardMain.MASTER_SHEET_NAME)
    
    Set source_w = Workbooks(sw)
    Set source_sh = source_w.Sheets(ss)
End Sub


Public Sub pomaluj_te_ktorych_nie_ma_w_aktualnym_td(ktora_kolumna_to_pn As Integer)
    
    Set tr = this_sh.Range("A2")
    Set sr = source_sh.Range("A" & CStr(ktora_kolumna_to_pn))
    Set sr = source_sh.Range(sr, sr.End(xlDown))
    
    
    Dim wynik As Range
    firstAdr = ""
    Do
    
        With sr
            Set wynik = .Find(CStr(Trim(tr)))
            
            If Not wynik Is Nothing Then
                firstAdr = wynik.Address
                Do
                    wynik.Interior.Color = RGB(0, 190, 230)
                    Set wynik = .FindNext(wynik)
                Loop While Not wynik Is Nothing And wynik.Address <> firstAdr
            End If
        End With
        Set tr = tr.Offset(1, 0)
    Loop Until Trim(tr) = ""
    
End Sub


