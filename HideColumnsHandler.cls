VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HideColumnsHandler"
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

Private m As Worksheet

Private Sub Class_Initialize()
    Set m = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
End Sub

Private Sub inner_show_poka(b As Boolean)

    Dim r As Range
    
    Set r = m.Cells(1, WizardMain.MRD2_QTY)
    r.EntireColumn.Hidden = b
    
    Set r = m.Range(m.Cells(1, WizardMain.MRD2_Ordered_date), m.Cells(1, WizardMain.MRD2_Total_PUS_STATUS))
    r.EntireColumn.Hidden = b
    
End Sub

Public Sub pokaz_kolumny_dla_mrd2()
    inner_show_poka False
End Sub

Public Sub showaj_kolumny_dla_mrd2()
    inner_show_poka True
End Sub

