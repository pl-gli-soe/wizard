VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PUSBUcket"
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

Private earliest_sdate As Date
Private latest_eda As Date

Private in_transit_mrd As Boolean
Private in_transit_today As Boolean
Private in_transit_custom As Boolean

Private indx As String

Public pus_collection As Collection
Private pus_item As PUSItemForPUSBucket

Private firstMRDDate As Date
Private todayDate As Date
Private customDate As Date


Public Sub setIndx(s As String)
    indx = CStr(s)
End Sub


Public Sub search_pus_sheet_and_prepare_bucket(psh As Worksheet)
    
    Dim r As Range
    Set r = psh.Cells(2, WizardMain.O_INDX)
    
    
    
    Do
        If CStr(indx) = Trim(CStr(r)) Then
        
        
            Set pus_item = Nothing
            Set pus_item = New PUSItemForPUSBucket
            
            With pus_item
                .pn = psh.Cells(r.Row, WizardMain.O_PN)
                .duns = psh.Cells(r.Row, WizardMain.O_DUNS)
                .fup_code = psh.Cells(r.Row, WizardMain.O_FUP_code)
                .nm = psh.Cells(r.Row, WizardMain.O_PUS_Number)
                
                .pus_date = CDate(psh.Cells(r.Row, WizardMain.O_Pick_up_date))
                .eda_date = CDate(psh.Cells(r.Row, WizardMain.O_Delivery_Date))
                
                .qty = CLng(psh.Cells(r.Row, WizardMain.O_Pick_up_Qty))
            End With
            
            pus_collection.Add pus_item
            
            Me.checkThisEDAAndAssignItIfItsLatest CDate(psh.Cells(r.Row, WizardMain.O_Delivery_Date))
            Me.checkThisSDateAndAssignItIfItsEarliest CDate(psh.Cells(r.Row, WizardMain.O_Pick_up_date))
            Me.setInTransit_customDate
            Me.setInTransit_todayDate
            Me.setInTransit_firstMRDDate
        End If
        Set r = r.Offset(1, 0)
    Loop Until r.Row = WizardMain.POLOWA_CAPACITY_ARKUSZA
End Sub

Public Function getIndx() As String
    getIndx = CStr(indx)
End Function


Public Function getInTransitMrd() As Boolean
    getInTransitMrd = in_transit_mrd
End Function

Public Function getInTransitToday() As Boolean
    getInTransitToday = in_transit_today
End Function

Public Function getInTransitCustom() As Boolean
    getInTransitCustom = in_transit_custom
End Function

Public Sub setInTransit_todayDate()
    
    If todayDate < latest_eda Then
    
        in_transit_today = True
    Else
        in_transit_today = False
    End If
End Sub

Public Sub setInTransit_firstMRDDate()
    
    If firstMRDDate < latest_eda Then
    
        in_transit_mrd = True
    Else
        in_transit_mrd = False
    End If
End Sub


Public Sub setInTransit_customDate()
    
    If customDate < latest_eda Then
    
        in_transit_custom = True
    Else
        in_transit_custom = False
    End If
End Sub


Public Function getEarliestSDate() As Date
    getEarliestSDate = earliest_sdate
End Function

Public Function getLatestEDA() As Date
    getLatestEDA = latest_eda
End Function


Public Sub checkThisSDateAndAssignItIfItsEarliest(d As Date)
    
    If d < earliest_sdate Then
        earliest_sdate = d
    End If
End Sub

Public Sub checkThisEDAAndAssignItIfItsLatest(d As Date)
    If d > latest_eda Then
        latest_eda = d
    End If
End Sub



Public Sub init(m_mrd_date As Date, m_custom_date As Date, m_today_date As Date)
    firstMRDDate = m_mrd_date
    todayDate = m_today_date
    customDate = m_custom_date
    ' heurystycznie
    earliest_sdate = Date + 10000
    latest_eda = 1
End Sub




Private Sub Class_Initialize()
    Set pus_collection = New Collection
End Sub

Private Sub Class_Terminate()
    Set pus_collection = Nothing
End Sub
