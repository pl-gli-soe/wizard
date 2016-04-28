Attribute VB_Name = "GlobalModule"
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


' pod niestd podmiane
Global Const G_WYBIERZ_KOLUMNE_PN = 2
Global Const G_WYBIERZ_KOLUMNE_CMNTS = 14
Global Const G_WYBIERZ_KOLUMNE_CUSTOMOWA = 28


Global details_handler As DetailsHandler
Global pickups_handler As PickupsHandler
Global dcs_handler As DCSHandler

Global Const MASTER_SHEET_NAME = "MASTER"
Global Const CONFIG_SHEET_NAME = "config"
Global Const CUSTOM_COPY_SHEET_NAME = "custom_copy"
Global Const COMMENT_SOURCE_SHEET_NAME = "comment_source"
Global Const DELIVERY_CONFIRMATION_SPECIAL_SHEET_NAME = "delivery_confirmation_special"
Global Const REGISTER_SHEET_NAME = "register"
Global Const DETAILS_SHEET_NAME = "DETAILS"
Global Const ORDERS_SHEET_NAME = "ORDERS"
Global Const PICKUPS_SHEET_NAME = "PICKUPS"
Global Const DCS_SHEET_NAME = "delivery_confirmation_special"


Global Const DCS = "delivery_confirmation_special"
' this is pointer for details sheet during working on init wizard on proj def
Global Const POINTER = "POINTER"


Global Const G_NO_LIST_IN_CACHE = "No list in cache!"
Global Const G_NO_PLATFORMS_TXT = "No Platforms! (Only in PPV needed)"
Global Const FMA_STR = "FMA"
Global Const FMA_WITH_STARS = "*FMA*"
Global Const COMMA = ","
Global Const RAW_TXT = "raw"
Global Const ENTER_STR = "ENTER"
Global Const NIC_NIE_WYBRANO_TXT = "nic nie wybrano"
Global Const PREFIX_TXT_EDIT_PUS_DEL_CONF = "Del Conf for selected PN: "
Global Const PREFIX_TXT_EDIT_PUS_FST_PUS_DATE = "First PUS Date for selected PN: "
Global Const TBD = "tbd"
Global Const CACHE = "CACHE"
Global Const GHOST = "GHOST"
Global Const DCS_STR = "delivery_confirmation_special"

Global Const MRD_KLUCZ_DO_PODMIANY = "{MRD}"

Global Const MIN_LEN_PROJ_NAME = 4

Global Const SELECTION_LIMIT = 256
' 2^14
Global Const TOP_EDIT_LIMIT = 16384
' Global Const TOP_EDIT_LIMIT = 50
Global Const ASCII_0 = 48
Global Const ASCII_9 = 57
Global Const ASCII_ENTER = 13

Global Const LISTBOX_CUSTOM_COLUMN_NAMES_LIMIT = 20

Global Const ALL_ORDERED_QTY = "ALL Ordered Qty"

Global Const G_PASS = "1985-07-10"

Global Const G_HOW_MANY_ROWS_WILL_BE_DELETED = 524288 ' 2^19 polowa capacity akursza excela
Global Const POLOWA_CAPACITY_ARKUSZA = 524288 ' 2^19 polowa capacity akursza excela
Global Const CAPACITY_ARKUSZA = 1048576

Global Const DWA_DO_16 = 65536 ' 2^10 polowa capacity akursza excela

Global Const SIX = 6

Global Const G_STEP_BETWEEN_PARALELL_USERS = 40000
' nawet polowa bufferu :D
Global Const USERS_LIMIT = 8

Public Function fnDateFromWeek(iYear As Integer, iWeek As Integer, iWeekDday As Integer)
    ' get the date from a certain day in a certain week in a certain year
      fnDateFromWeek = CDate(CStr(iYear) & "-01-01")
      
      Do
        fnDateFromWeek = fnDateFromWeek + 1
      Loop Until Int(Application.WorksheetFunction.IsoWeekNum(fnDateFromWeek)) = Int(iWeek)
      
End Function


' global sub
Public Sub nowy_schemat_offsetu_w_arkuszu_pickups(ByRef i As Range)


    Set i = i.Offset(1, 0)
    If Trim(i) = "" Then
        Set i = i.End(xlDown)
    End If
End Sub




' tego jeszcze nie testowalem :(
Public Sub users_status_usun_moje_stare_instancje(u As String)

    '    MsgBox Application.UserName
    '
    Dim d As Date
    ile = 0
    Users = ThisWorkbook.UserStatus
    With ThisWorkbook.ActiveSheet
        For x = 1 To UBound(Users, 1)
        
            If Users(x, 1) = CStr(u) Then
                ile = ile + 1
            End If
        Next x
    End With
    
    ' jakos tak sie sklada ze mam wiecej instancji
    
    If ile > 1 Then
        x = 1
        y = 1
        Do
            y = 1
            Do
                If Users(y, 1) = CStr(u) Then
                    Users(x).Delete
                    x = x + 1
                    Exit Do
                End If
                
                y = y + 1
            Loop Until y = UBound(Users, 1)
        Loop Until x = ile

    End If
    
    
    
    
    
    
    
    
    
    
    '        For Row = 1 To UBound(Users, 1)
    '            .Cells(Row, 1) = Users(Row, 1)
    '            .Cells(Row, 2) = Users(Row, 2)
    '            Select Case Users(Row, 3)
    '                Case 1
    '                    .Cells(Row, 3).Value = "Exclusive"
    '                Case 2
    '                    .Cells(Row, 3).Value = "Shared"
    '            End Select
    '        Next
    '    End With
    
    

End Sub

