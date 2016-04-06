VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public details_dictionary As Dictionary
Private details_range As Range
Private rkey As Range


Private dfm As DetailsFormsManager
Private e As E_NEW_PROJECT_ITEM

Private sh As Worksheet


Public Function get_e() As E_NEW_PROJECT_ITEM
    get_e = e
End Function
    


Private Sub Class_Initialize()

    Set details_dictionary = New Dictionary
    Set details_range = ThisWorkbook.Sheets(DETAILS_SHEET_NAME).Range("a1")
    Set details_range = ThisWorkbook.Sheets(DETAILS_SHEET_NAME).Range(details_range, details_range.End(xlDown))
    
    For Each rkey In details_range
        If Not details_dictionary.Exists(CStr(rkey)) Then
            details_dictionary.Add CStr(rkey), CStr(rkey.Offset(0, 1))
        End If
    Next rkey
    
    Set dfm = New DetailsFormsManager
End Sub

Private Sub Class_Terminate()
    Set details_dictionary = Nothing
End Sub


Public Sub init_wizard_for_details(new_definition As Boolean)
    
    
    
    Set sh = ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
    'sh.Unprotect 123
    
    
    
    dfm.new_definition = new_definition
    
    
    e = PLT
    dfm.init e, Me
    
    Set dfm = Nothing
    
    
    'sh.Protect 123
End Sub


' where arg is me
Public Sub dalej(arg As Variant)

    arg.hide
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
    
    
    ' lecimy po kolei z updatem pol
    With ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
        
        details_range.item(e).Offset(0, 1).Value = arg.TextBoxBufor.Value
    End With
    ' dodatkowa logika sprawdzajaca czy nie przekroczylismy przypadkiem juz naszych pol
    e = e + 1
    
    If e > E_UNIQUE_ID Then
        MsgBox "Gotowe!" & Chr(10) & _
            "Twoje wybory sa widoczne w arkuszu details"
        
        'ThisWorkbook.Activate
        'Sheets(DETAILS_SHEET_NAME).Select
        'Range("A1").Select
    Else
    
        dfm.init e, Me
    End If
    
    Set arg = Nothing

End Sub

' where arg is me
Public Sub cofnij(arg As Variant)

    arg.hide
    ThisWorkbook.Sheets(CONFIG_SHEET_NAME).Range("form_activatedd") = 0
    If e = PLT Then
        MsgBox "ten msgbox nigdy nie powinien sie pojawic - nie jestesmy wstanie confac sie dalej niz do plantu!"
    Else
    
        
        ' lecimy po kolei z updatem pol
        With ThisWorkbook.Sheets(DETAILS_SHEET_NAME)
            
            details_range.item(e).Offset(0, 1).Value = arg.TextBoxBufor.Value
        End With
    
        e = e - 1
        dfm.init e, Me
    End If


    Set arg = Nothing
End Sub
