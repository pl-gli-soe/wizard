Attribute VB_Name = "DeleteModule"
Public Sub usun_ten_arkusz(ictrl As IRibbonControl)
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If (ActiveSheet.Name Like "*MASTER*") Or (ActiveSheet.Name Like "*DETAILS*") Or _
        (ActiveSheet.Name Like "*PICKUPS*") Or (ActiveSheet.Name Like "*register*") Or _
        (ActiveSheet.Name Like "*config*") Or (ActiveSheet.Name Like "*delivery_confirmation_special*") Or _
        (ActiveSheet.Name Like "*custom_copy*") Or (ActiveSheet.Name Like "*comment_source*") Or _
        (ActiveSheet.Name Like "*CACHE*") Then
            
            MsgBox "you can't delete this sheet!"
    Else
        ActiveSheet.Delete
    End If
    
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


Public Sub usun_wszystkie_arkusze(ictrl As IRibbonControl)

    Application.EnableEvents = False

    ret = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If ret = vbYes Then
    
    
        Application.DisplayAlerts = False
        
        x = 1
        Do
            If (ActiveSheet.Name Like "*MASTER*") Or (ActiveSheet.Name Like "*DETAILS*") Or _
                (ActiveSheet.Name Like "*PICKUPS*") Or (ActiveSheet.Name Like "*register*") Or _
                (ActiveSheet.Name Like "*config*") Or (ActiveSheet.Name Like "*delivery_confirmation_special*") Or _
                (ActiveSheet.Name Like "*custom_copy*") Or (ActiveSheet.Name Like "*comment_source*") Or _
                (ActiveSheet.Name Like "*CACHE*") Then
                    
                    x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.Count
        Application.DisplayAlerts = True
    End If
    
    
    Application.EnableEvents = True
End Sub
