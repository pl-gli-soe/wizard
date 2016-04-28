Attribute VB_Name = "DeleteModule"
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
