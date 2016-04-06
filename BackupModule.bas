Attribute VB_Name = "BackupModule"
Public Sub create_backup()
    
    If ThisWorkbook.Sheets("config").Range("backup_system").Value = 1 Then
        ThisWorkbook.SaveCopyAs Filename:="" _
                & ThisWorkbook.FullName & "_backup__timestamp_" & _
                Replace(CStr(CDbl(Now)), ",", "_") & ".xlsm"
    End If
End Sub


Public Sub zrob_backup(ictrl As IRibbonControl)

    ThisWorkbook.Sheets("config").Range("backup_system").Value = 1
    create_backup
    ThisWorkbook.Sheets("config").Range("backup_system").Value = 0
End Sub
