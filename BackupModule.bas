Attribute VB_Name = "BackupModule"
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
