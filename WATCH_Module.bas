Attribute VB_Name = "WATCH_Module"
Public Sub add_watch(ictrl As IRibbonControl)
    addWatch
End Sub


Private Sub addWatch()
Attribute addWatch.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim wh As WatchHandler
    Set wh = New WatchHandler
    
    wh.createWatchForThisLine ActiveCell
    
    Set wh = Nothing
End Sub
