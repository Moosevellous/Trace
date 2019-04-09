Attribute VB_Name = "OnlineHelp"
Sub GetHelp()

Set sh = ActiveSheet

    If sh Is Nothing Then
    Workbooks.Add
    End If
    
ActiveWorkbook.FollowHyperlink _
        Address:="https://github.com/Moosevellous/Trace/wiki", _
        NewWindow:=True, _
        AddHistory:=True
Application.WindowState = xlNormal

End Sub


