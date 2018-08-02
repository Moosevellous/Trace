Attribute VB_Name = "OnlineHelp"
Sub GetHelp()
ActiveWorkbook.FollowHyperlink _
        Address:="https://github.com/Moosevellous/Trace/wiki", _
        NewWindow:=True, _
        AddHistory:=True
    Application.WindowState = xlNormal
End Sub
