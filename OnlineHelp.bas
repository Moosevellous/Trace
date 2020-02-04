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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'The link goes to the wiki, which says:


'Welcome to the Trace wiki!
'
'Trace has been developed with Acoustic Engineers in mind, with tools and templates developed as needed. All results are intended to be presentable as an
'Appendix and all calculations are completely transparent.
'
'Trace was first developed by [WSP](https://www.wsp.com/en-AU) in Melbourne Australia. Other calculation toolkits for Acoustics have been developed as
'proprietary software, however this is the first to be developed as an open source project.
'
'Trace has lots of potential to be used and expanded, and contributors are asked to help in any of the following ways:
'- Developing new functions
'- Testing the system
'- Writing the Wiki
'- Requesting additional functionality
'
'Contact us through this page for more or alternatively email philip.setton@wsp.com

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
