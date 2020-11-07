Attribute VB_Name = "OnlineHelp"
'==============================================================================
' Name:     GetHelp
' Author:   PS
' Desc:     Open the Wiki page on the GitHub site
' Args:     None
' Comments: (1) The link goes to the wiki, which says:

' Welcome to the Trace wiki!
'
' Trace has been developed with Acoustic Engineers in mind, with tools and
' templates developed as needed. All results are intended to be presentable
' as an Appendix and all calculations are completely transparent.
'
' Trace was first developed by [WSP](https://www.wsp.com/en-AU) in Melbourne
' Australia. Other calculation toolkits for Acoustics have been developed as
' proprietary software, however this is the first to be developed as an open
' source project.
'
' Trace has lots of potential to be used and expanded, and contributors are
' asked to help in any of the following ways:
' - Developing new functions
' - Testing the system
' - Writing the Wiki
' - Requesting additional functionality
'
' Contact us through this page for more or alternatively email
' philip.setton@wsp.com

'==============================================================================
Sub GetHelp(Optional PathStr As String)

Set sh = ActiveSheet

    If sh Is Nothing Then
    Workbooks.Add
    End If

GotoWikiPage

End Sub

'==============================================================================
' Name:     GotoWikiPage
' Author:   PS
' Desc:     Opens a browser window with the default path, or to correct area
'           of the wiki if the variable PathStr is provided
' Args:     PathStr - URL string to be appended
' Comments: (1)
'==============================================================================
Sub GotoWikiPage(Optional PathStr As String)
Dim LinkPath As String

LinkPath = "https://github.com/Moosevellous/Trace/wiki/" & PathStr

ActiveWorkbook.FollowHyperlink Address:=LinkPath, NewWindow:=True
End Sub
