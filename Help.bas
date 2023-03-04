Attribute VB_Name = "Help"
'==============================================================================
'==============================================================================
'HELP MODULE
'==============================================================================
'==============================================================================

Sub btnOnlineHelp(control As IRibbonControl)
GetHelp 'TODO: remove this?
End Sub

Sub btnAbout(control As IRibbonControl)
frmAbout.Show
End Sub

Sub btnSettings(control As IRibbonControl)
frmSettings.Show
End Sub

Sub btnFeedback(control As IRibbonControl)
CreateEmail
End Sub

'==============================================================================
' Name:     OpenTraceDirectory
' Author:   PS
' Desc:     Opens Windows Explorer wherever Trace is installed
' Args:     None
' Comments: (1) Where am I?!
'==============================================================================
Sub OpenTraceDirectory()

On Error GoTo catch

Dim TraceDir As String

TraceDir = Application.AddIns("Trace").Path
    If TraceDir <> "" Then
    Call Shell("explorer.exe" & " " & TraceDir, vbNormalFocus)
    End If
    End
    
catch:
Debug.Print "Error: Add-In path not found"

End Sub


'==============================================================================
' Name:     btnContact
' Author:   PS
' Desc:     Creates an email to the document author, based on the custom
'           properties imported into the document, including the boookmark
' Args:     control - object for the button
' Comments: (1)bookmarkRef - name of the bookmark to refer to in the emalil
'==============================================================================
Sub CreateEmail(Optional bookmarkRef As String)
'variables
Dim OL As Object
Dim EmailItem As Object

Dim BodyText As String
Dim SourceBkName As String

'initialise objects
Set OL = CreateObject("Outlook.Application")
Set EmailItem = OL.CreateItem(olMailItem)

'get fields
SourceBkName = ActiveWorkbook.Name

'body of email, now in HTML format
BodyText = "<p style='font-family:Times New Roman;font-size: 11pt'> " & _
                "Hi Trace,<br>" & _
                "<br>" & _
                "This is a message about Trace about: " & "<br>" & _
                SourceBkName & "<br>"
            
    If bookmarkRef <> "" Then
    BodyText = BodyText & "<br>" & _
        "Function Name: <br>" & _
        "Description of issue:" & "<br>" & _
        "<br>" & _
        "Proposed change:"
    End If
    
    'generate email
    With EmailItem
        .Subject = "Trace v" & T_VersionNo & " user feedback"
        .HTMLBody = BodyText
        .To = "philip.setton@wsp.com"
        .Display
    End With

End Sub

