Attribute VB_Name = "DevTools"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     findTraceVBProject
' Author:   PS
' Desc:     Returns a the location of the Trace.xlam addin file
' Args:     None
' Comments: (1) Called from EXPORT_TRACE_SOURCE_CODE
'==============================================================================
Function findTraceVBProject()
    For i = 1 To Application.VBE.VBProjects.Count
    'Debug.Print Application.VBE.VBProjects(i).Name
        If Application.VBE.VBProjects(i).Name = "Trace3" Then
        findTraceVBProject = i
        Exit Function
        End If
    Next i
End Function

'==============================================================================
' Name:     GetFolder
' Author:   PS
' Desc:     Returns a folder from a user input
' Args:     None
' Comments: (1) Called from EXPORT_TRACE_SOURCE_CODE
'==============================================================================
Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     EXPORT_TRACE_SOURCE_CODE
' Author:   PS
' Desc:     Saves all modules and forms as .bas files, for use in Git etc
' Args:     None
' Comments: (1) An essential part of tracking differences in the code.
'==============================================================================
Sub EXPORT_TRACE_SOURCE_CODE()
Dim TraceIndex As Integer
Dim fldr As String
Dim numFiles As Integer
Dim TraceComponent As Object
Dim SavePath As String

TraceIndex = findTraceVBProject 'calls function to find which add-in is Trace

    If TraceIndex = 0 Then
    msg = MsgBox("Error, can't find Trace Add-in index. " & chr(10) & _
        "Try closing opening Excel.", vbOKOnly, "Add-in index error")
    End
    End If
    
numFiles = 0

fldr = GetFolder

If fldr = "" Then End

    For Each TraceComponent In Application.VBE.VBProjects(TraceIndex).VBComponents
    
    Debug.Print "Name: "; TraceComponent.Name & "     Type: " & TraceComponent.Type
    Application.StatusBar = "Exporting: " & TraceComponent.Name
        
        'only export modules and forms
        If TraceComponent.Type = 1 Or TraceComponent.Type = 3 Then
        
            If Left(TraceComponent.Name, 3) = "frm" Then 'put in forms subfolder
            'folder doesn't exist, make one!
                If Dir(fldr & "\form", vbDirectory) = Empty Then
                MkDir fldr & "\form"
                End If
            SavePath = fldr & "\form\" & TraceComponent.Name & ".bas"
            Else
            SavePath = fldr & "\" & TraceComponent.Name & ".bas"
            End If
        
        'Debug.Print SavePath
        TraceComponent.Export (SavePath)
        Debug.Print "EXPORTED"
        numFiles = numFiles + 1
        Else
        Debug.Print "SKIPPED"
        End If
    Next

msg = MsgBox("Export process complete: " & numFiles & " files", vbOKOnly, _
    "Dev Tools - Export")

Application.StatusBar = False

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


