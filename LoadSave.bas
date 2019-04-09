Attribute VB_Name = "CODE_EXPORT"
Function findTraceVBProject()
    For i = 1 To Application.VBE.VBProjects.count
    If Application.VBE.VBProjects(i).Name = "Trace" Then findTraceVBProject = i
    Next i
End Function

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




'''''''''''''''''''''''''''''''''''''''''''''''
'RUN THIS CODE TO EXPORT
'''''''''''''''''''''''''''''''''''''''''''''''
Sub EXPORT_TRACE_SOURCE_CODE()
'For Each VbComp In ActiveWorkbook.VBProject.VBComponents
Dim TraceIndex As Integer
Dim fldr As String
Dim numfiles As Integer

TraceIndex = findTraceVBProject 'calls function to find which add-in is Trace
numfiles = 0

fldr = GetFolder

If fldr = "" Then End

    For Each TraceComponent In Application.VBE.VBProjects(TraceIndex).VBComponents
    
    Debug.Print "Name: "; TraceComponent.Name & " Type: " & TraceComponent.Type
        If TraceComponent.Type = 1 Or TraceComponent.Type = 3 Then 'modules and forms
        Debug.Print "EXPORTED"
        'MkDir "C:\Users\AUPS02932\Documents\Development\Trace\EXPORT\"
        SavePath = fldr & "\" & TraceComponent.Name & ".txt"
        TraceComponent.EXPORT (SavePath)
        numfiles = numfiles + 1
        Else
        Debug.Print "SKIPPED"
        End If
    Next

msg = MsgBox("Process complete. " & numfiles & " files exported", vbOKOnly, "Source Code Export")

End Sub


