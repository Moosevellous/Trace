Attribute VB_Name = "RibbonControls"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     NamedRangeExists
' Author:   PS
' Desc:     Checks if a given named range exists, returns true if found
' Args:     strRangeName (a named range to be searched for)
' Comments: (1) Called as part of SetSheetTypeControls and elsewhere
'           (2) Used to be called IsNamedRange, which was a bad name
'==============================================================================
Function NamedRangeExists(strRangeName As String) As Boolean
Dim rngExists  As Range
On Error Resume Next
Set rngExists = Range(strRangeName)
NamedRangeExists = True

    If rngExists Is Nothing Then
    NamedRangeExists = False
    msg = MsgBox("Error: Named Range TYPECODE missing!" & chr(10) & chr(10) & _
    "This Trace function requires a blank sheet." & chr(10) & chr(10) _
    & "Try clicking '+ Sheet' on the Trace Ribbon (top left, in the 'New' group).", _
    vbOKOnly, "Sorryyyyyyyyyy")
    End
    End If
    On Error GoTo 0
    
End Function


'NOTE:
'The shorthand for the types is: % -integer; & -long; @ -currency; # -double; ! -single; $ -string


'==============================================================================
'==============================================================================
'UNIVERSAL RIBBON CALLS
'==============================================================================
'==============================================================================

'==============================================================================
' Name:     TraceFunctionWithInput
' Author:   PS
' Desc:     Handles all calls with inputs from the ribbon, except styles and
'           number formats
' Args:     control - set by the Ribbon
' Comments: (1) Calls other functions using Application.Run, feeds in the
'           control.ID as the function input.
'           (2) Sets the sheet parameters using SetSheetTypeControls
'==============================================================================
Sub TraceFunctionWithInput(control As IRibbonControl)

Dim FuncName As String
    
    'check for input
    If Len(control.id) > 0 Then
        'Skip controls for certain functions
        If Left(control.Tag, 1) = "`" Then
        'trim function name of backtick
        FuncName = Right(control.Tag, Len(control.Tag) - 1)
        Else
        FuncName = control.Tag
        SetSheetTypeControls 'set sheet control variables
        End If
    Application.Run FuncName, control.id
    End If
    
End Sub

'==============================================================================
' Name:     TraceFunction
' Author:   PS
' Desc:     Handles all calls without inputs from the ribbon, except styles and
'           number formats
' Args:     control - set by the Ribbon
' Comments: (1) Calls other functions using Application.Run
'==============================================================================
Sub TraceFunction(control As IRibbonControl)

Dim FuncName As String

On Error GoTo errorCatch
'Debug.Print "TypeName: "; TypeName(Selection)

'guard clause for no workbook
    If Selection Is Nothing Then
    Application.Workbooks.Add
    DoEvents
    End If

    If TypeName(Selection) = "Range" Then
        'Skip controls for certain functions
        If Left(control.Tag, 1) = "`" Then
        'trim function name of backtick
        FuncName = Right(control.Tag, Len(control.Tag) - 1)
        Else
        FuncName = control.Tag
        SetSheetTypeControls 'set sheet control variables
        End If
    ElseIf TypeName(Selection) = "ChartArea" Then 'chart object selected
    FuncName = control.Tag
    End If
    
Application.Run FuncName
Exit Sub

'ERRORS GO HERE
errorCatch:
Debug.Print "ERROR: "; Err.Number; Err.Description
'TODO: Pull into own function?
    If Err.Number = 1004 Then
    msg = MsgBox("Error - The function '" & control.Tag & "' was not found!" & _
        chr(10) & "Check the XML and the VBA function name match", _
        vbOKOnly, "Error 1004")
    Exit Sub
'    ElseIf Err.Number = 449 Then
'    msg = MsgBox("Error - The argument'" & control.ID & _
'        "' was not accepted by the function '" & control.Tag & "'", _
'        vbOKOnly, "I came here for an argument")
'    Exit Sub
    Else
    msg = MsgBox("ERROR " & Err.Number & chr(10) & _
        "Description: " & Err.Description & chr(10) & _
        "Function name: " & control.Tag, vbOKOnly, "ERROR")
    Exit Sub
    End If
    
End Sub



'==============================================================================
'==============================================================================
'LOAD MODULE
'==============================================================================
'==============================================================================
Sub btnLoad(control As IRibbonControl)
Dim NameString As String
'Type of sheet-> lose the characters 'btnAdd'
NameString = Right(control.id, Len(control.id) - 3)
    New_Tab (NameString)
End Sub

Sub btnSameType(control As IRibbonControl)
    Same_Type
End Sub

Sub btnStandardCalc(control As IRibbonControl)
    LoadCalcFieldSheet ("Standard")
End Sub

Sub btnFieldSheet(control As IRibbonControl)
    LoadCalcFieldSheet ("Field")
End Sub

Sub btnEquipmentImport(control As IRibbonControl)
    LoadCalcFieldSheet ("EquipmentImport")
End Sub

'==============================================================================
'==============================================================================
'FORMAT / STYLE MODULE
'==============================================================================
'==============================================================================

Sub btnUnits(control As IRibbonControl)
SetSheetTypeControls
SetUnits control.Tag, Selection.Column, 0, _
    Selection.Column + Selection.Columns.Count - 1
End Sub

Sub btnStyle(control As IRibbonControl)
SetSheetTypeControls
SetTraceStyle control.Tag, False
End Sub


'==============================================================================
'==============================================================================
'HELP MODULE
'==============================================================================
'==============================================================================
Sub btnOnlineHelp(control As IRibbonControl)
GetHelp
End Sub

Sub btnAbout(control As IRibbonControl)
frmAbout.Show
End Sub

'==============================================================================
'==============================================================================
'ERROR CATCHING SUBROUTINES
'==============================================================================
'==============================================================================
Public Sub ErrorTypeCode()
msg = MsgBox("Error: Named Range ""TYPECODE"" can't be found or does not match the standard codes." & chr(10) & _
    chr(10) & _
    "Ribbon controls only implemeted for:" & chr(10) & _
    "OCT, OCTA, TO, TOA, MECH, LF_TO, LF_OCT, and CVT" & chr(10) & _
    chr(10) & _
    "Please use a Trace sheet layout.", vbOKOnly, "Waggling finger of shame")
    
End
End Sub

Sub ErrorDoesNotExist()
msg = MsgBox("Error: Feature does not exist yet - please try again later", vbOKOnly, "Maybe one day....?")
End Sub

Sub ErrorOctOnly()
msg = MsgBox("Error: Function only possible in octave bands.", vbOKOnly, "Once.....twice.....three times an octave")
End
End Sub

Sub ErrorThirdOctOnly()
msg = MsgBox("Error: Function only possible in one-third octave bands.", vbOKOnly, "Fool me three times.....")
End
End Sub

Sub ErrorLFTOOnly()
msg = MsgBox("Error: Function only possible in low-frequency one-third octave bands.", vbOKOnly, "All about that bass")
End
End Sub

Sub ErrorOCTTOOnly()
msg = MsgBox("Error: Function only possible in the following Sheet Types: " _
& chr(10) & "OCT / OCTA / TO / TOA", vbOKOnly, "Aw sheet")
End
End Sub

Sub ErrorFrequencyBand()
msg = MsgBox("Error: Frequency band mis-match.", vbOKOnly, "Love Hertz")
End
End Sub

Sub ErrorUnexpectedValue()
msg = MsgBox("Error: Unexpected value.", vbOKOnly, "*confused noise*")
End
End Sub
