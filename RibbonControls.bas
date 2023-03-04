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
Function NamedRangeExists(strRangeName As String, Optional skipErrorMessage) As Boolean
Dim rngExists  As Range
On Error Resume Next
Set rngExists = Range(strRangeName)
NamedRangeExists = True

    If rngExists Is Nothing Then
    NamedRangeExists = False
        If skipErrorMessage = False Then
        msg = MsgBox("Error: Named Range TYPECODE missing!" & chr(10) & chr(10) & _
        "This Trace function requires a blank sheet." & chr(10) & chr(10) _
        & "Try clicking '+ Sheet' on the Trace Ribbon (top left, in the 'New' group).", _
        vbOKOnly, "Sorryyyyyyyyyy")

        End If
    End If
    On Error GoTo 0
    
End Function

'==============================================================================
' Name:     CheckFunctionName
' Author:   PS
' Desc:     Checks for backtick which denotes skipping setting the sheet types
' Args:
' Comments: (1)
'==============================================================================
Function CheckFunctionName(inputString As String, Optional SkipSetSheetTypes As Boolean)

    'Skip setting controls for certain functions
    If Left(inputString, 1) = "`" Then
    'trim function name of backtick
    CheckFunctionName = Right(inputString, Len(inputString) - 1)
    SkipSetSheetTypes = True 'backtick over-rides the default input
    Else
    CheckFunctionName = inputString
    End If
    
    'Some functions don't need sheet controls to be set
    If SkipSetSheetTypes = False Then 'don't skip
    SetSheetTypeControls
    End If
    
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
    FuncName = CheckFunctionName(control.Tag)
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
    FuncName = CheckFunctionName(control.Tag)
    ElseIf TypeName(Selection) = "ChartArea" Then 'chart object selected
    FuncName = CheckFunctionName(control.Tag, True)
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
'ERROR CATCHING SUBROUTINES
'==============================================================================
'==============================================================================
Public Sub ErrorTypeCode()
MsgBox "Error: Named Range ""TYPECODE"" can't be found or does not match the standard codes." & chr(10) & _
    chr(10) & _
    "Ribbon controls only implemeted for:" & chr(10) & _
    "OCT, OCTA, TO, TOA, MECH, LF_TO, LF_OCT, and CVT" & chr(10) & _
    chr(10) & _
    "Please use a Trace sheet layout.", vbOKOnly, "Waggling finger of shame"
    
End
End Sub

Sub ErrorDoesNotExist()
MsgBox "Error: Feature does not exist yet - please try again later", _
    vbOKOnly, "Maybe one day....?"
End Sub

Sub ErrorOctOnly()
MsgBox "Error: Function only possible in octave bands." & chr(10) & _
    "Try adding an a new octave band sheet (Sheet>OCT).", vbOKOnly, _
    "Once.....twice.....three times an octave"
End
End Sub

Sub ErrorThirdOctOnly()
MsgBox "Error: Function only possible in one-third octave bands." & chr(10) & _
    "Try adding an a new one-third octave band sheet (Sheet>TO).", vbOKOnly, _
    "Fool me three times....."
End
End Sub

Sub ErrorLFTOOnly()
MsgBox "Error: Function only possible in low-frequency one-third octave bands." & chr(10) & _
    "Try adding an a new low frequency sheet (Sheet>LF_TO).", vbOKOnly, _
    "All about that bass"
End
End Sub

Sub ErrorOCTTOOnly()
MsgBox "Error: Function only possible in the following Sheet Types: " & chr(10) _
    & "OCT / OCTA / TO / TOA", vbOKOnly, "Aw sheet"
End
End Sub

Sub ErrorFrequencyBand()
MsgBox "Error: Frequency band mis-match.", vbOKOnly, "Love Hertz"
End
End Sub

Sub ErrorUnexpectedValue()
MsgBox "Error: Unexpected value.", vbOKOnly, "*confused noise*"
End
End Sub

Sub ErrorFrequencyBandMissing()
MsgBox "Start or end bands are missing or may be switched off. " & chr(10) & _
    "Check the active range in the dropdown menu in Basics Group on the ribbon", _
    vbOKOnly, "Error - frequency bands"
End
End Sub
