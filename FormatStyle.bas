Attribute VB_Name = "FormatStyle"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public targetType As String
Public targetRange As String
Public targetLimitValue As Single
Public targetMarginValue As Single
Public targetCompliantValue As Single
Public targetRoundingWholeNumber As Boolean
Public targetLimitColour As Long
Public targetMarginColour As Long
Public targetCompliantColour As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'==============================================================================
' Name:     GetStyleRange
' Author:   PS
' Desc:     Returns range address given a row, with switches for parameter
'           columns
' Args:     InputRw - row number
'           isParamCol - set to TRUE to return Parameter range
' Comments: (1) Called by ApplyTraceStyle
'==============================================================================
Function GetStyleRange(InputRw As Integer, Optional isParamCol As Boolean) As String

    If isParamCol Then
    GetStyleRange = Range(Cells(InputRw, T_ParamStart), _
        Cells(InputRw, T_ParamEnd)).Address
    Else 'isParamCol defaults to false
    GetStyleRange = Range(Cells(InputRw, T_Description), _
        Cells(InputRw, T_LossGainEnd)).Address
    End If
 
End Function

'==============================================================================
' Name:     StyleExists
' Author:   PS
' Desc:     Checks if a give Style exists
' Args:     StyleName - name of Style to be checked
' Comments: (1)
'==============================================================================
Function StyleExists(StyleName As String) As Boolean
StyleExists = False

    For i = 1 To ActiveWorkbook.Styles.Count
    'Debug.Print ActiveWorkbook.Styles(i).Name
        If ActiveWorkbook.Styles(i).Name = StyleName Then
        StyleExists = True
        End If
    Next i
    
End Function

'==============================================================================
' Name:     BuildNumDigitsString
' Author:   PS
' Desc:     Returns string for formatting cells with number of digits
' Args:     numDigits - number of digits after the decimal point
' Comments: (1)
'==============================================================================
Function BuildNumDigitsString(numDigits As Integer) As String
Dim i
    If numDigits <= 0 Then
    BuildNumDigitsString = "0"
    Else
    fmtString = "0."
        For i = 1 To numDigits
        fmtString = fmtString & "0"
        Next i
    BuildNumDigitsString = fmtString
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'==============================================================================
' Name:     SetTraceStyle
' Author:   PS
' Desc:     Applied Trace Style to currently selected row
' Args:     StyleName - short name of style
'           isParamCol - set to TRUE to apply to parameter columns
' Comments: (1)
'==============================================================================
Sub SetTraceStyle(StyleName As String, Optional isParamCol As Boolean)
Dim FullStyleName As String
FullStyleName = "Trace " & StyleName
    If IsEmpty(isParamCol) Then
    ApplyTraceStyle FullStyleName, Selection.Row
    Else
    ApplyTraceStyle FullStyleName, Selection.Row, isParamCol
    End If
End Sub

'==============================================================================
' Name:     SetUnits
' Author:   PS
' Desc:     Applies number formats for commonly used units
' Args:     UnitType - unit to be applied, metres, dB etc
'           colStart - First column of range
'           numDigits - number of digits after decimal point
'           colEnd - last column of range
' Comments: (1) Replaces a bunch of individual subroutines
'==============================================================================
Sub SetUnits(UnitType As String, ColStart As Integer, _
    Optional numDigits As Integer, Optional ColEnd As Integer)
    
Dim fmtString As String

    If ColEnd = 0 Then ColEnd = ColStart
fmtString = "NotSet" 'for error catching
    'set fmtString
    Select Case UnitType
    Case Is = "m"
    fmtString = BuildNumDigitsString(numDigits) & " ""m"""
    Case Is = "m2"
    fmtString = BuildNumDigitsString(numDigits) & " ""m" & chr(178) & """"
    Case Is = "mps"
    fmtString = BuildNumDigitsString(numDigits) & " ""m/s"""
    Case Is = "m2ps"
    fmtString = BuildNumDigitsString(numDigits) & " ""m" & chr(178) & "/s"""
    Case Is = "m3ps"
    fmtString = BuildNumDigitsString(numDigits) & " ""m" & chr(179) & "/s"""
    Case Is = "lps"
    fmtString = "0 ""L/s"""
    Case Is = "mm"
    fmtString = BuildNumDigitsString(numDigits) & """ mm"""
    Case Is = "dB"
    fmtString = "0 ""dB"""
    Case Is = "dBA"
    fmtString = "0 ""dBA"""
    Case Is = "kW"
    fmtString = "0 ""kW"""
    Case Is = "MW"
    fmtString = "0 ""MW"""
    Case Is = "Pa"
    fmtString = "0 ""Pa"""
    Case Is = "Q"
    fmtString = "Q=0"
    Case Is = "Clear"
    fmtString = "0" 'default as no decimal places
    End Select

'catch error when Select Case doesn't apply anything
If fmtString = "NotSet" Then fmtString = "General"

'set format to all selected rows
Range(Cells(Selection.Row, ColStart), _
    Cells(Selection.Row + Selection.Rows.Count - 1, ColEnd)) _
    .NumberFormat = fmtString

End Sub

'==============================================================================
' Name:     Target
' Author:   PS
' Desc:     Sets conditional formatting colours
' Args:     None
' Comments: (1)
'==============================================================================
Sub Target()

Dim numconditions As Integer

frmTarget.Show

If btnOkPressed = False Then End

If targetType = "" Then End

    Select Case targetType
    Case Is = "dB"
    targetRange = Cells(Selection.Row, T_LossGainStart - 2).Address
    Case Is = "dBA"
    'Cells(Selection.Row, 4).Value = targetLimitValue
    targetRange = Cells(Selection.Row, T_LossGainStart - 1).Address
    Case Is = "dBC" '<- TODO check for C-weighted range
    targetRange = Cells(Selection.Row, T_LossGainStart - 1).Address
    Case Is = "NR"
    PutNR
    Case Is = "Band"
    '<- TODO band limits
    End Select

Range(targetRange).Select

    If Selection.FormatConditions.Count > 0 Then
    Selection.FormatConditions.Delete
    End If

numconditions = 1

    If targetLimitValue <> 0 Then
        With Selection
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=" & targetLimitValue
        .FormatConditions(numconditions).Interior.Color = targetLimitColour
        End With
    numconditions = numconditions + 1
    End If
    
    If targetCompliantValue <> 0 Then
        With Selection
        'marginal format
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
            Formula1:=targetMarginValue, Formula2:=targetLimitValue
        .FormatConditions(numconditions).Interior.Color = targetMarginColour
        numconditions = numconditions + 1
        'compliant format
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _
            Formula1:="=" & targetCompliantValue
        .FormatConditions(numconditions).Interior.Color = targetCompliantColour
        End With
    End If
    
'I think this line is just good housekeeping?
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

End Sub

'==============================================================================
' Name:     ImportTraceStyles
' Author:   PS
' Desc:     Imports styles from template sheets directory
' Args:     None
' Comments: (1)
'==============================================================================
Sub ImportTraceStyles()
Dim BlankBookName As String
Dim CurrentBookName As String
Dim StylesSheet As String
'set variables
GetSettings
CurrentBookName = ActiveWorkbook.Name
StylesSheet = TEMPLATELOCATION & "\STYLE.xlsm"
'open the style sheet
Application.Workbooks.Open (StylesSheet)
BlankBookName = ActiveWorkbook.Name
Workbooks(CurrentBookName).Activate
'merge styles
ActiveWorkbook.Styles.Merge (BlankBookName)
Workbooks(BlankBookName).Close (False)
End Sub

'==============================================================================
' Name:     ApplyTraceStyle
' Author:   PS
' Desc:     Checks is style exists, then applies the style to the row.
' Args:     StyleName - Full name of style
'           ApplyTorow
' Comments: (1)
'==============================================================================
Sub ApplyTraceStyle(StyleName As String, ApplyToRow As Integer, _
Optional isParamCol As Boolean) '<--TODO set range description, not just isParamCol
Dim askForStyleImport As Integer

    If StyleExists(StyleName) = False Then
    askForStyleImport = MsgBox("No styles in this document. Do you want to import?", _
        vbYesNo, "Style(ish)!")
        
        If askForStyleImport = vbYes Then
        ImportTraceStyles
        Else
        End
        End If
        
    End If
    
Range(GetStyleRange(ApplyToRow, isParamCol)).Style = StyleName
    
'A-weighted column goes bold!
Cells(ApplyToRow, T_LossGainStart - 1).Font.Bold = True

    If T_RegenStart > 0 Then 'skips if value is -1
    Cells(ApplyToRow, T_RegenStart - 1).Font.Bold = True
    End If
    
End Sub

'==============================================================================
' Name:     DeleteNonTraceStyle
' Author:   PS
' Desc:     Removes any style that doesn't contain the string "Trace"
' Args:     None
' Comments: (1) Skips "Normal" style
'==============================================================================
Sub DeleteNonTraceStyles()
Dim LastStyleIndex As Integer
Dim i As Integer

LastStyleIndex = ActiveWorkbook.Styles.Count
  For i = 1 To LastStyleIndex
    'Debug.Print ActiveWorkbook.Styles(i).Name
        If InStr(1, ActiveWorkbook.Styles(i).Name, "Trace", vbTextCompare) = 0 _
            And ActiveWorkbook.Styles(i).Name <> "Normal" Then
        ActiveWorkbook.Styles(i).Delete
        i = i - 1 'don't skip
        LastStyleIndex = LastStyleIndex - 1
        End If
    Next i
End Sub
