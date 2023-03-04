Attribute VB_Name = "FormatStyle"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public targetType As String
Public targetBands As Boolean
Public targetRange As String
Public targetLimitValue As Single
Public targetMarginValue As Single
Public targetCompliantValue As Single
Public targetRoundingWholeNumber As Boolean
Public targetLimitColour As Long
Public targetMarginColour As Long
Public targetCompliantColour As Long
Public ApplyHeatMap As Boolean

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
        Exit Function
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
Function BuildNumDigitsString(NumDigits As Integer) As String
Dim i
    If NumDigits <= 0 Then
    BuildNumDigitsString = "0"
    Else
    fmtString = "0."
        For i = 1 To NumDigits
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
Dim rw As Integer
FullStyleName = "Trace " & StyleName
    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        If IsEmpty(isParamCol) Then
        ApplyTraceStyle FullStyleName, rw
        Else
        ApplyTraceStyle FullStyleName, rw, isParamCol
        End If
    Next rw
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
    Optional NumDigits As Integer, Optional ColEnd As Integer)
    
Dim fmtString As String

    If ColEnd = 0 Then ColEnd = ColStart
fmtString = "NotSet" 'for error catching
    'set fmtString
    Select Case UnitType
    Case Is = "m"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m"""
    Case Is = "m2"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m" & chr(178) & """"
    Case Is = "m3"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m" & chr(179) & """"
    Case Is = "mps"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m/s"""
    Case Is = "m2ps"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m" & chr(178) & "/s"""
    Case Is = "m3ps"
    fmtString = BuildNumDigitsString(NumDigits) & " ""m" & chr(179) & "/s"""
    Case Is = "lps"
    fmtString = "0 ""L/s"""
    Case Is = "mm"
    fmtString = BuildNumDigitsString(NumDigits) & """ mm"""
    Case Is = "dB"
    fmtString = "0 ""dB"""
    Case Is = "dBA"
    fmtString = "0 ""dBA"""
    Case Is = "kW"
    fmtString = BuildNumDigitsString(NumDigits) & """ kW"""
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

'Set conditional formatting to a range
'TODO: set colouring for all bands
    Select Case targetType
    Case Is = "dB"
    targetRange = Cells(Selection.Row, T_LossGainStart - 2).Address
    Case Is = "dBA"
    'Cells(Selection.Row, 4).Value = targetLimitValue
    targetRange = Cells(Selection.Row, T_LossGainStart - 1).Address
    Case Is = "dBC" '<- TODO check for C-weighted range
    targetRange = Cells(Selection.Row, T_LossGainStart - 1).Address
    Case Is = "NR"
    MoveDown
    PutNR
    targetRange = Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Address
    End Select

'clear previous conditional formatting
Range(targetRange).Select

    If Selection.FormatConditions.Count > 0 Then
    Selection.FormatConditions.Delete
    End If

numconditions = 1

    'Add conditional formatting for limit colours
    If targetLimitValue <> 0 Then
        With Selection
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=" & targetLimitValue
        .FormatConditions(numconditions).Interior.Color = targetLimitColour
        End With
    numconditions = numconditions + 1
    End If
    
    'Add conditional formatting for marginal colours
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
' Comments: (1) TODO: make this apply style to Regen columns in a clever way?
'==============================================================================
Sub ApplyTraceStyle(StyleName As String, ApplyToRow As Integer, _
Optional isParamCol As Boolean) '<--TODO: set range description, not just isParamCol
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

'==============================================================================
' Name:     ApplyTraceMarker
' Author:   PS
' Desc:     Adds symbols to first column
' Args:     None
' Comments: (1) Formerly named 'MarkRowAs'
'==============================================================================
Sub ApplyTraceMarker(MarkerType As String)

'catch function calls from ribbon which start with 'Mrk'
If Left(MarkerType, 3) = "Mrk" Then
    MarkerType = Right(MarkerType, Len(MarkerType) - 3)
End If

For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
    Select Case MarkerType
    Case Is = "Clear"
    Cells(rw, 1).ClearContents
    Case Is = "Sum"
    Cells(rw, 1) = ChrW(T_MrkSum)
    Case Is = "Average"
    Cells(rw, 1) = ChrW(T_MrkAverage)
    Case Is = "Silencer"
    Cells(rw, 1) = ChrW(T_MrkSilencer)
    Case Is = "Louvre"
    Cells(rw, 1) = ChrW(T_MrkLouvre)
    Case Is = "Result"
    Cells(rw, 1) = ChrW(T_MrkResult)
    Case Is = "Schedule"
    Cells(rw, 1) = ChrW(T_MrkSchedule)
    Case Else
    MsgBox "Error: Symbol 'Mrk" & MarkerType & "' not found.", vbOKOnly, _
        "Function: ApplyTraceMarker()"
    End Select
Next rw
End Sub

'==============================================================================
' Name:     FormatBorders
' Author:   PS
' Desc:     Makes boders to match the Trace Style
' Args:     None
' Comments: (1)
'==============================================================================
Sub FormatBorders()
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'==============================================================================
' Name:     Plot
' Author:   PS
' Desc:     Plots spectrum for each row, calls the form frmPlot, which makes
'           it look nice gives more formatting tools.
' Args:     None
' Comments: (1)
'==============================================================================
Sub Plot()

Dim StartRw As Integer
Dim endRw As Integer
Dim TraceChartObj As ChartObject
Dim XaxisTitle As String
Dim YaxisTitle As String
Dim TraceChartTitle As String
Dim SheetName As String
Dim SeriesNameStr As String
Dim SeriesNo As Integer
Dim DefaultWidth As Integer
Dim DefaultHeight As Integer

    'catch if chart is already selected, build chart if not
    If ActiveChart Is Nothing Then
    
        'check if sheet name contains space and needs quotation marks
        If Left(ActiveSheet.Name, 1) <> "'" And _
            Right(ActiveSheet.Name, 1) <> "'" Then
        SheetName = "'" & ActiveSheet.Name & "'!"
        Else
        SheetName = ActiveSheet.Name & "!"
        End If
    
    'set plot ranges
    StartRw = Selection.Row
    endRw = Selection.Row + Selection.Rows.Count - 1
    
        'set X-axis title
        Select Case T_BandType
        Case Is = "oct"
        XaxisTitle = "Octave Band Centre Frequency, Hz"
        Case Is = "to"
        XaxisTitle = "One-Third Octave Band Centre Frequency, Hz"
        Case Is = "cvt"
        XaxisTitle = "One-Third Octave Band Centre Frequency, Hz"
        End Select
        
        'check for A-weighting for Y-axis title
        If Right(T_SheetType, 1) = "A" Then
        YaxisTitle = "Sound Pressure Level, dBA"
        Else
        YaxisTitle = "Sound Pressure Level, dB"
        End If
    

    DefaultHeight = Application.CentimetersToPoints(14)
    DefaultWidth = Application.CentimetersToPoints(19)
    
        'create chart
    '    Left, Top,                Width, Height
    Set TraceChartObj = ActiveSheet.ChartObjects.Add _
        (600, Cells(StartRw, 1).Top + 5, DefaultWidth, DefaultHeight)
    TraceChartObj.Chart.ChartType = xlLine
    TraceChartObj.Placement = xlFreeFloating 'don't resize with cells
    TraceChartObj.ShapeRange.Line.Visible = msoFalse
    'add series
    SeriesNo = 1
        For plotrw = StartRw To endRw
            'set name and values
            With TraceChartObj.Chart.SeriesCollection.NewSeries
            .Name = "=" & SheetName & Cells(plotrw, 2).Address
            .values = Range(Cells(plotrw, T_LossGainStart), _
                            Cells(plotrw, T_LossGainEnd))
            End With
        'set x-axis values
        TraceChartObj.Chart.FullSeriesCollection(SeriesNo).XValues = _
            "=" & SheetName & Range(Cells(T_FreqRow, T_LossGainStart), _
            Cells(T_FreqRow, T_LossGainEnd)).Address
        SeriesNo = SeriesNo + 1
        Next plotrw
    DoEvents
        
        'format legend, axis labels etc
        With TraceChartObj.Chart
        
        'legend
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.size = 9
        .SetElement (msoElementPrimaryCategoryAxisTitleBelowAxis)
        .SetElement (msoElementPrimaryValueAxisTitleBelowAxis)
        
        'chart titles
'        .SetElement (msoElementChartTitleAboveChart)
'        .ChartTitle.Font.size = 12
        .SetElement (msoElementChartTitleNone)
        
        'axis
        .Axes(xlValue, xlPrimary).MajorUnit = 10
        .Axes(xlValue, xlPrimary).MinorUnit = 5
        .Axes(xlValue, xlPrimary).HasMinorGridlines = True
        .Axes(xlCategory, xlPrimary).AxisBetweenCategories = False
        'set 60dB range
        .Axes(xlValue, xlPrimary).MinimumScale = _
            .Axes(xlValue, xlPrimary).MaximumScale - 60
            
        
        'variable YaxisTitle is set earlier in the code
        .Axes(xlValue, xlPrimary).AxisTitle.Text = YaxisTitle
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = XaxisTitle
        End With
    
    'Call graph formatter
    TraceChartObj.Select
    End If

'launch the formatting tool!
frmPlotTool.Show

End Sub


'==============================================================================
' Name:     HeatMap
' Author:   PS
' Desc:     Applies conditional formatting for the spectrum
' Args:     None
' Comments: (1) Includes options for row-by-row formatting, or the entire range
'           (2) Calls GreenYellowRed. Might add other colour schemes later.
'           (3) Added optional input, which skips the check and applies to
'               whole group. Called by Options Analysis subroutine.
'==============================================================================
Sub HeatMap(Optional SkipCheck As Boolean)
Dim RowByRow As Boolean
Dim StartRw As Integer
Dim endRw As Integer
Dim SelectRw As Integer
Dim InitialSelection As String

InitialSelection = Selection.Address

StartRw = Selection.Row
endRw = StartRw + Selection.Rows.Count - 1

    If StartRw = endRw Then 'only one row
    RowByRow = False
    ElseIf SkipCheck = True Then
    RowByRow = False
    Else
    
    msg = MsgBox("Apply heat map row-by-row?", vbYesNo, _
        "I love a sunburnt country")
        
        If msg = vbYes Then
        RowByRow = True
        ElseIf msg = vbNo Then
        RowByRow = False
        Else 'just in case
        End
        End If
        
    End If


'clear any existing formatting
Range(Cells(StartRw, T_Description), Cells(endRw, T_LossGainEnd)).Select
Selection.FormatConditions.Delete
    
    If RowByRow = True Then
        For SelectRw = StartRw To endRw 'loop for each row
        'select one row
        Range(Cells(SelectRw, T_LossGainStart), _
            Cells(SelectRw, T_LossGainEnd)).Select
        'make-a-the-pretty-colours!
        GreenYellowRed
        Next SelectRw
    Else
    Range(Cells(StartRw, T_LossGainStart), _
            Cells(endRw, T_LossGainEnd)).Select
        GreenYellowRed
    End If
    
'go back to initially selected range
Range(InitialSelection).Select

End Sub

'==============================================================================
' Name:     ClearHeatMap
' Author:   PS
' Desc:     Deletes conditional formatting for all selected rows
' Args:     None
' Comments: (1)
'==============================================================================
Sub ClearHeatMap()
Dim StartRw As Integer
Dim endRw As Integer
StartRw = Selection.Row
endRw = StartRw + Selection.Rows.Count - 1
'remove heatmap
Range(Cells(StartRw, T_Description), Cells(endRw, T_LossGainEnd)).FormatConditions.Delete
End Sub

'==============================================================================
' Name:     GreenYellowRed
' Author:   PS
' Desc:     Applies formatting style for heat hap
' Args:     None
' Comments: (1)
'==============================================================================
Sub GreenYellowRed()
'add colour scale
Selection.FormatConditions.AddColorScale ColorScaleType:=3
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1)
    'green
    .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        With .ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
        End With
    
    'yellow
    .ColorScaleCriteria(2).Type = xlConditionValuePercentile
    .ColorScaleCriteria(2).Value = 50
        With .ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
        End With
        
    'red
    .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        With .ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
        End With
        
    End With
    
End Sub



