Attribute VB_Name = "SheetFunctions"
Sub HeaderBlock(TypeCode As String)
    If TypeCode = "NR1L" Or TypeCode = "R2R" Or TypeCode = "RT" Or TypeCode = "N1L" Then
    'do nothing
    Else
    
    'Date
    Cells(1, 10).Value = Now
    'Engineer
        If ENGINEER = "" Then Update_ENGINEER
    Cells(2, 11).Value = ENGINEER
    
    ScanProjectInfoDirectory
    
    'Project No and Name
    GetProjectInfo
    Cells(1, 3).Value = PROJECTNO
    Cells(2, 3).Value = PROJECTNAME
    End If
    
End Sub


Sub ClearHeaderBlock(TypeCode As String)
    If TypeCode = "NR1L" Or TypeCode = "R2R" Or TypeCode = "RT" Or TypeCode = "N1L" Then
    'do nothing
    Else
    msg = MsgBox("Are you sure?", vbYesNo, "Choose wisely...")
        If msg = vbYes Then
        Cells(1, 3).Value = ""
        Cells(2, 3).Value = ""
        Cells(3, 3).Value = ""
        Cells(1, 10).Value = ""
        Cells(2, 11).Value = ""
        End If
    End If
End Sub


Sub Update_ENGINEER()
Dim StrUserName As String
StrUserName = Application.UserName
splitStr = Split(StrUserName, " ", Len(StrUserName), vbTextCompare)
ENGINEER = Left(splitStr(1), 1) & Left(splitStr(0), 1)
End Sub


Sub GetProjectInfo()

On Error GoTo closefile
Dim ReadStr() As String
Dim SplitHeader() As String
Dim splitData() As String
Dim jobNoCol As Integer
Dim jobNameCol As Integer

'Get Array from text
Close #1

Open PROJECTINFODIRECTORY For Input As #1  'global

'header is line 0
ReDim Preserve ReadStr(0)
Line Input #1, ReadStr(0)
'Debug.Print ReadStr(0)
SplitHeader = Split(ReadStr(0), ";", Len(ReadStr(0)), vbTextCompare)
    For c = 0 To UBound(SplitHeader)
        If SplitHeader(c) = "Job number*" Then
        jobNoCol = c
        End If
        
        If SplitHeader(c) = "Job name*" Then
        jobNameCol = c
        End If
    Next c
'data is line 1
ReDim Preserve ReadStr(1)
Line Input #1, ReadStr(1)
'Debug.Print ReadStr(1)

splitData = Split(ReadStr(1), ";", Len(ReadStr(1)), vbTextCompare)
PROJECTNO = splitData(jobNoCol)
PROJECTNAME = splitData(jobNameCol)

closefile:
Close #1
End Sub


Sub ScanProjectInfoDirectory()
Dim splitDir() As String
Dim searchPath As String
Dim testPath As String
Dim foundProjectDirectory As Boolean
Dim searchlevel As Integer
Dim checkExists As String

foundProjectDirectory = False
searchlevel = 0
    While foundProjectDirectory = False And searchlevel <= 3 'max 3 searchlevels
        testPath = ""
        splitDir = Split(ActiveWorkbook.Path, "\", Len(ActiveWorkbook.Path), vbTextCompare)
        
            For i = 0 To UBound(splitDir) - searchlevel
            testPath = testPath & "\" & splitDir(i)
            Next i
        
        testPath = Right(testPath, Len(testPath) - 1) & "\" & "ProjectInfo.txt"
        'Debug.Print testPath
        checkExists = Dir(testPath)
        
            If Len(checkExists) > 0 Then
            foundProjectDirectory = True
            PROJECTINFODIRECTORY = testPath
            End If
        searchlevel = searchlevel + 1
    Wend
End Sub


Sub FormatBorders()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub Plot(TypeCode As String)

Dim OneThirdsCheck As Boolean
Dim CheckAWeight As Boolean

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(TypeCode, 3) = "OCT" Then
    StartRw = Selection.Row
    endrw = Selection.Row + Selection.Rows.Count - 1
    Range(Cells(StartRw, 4), Cells(endrw, 13)).Select
    ElseIf Left(TypeCode, 2) = "TO" Then
    'do nothing
    End If
    
    'check for A-weighting
    If Right(TypeCode, 1) = "A" Then
    CheckAWeight = True
    Else
    CheckAWeight = False
    End If


    'check for one thirds
    If Left(TypeCode, 2) = "TO" Then
    OneThirdsCheck = True
    Else
    OneThirdsCheck = False
    End If

AxisTitle = InputBox("Name of the Chart?", "Chart Title", Cells(Selection.Row, 2).Value)

SeriesRange = Selection.Address

'create chart
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=Range(SeriesRange)
ActiveChart.ChartType = xlLine
ChartName = ActiveChart.Name
    
        With ActiveChart
        .Legend.Delete
        .Axes(xlValue).TickLabels.NumberFormat = "0"
        .Axes(xlCategory).AxisBetweenCategories = False
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = AxisTitle
        .ChartTitle.Font.size = 12
        .ChartArea.Width = 380
        .ChartArea.Height = 470
            If CheckAWeight = True Then
             .SetElement (msoElementPrimaryValueAxisTitleRotated)
             .Axes(xlValue, xlPrimary).AxisTitle.Text = "Sound Pressure Level, dBA"
            ElseIf CheckAWeight = False Then
             .SetElement (msoElementPrimaryValueAxisTitleRotated)
             .Axes(xlValue, xlPrimary).AxisTitle.Text = "Sound Pressure Level, dB"
            End If
            
            If OneThirdsCheck = True Then
             .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
             .Axes(xlCategory, xlPrimary).AxisTitle.Text = "1/3 Octave Band Centre Frequency, Hz"
            ElseIf OneThirdsCheck = False Then
             .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
             .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Octave Band Centre Frequency, Hz"
            End If
        .Axes(xlValue).MinimumScale = 0
        End With

    'x-axis labels
    If OneThirdsCheck = True Then 'One third octave
    ActiveChart.FullSeriesCollection(1).XValues = "=" & ActiveSheet.Name & "!$D$6:$Y$6"
    Else 'Octave
    ActiveChart.FullSeriesCollection(1).XValues = "=" & ActiveSheet.Name & "!$D$6:$M$6"
    End If

NumSeries = ActiveChart.SeriesCollection.Count
    For SeriesX = 1 To NumSeries
        
        'format series
        ActiveChart.SeriesCollection(SeriesX).Select
                With Selection
                .MarkerStyle = 1
                .MarkerSize = 3
                .Border.Weight = 2
                End With
    
        'format first/final point
            NumPoints = ActiveChart.SeriesCollection(SeriesX).Points.Count
            ActiveChart.SeriesCollection(SeriesX).Points(1).Select
                With Selection
                .MarkerStyle = 1
                .MarkerSize = 3
                .HasDataLabel = True
                End With
            ActiveChart.SeriesCollection(SeriesX).Points(2).Select
                With Selection
                .Border.LineStyle = xlNone
                End With
    
        
    Next SeriesX

End Sub

Sub HeatMap(SheetType As String)
Dim RowByRow As Boolean

msg = MsgBox("Apply heat map row-by-row?", vbYesNoCancel, "I love a sunburnt country")
If msg = vbCancel Then End

    If msg = vbYes Then
    RowByRow = True
    ElseIf msg = vbNo Then
    RowByRow = False
    Else
    msg = MsgBox("Prompt not recognised. Macro aborted.", vbOKOnly, "YOU SUCK")
    End If


StartRw = Selection.Row
LastRw = StartRw + Selection.Rows.Count - 1

    If Left(SheetType, 3) = "OCT" Then ' OCT or OCTA
    Range(Cells(StartRw, 3), Cells(LastRw, 13)).Select
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Range(Cells(StartRw, 3), Cells(LastRw, 25)).Select
    End If
Selection.FormatConditions.Delete
    
    If RowByRow Then
        For selectrw = StartRw To LastRw
            If Left(SheetType, 3) = "OCT" Then ' OCT or OCTA
            Range(Cells(StartRw, 3), Cells(LastRw, 13)).Select
            ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
            Range(Cells(StartRw, 3), Cells(LastRw, 25)).Select
            End If
        GreenYellowRed
        Next selectrw
    Else
        GreenYellowRed
    End If

End Sub

Sub GreenYellowRed()

Selection.FormatConditions.AddColorScale ColorScaleType:=3
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
xlConditionValueLowestValue
With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
.Color = 8109667
.TintAndShade = 0
End With

Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
xlConditionValuePercentile
Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
.Color = 8711167
.TintAndShade = 0
End With

Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
xlConditionValueHighestValue
With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
.Color = 7039480
.TintAndShade = 0
End With
End Sub

Sub FixReferences(SheetType As String)
'find exclamation mark character
InputFormula = Selection.Formula
ExPos = InStr(1, InputFormula, "!", vbTextCompare)
AposPos = InStr(1, InputFormula, "'", vbTextCompare)
PurgeStr = Mid(InputFormula, AposPos, ExPos - AposPos + 1)
'Debug.Print PurgeStr
Cells.Replace What:=PurgeStr, Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
msg = MsgBox("Reference string " & Chr(10) & PurgeStr & Chr(10) & "has been removed.", vbOKOnly, "THE PURGE!")
End Sub
