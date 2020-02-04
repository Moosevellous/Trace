Attribute VB_Name = "SheetTools"
Sub HeaderBlock(TypeCode As String)
    If TypeCode = "NR1L" Or TypeCode = "R2R" Or TypeCode = "RT" Or TypeCode = "N1L" Or TypeCode = "BA" Then
    msg = MsgBox("Header block not supported for Sheet Type '" & TypeCode & "'", vbOKOnly, "Error - header block") 'do nothing
    Else
    

    'Project No and Name
    ScanProjectInfoHTML
    GetProjectInfoHTML 'OLD GetProjectInfo
    Cells(1, 3).Value = PROJECTNO
    Cells(2, 3).Value = PROJECTNAME
    
    'Date
    Cells(1, 10).Value = Now
    'Engineer
    If ENGINEER = "" Then Update_ENGINEER
    Cells(2, 11).Value = ENGINEER
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
SplitStr = Split(StrUserName, " ", Len(StrUserName), vbTextCompare)
ENGINEER = Left(SplitStr(1), 1) & Left(SplitStr(0), 1)
End Sub

'Legacy code for text files
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

Application.StatusBar = "Opening file: " & PROJECTINFODIRECTORY

'header is line 0
ReDim Preserve ReadStr(0)
Line Input #1, ReadStr(0)
'Debug.Print ReadStr(0)
SplitHeader = Split(ReadStr(0), ";", Len(ReadStr(0)), vbTextCompare)
    For C = 0 To UBound(SplitHeader)
        If SplitHeader(C) = "Job number*" Then
        jobNoCol = C
        End If
        
        If SplitHeader(C) = "Job name*" Then
        jobNameCol = C
        End If
    Next C
'data is line 1
ReDim Preserve ReadStr(1)
Line Input #1, ReadStr(1)
'Debug.Print ReadStr(1)

splitData = Split(ReadStr(1), ";", Len(ReadStr(1)), vbTextCompare)
PROJECTNO = splitData(jobNoCol)
PROJECTNAME = splitData(jobNameCol)

closefile:
Close #1

Application.StatusBar = False

End Sub

Sub GetProjectInfoHTML()
Dim scanBookName As String
Dim MainBookName As String
MainBookName = ActiveWorkbook.Name
'status bar
Application.StatusBar = "Opening HTML file: " & PROJECTINFODIRECTORY
Application.ScreenUpdating = False
'open file
Workbooks.Open Filename:=PROJECTINFODIRECTORY
DoEvents
scanBookName = ActiveWorkbook.Name
'set global variables
PROJECTNO = Cells(3, 2).Value
PROJECTNAME = Cells(5, 2).Value
'close file
Workbooks(scanBookName).Close (False)
DoEvents
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub

'LEGACY CODE FOR SCANNING
Sub ScanProjectInfoDirectory()
Dim splitDir() As String
Dim searchPath As String
Dim testPath As String
Dim foundProjectDirectory As Boolean
Dim searchlevel As Integer
Dim checkExists As String

foundProjectDirectory = False
searchlevel = 0
    While foundProjectDirectory = False And searchlevel <= 4 'max 4 searchlevels
        testPath = ""
        splitDir = Split(ActiveWorkbook.Path, "\", Len(ActiveWorkbook.Path), vbTextCompare)
        
            For i = 0 To UBound(splitDir) - searchlevel
            testPath = testPath & "\" & splitDir(i)
            Next i
            
        If Len(testPath) = 0 Then End
        
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

Sub ScanProjectInfoHTML()
Dim splitDir() As String
Dim SplitPS() As String
Dim searchPath As String
Dim testPath As String
Dim HTMLFilePath As String
Dim foundProjectDirectory As Boolean
Dim searchlevel As Integer
Dim checkExists As String
Dim ProjNoExtract As String

foundProjectDirectory = False
searchlevel = 0

    While foundProjectDirectory = False And searchlevel <= 4 'max 4 searchlevels
        testPath = ""
        splitDir = Split(ActiveWorkbook.Path, "\", Len(ActiveWorkbook.Path), vbTextCompare)
        
            For i = 0 To UBound(splitDir) - searchlevel
                If i = 0 Then 'first element
                testPath = splitDir(i)
                Else
                testPath = testPath & "\" & splitDir(i)
                End If
            Next i
              
        If Len(testPath) = 0 Then End 'path not found
        SplitPS = Split(testPath, "PS", Len(testPath), vbTextCompare)
        ProjNoExtract = "PS" & Left(SplitPS(1), 6)
        HTMLFilePath = Right(testPath, Len(testPath)) & "\*" & ProjNoExtract & "*.html"
        'Debug.Print testPath
        checkExists = Dir(HTMLFilePath)
        
            If Len(checkExists) > 0 Then
            foundProjectDirectory = True
            PROJECTINFODIRECTORY = testPath & "\" & checkExists

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
Dim StartRw As Integer
Dim EndRw As Integer
Dim StartCol As Integer
Dim EndCol As Integer
Dim TraceChartObj As ChartObject
Dim XaxisTitle As String
Dim YaxisTitle As String

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(TypeCode, 3) = "OCT" Then
    StartRw = Selection.Row
    EndRw = Selection.Row + Selection.Rows.count - 1
    StartCol = 5
    EndCol = 13
    OneThirdsCheck = False
    XaxisTitle = "Octave Band Centre Frequency, Hz"
    ElseIf Left(TypeCode, 2) = "TO" Then
    StartCol = 5
    EndCol = 25
    OneThirdsCheck = True
    XaxisTitle = "One-Third Octave Band Centre Frequency, Hz"
    End If
    
    'check for A-weighting
    If Right(TypeCode, 1) = "A" Then
    YaxisTitle = "Sound Pressure Level, dBA"
    Else
    YaxisTitle = "Sound Pressure Level, dB"
    End If
    
TraceChartTitle = InputBox("Name of the chart?", "Nom de le carte", Cells(Selection.Row, 2).Value)

'create chart
Set TraceChartObj = ActiveSheet.ChartObjects.Add(600, 70, 340, 400) 'L, T, W, H
TraceChartObj.Chart.ChartType = xlLine

    'add series
    For plotrw = StartRw To EndRw
    'TraceChartObj.Chart.Add '(Range(Cells(plotrw, 5), Cells(plotrw, 13)))
        With TraceChartObj.Chart.SeriesCollection.NewSeries
        .Name = Cells(plotrw, 2)
        .Values = Range(Cells(plotrw, StartCol), Cells(plotrw, EndCol))
        End With
    Next plotrw
    
    With TraceChartObj.Chart
    .Legend.Position = xlLegendPositionBottom
    .Legend.Font.size = 9
    .SetElement (msoElementPrimaryCategoryAxisTitleBelowAxis)
    .SetElement (msoElementPrimaryValueAxisTitleBelowAxis)
    .SetElement (msoElementChartTitleAboveChart)
    .ChartTitle.Text = TraceChartTitle
    .ChartTitle.Font.size = 12
    .Axes(xlValue, xlPrimary).MajorUnit = 10
    .Axes(xlCategory, xlPrimary).AxisBetweenCategories = False
    .Axes(xlValue, xlPrimary).AxisTitle.Text = YaxisTitle
    .Axes(xlCategory, xlPrimary).AxisTitle.Text = XaxisTitle
    End With

    'x-axis labels
    If OneThirdsCheck = True Then 'One third octave
    TraceChartObj.Chart.FullSeriesCollection(1).XValues = "=" & ActiveSheet.Name & "!$E$6:$Y$6"
    Else 'Octave
    TraceChartObj.Chart.FullSeriesCollection(1).XValues = "=" & ActiveSheet.Name & "!$E$6:$M$6"
    End If

'Call graph formatter
TraceChartObj.Select
frmChartFormatter.Show

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
LastRw = StartRw + Selection.Rows.count - 1

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
Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
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

Dim AposPos As Integer 'Position of the apostrophe' in the string
Dim ExPos As Integer 'Position of the exclamation mark! in the string

'find exclamation mark character
InputFormula = Selection.Formula
ExPos = InStr(1, InputFormula, "!", vbTextCompare)
AposPos = InStr(1, InputFormula, "'", vbTextCompare)

'catch error
If AposPos = 0 Or ExPos = 0 Then End

PurgeStr = Mid(InputFormula, AposPos, ExPos - AposPos + 1)
'Debug.Print PurgeStr
Cells.Replace What:=PurgeStr, Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
msg = MsgBox("Reference string " & Chr(10) & PurgeStr & Chr(10) & "has been removed.", vbOKOnly, "THE PURGE!")
End Sub
