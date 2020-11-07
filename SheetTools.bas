Attribute VB_Name = "SheetTools"
'==============================================================================
' Name:     FillHeaderBlock
' Author:   PS
' Desc:     Puts project info in the header block
' Args:     None
' Comments: (1) Requires standardised header layout
'           (2) Calls ScanProjectInfoHTML and stores info in public variables
'==============================================================================
Sub FillHeaderBlock()

    If T_SheetType = "NR1L" Or T_SheetType = "R2R" Or T_SheetType = "N1L" Or _
        T_SheetType = "BA" Then '<-TODO make this more flexible
    msg = MsgBox("Header block not supported for Sheet Type '" & _
        T_SheetType & "'", vbOKOnly, "Error - header block")
    End
    End If

'Project No and Name
GetProjectInfoHTML
Cells(1, 3).Value = PROJECTNO
Cells(2, 3).Value = PROJECTNAME
'Date
Cells(1, 10).Value = Now
'Engineer
    If ENGINEER = "" Then Update_ENGINEER
Cells(2, 11).Value = ENGINEER

End Sub

'==============================================================================
' Name:     ClearHeaderBlock
' Author:   PS
' Desc:     Clear project info in the header block
' Args:     None
' Comments: (1) Requires standardised header layout
'==============================================================================
Sub ClearHeaderBlock()

    If T_SheetType = "NR1L" Or T_SheetType = "R2R" Or T_SheetType = "N1L" Or _
        T_SheetType = "BA" Then '<-TODO make this more flexible
    msg = MsgBox("Header block not supported for Sheet Type '" & _
        T_SheetType & "'", vbOKOnly, "Error - header block")
    End
    End If
    
'user confirmation
msg = MsgBox("Are you sure?", vbYesNo, "Choose wisely...")

    If msg = vbYes Then
    Range("C1:H1").ClearContents
    Range("C2:H2").ClearContents
    Range("C3:H3").ClearContents
    Range("J1:M1").ClearContents
    Range("K2:M2").ClearContents
    Range("K3:M3").ClearContents
    End If

End Sub

'==============================================================================
' Name:     Update_ENGINEER
' Author:   PS
' Desc:     Gets engineer initials from Windows
' Args:     None
' Comments: (1)
'==============================================================================
Sub Update_ENGINEER()
Dim StrUserName As String
StrUserName = Application.UserName
SplitStr = Split(StrUserName, " ", Len(StrUserName), vbTextCompare)
ENGINEER = Left(SplitStr(1), 1) & Left(SplitStr(0), 1)
End Sub

'==============================================================================
' Name:     GetProjectInfoHTML
' Author:   PS
' Desc:     Gets project information from HTML file, created by Project system
' Args:     None
' Comments: (1) Calls ScanProjectInfoHTML
'==============================================================================
Sub GetProjectInfoHTML()
Dim scanBookName As String
Dim MainBookName As String

'find HTML file
ScanProjectInfoHTML

'set workbook
MainBookName = ActiveWorkbook.Name

'status bar
Application.StatusBar = "Opening HTML file: " & PROJECTINFODIRECTORY
Application.ScreenUpdating = False

'open file
    If PROJECTINFODIRECTORY <> "" Then
    Workbooks.Open fileName:=PROJECTINFODIRECTORY
    DoEvents
    scanBookName = ActiveWorkbook.Name
    'set public variables
    PROJECTNO = Cells(3, 2).Value
    PROJECTNAME = Cells(5, 2).Value
    'close file
    Workbooks(scanBookName).Close (False)
    End If
DoEvents

'status bar updates
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub

'==============================================================================
' Name:     ScanProjectInfoHTML
' Author:   PS
' Desc:     Looks for HTML file, created when project was created by the system
' Args:     None
' Comments: (1)
'==============================================================================
Sub ScanProjectInfoHTML()
Dim splitDir() As String
Dim SplitPS() As String 'project code starts with PS
Dim searchPath As String
Dim testPath As String
Dim HTMLFilePath As String
Dim foundProjectDirectory As Boolean
Dim searchLevel As Integer
Dim checkExists As String
Dim ProjNoExtract As String
Dim elem As Integer
Dim MaxSearchLevels As Integer

foundProjectDirectory = False
searchLevel = 0
MaxSearchLevels = 10

    While foundProjectDirectory = False And searchLevel <= MaxSearchLevels
        testPath = ""
        splitDir = Split(ActiveWorkbook.Path, "\", Len(ActiveWorkbook.Path), _
            vbTextCompare)
        
            For i = 0 To UBound(splitDir) - searchLevel
                If i = 0 Then 'first element
                testPath = splitDir(i)
                Else
                testPath = testPath & "\" & splitDir(i)
                End If
            Next i

            If Len(testPath) = 0 Or InStr(1, testPath, "https://", _
                vbTextCompare) > 0 Then
            'skip! sharepoint location not allowed, nor are blank file paths
            PROJECTINFODIRECTORY = ""
            Else
            
            SplitPS = Split(testPath, "PS", Len(testPath), vbTextCompare)
                If UBound(SplitPS) > 0 Then
                    'for projects in the new ProjectsAU folders,
                    'e.g. U:\ProjectsAU\PS117xxx
                    If InStr(1, testPath, "xxx", vbTextCompare) > 0 Then
                    elem = 2
                    Else 'old project structure
                    elem = 1
                    End If
                
                'use the letters PS to get the file path
                ProjNoExtract = "PS" & Left(SplitPS(elem), 6)
                HTMLFilePath = Right(testPath, Len(testPath)) & "\*" & _
                    ProjNoExtract & "*.html"
                Debug.Print "Checking path: " & HTMLFilePath
    
                Application.StatusBar = "Scanning: " & testPath
                checkExists = Dir(HTMLFilePath)
                
                    'if HTML file was found, stores in public Variable
                    If Len(checkExists) > 0 Then
                    Application.StatusBar = "Project HTML file found!"
                    foundProjectDirectory = True
                    PROJECTINFODIRECTORY = testPath & "\" & checkExists
                    Debug.Print "****PATH FOUND****"
                    End If
                End If
            End If
        searchLevel = searchLevel + 1
    Wend
'status bar
Application.StatusBar = False
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
Dim EndRw As Integer
Dim TraceChartObj As ChartObject
Dim XaxisTitle As String
Dim YaxisTitle As String
Dim TraceChartTitle As String
Dim SheetName As String
Dim SeriesNameStr As String
Dim SeriesNo As Integer

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
    EndRw = Selection.Row + Selection.Rows.Count - 1
    
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
    
    'create chart
    '                                              Left, Top, Width, Height
    Set TraceChartObj = ActiveSheet.ChartObjects.Add(600, 70, 340, 400)
    TraceChartObj.Chart.ChartType = xlLine
    
    'add series
    SeriesNo = 1
        For plotrw = StartRw To EndRw
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
        .Legend.Position = xlLegendPositionBottom
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
        '<--------------------------------------------------TODO: set 60dB range?
        '.Axes(xlValue, xlPrimary).MinimumScale = _
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
'==============================================================================
Sub HeatMap()
Dim RowByRow As Boolean
Dim StartRw As Integer
Dim EndRw As Integer

StartRw = Selection.Row
EndRw = StartRw + Selection.Rows.Count - 1

    If StartRw = EndRw Then 'only one row
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
Range(Cells(StartRw, T_Description), Cells(EndRw, T_LossGainEnd)).Select
Selection.FormatConditions.Delete
    
    If RowByRow = True Then
        For selectrw = StartRw To EndRw 'loop for each row
        'select one row
        Range(Cells(selectrw, T_LossGainStart), _
            Cells(selectrw, T_LossGainEnd)).Select
        'make-a-the-pretty-colours!
        GreenYellowRed
        Next selectrw
    Else
        GreenYellowRed
    End If

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


'==============================================================================
' Name:     FixReferences
' Author:   PS
' Desc:     Fixes formula references to other copies of Trace
' Args:     Mode - String set by Ribbon control to loop through all sheets
'           within a workbook
' Comments: (1)
'==============================================================================
Sub FixReferences(Mode As String)

Dim AposPos As Integer 'Position of the apostrophe' in the string
Dim ExPos As Integer 'Position of the exclamation mark! in the string
Dim inputFormula As String
Dim inputFormulaVar As Variant
Dim PurgeStr As String 'string to be deleted

    'find exclamation mark character
    If TypeName(Selection.Formula) = "Variant()" Then 'catch merged cells
    inputFormulaVar = Selection.Formula
    inputFormula = inputFormulaVar(1, 1)
    Else
    inputFormula = Selection.Formula
    End If

'location position within the inputFormula
ExPos = InStr(1, inputFormula, "!", vbTextCompare)
AposPos = InStr(1, inputFormula, "'", vbTextCompare)

    'catch error with file paths with no spaces, use
    If AposPos = 0 Then
    AposPos = InStr(1, inputFormula, "\\", vbTextCompare)
    End If
    
    'catch error if not found
    If AposPos = 0 Or ExPos = 0 Then
    msg = MsgBox("Reference not found!" & chr(10) & _
        "Try selecting a cell with the reference to be fixed and try again.", _
        vbOKOnly, "Search Error")
    End
    End If

PurgeStr = Mid(inputFormula, AposPos, ExPos - AposPos + 1)
'Debug.Print "Purging: " PurgeStr

    'if all sheets, then loop through
    If Mode = "FixReferencesAll" Or Mode = "FixReferencesDefault" Then
    ReplaceFormulaReferences PurgeStr, "", False
    Else 'current sheet only, no loops
    ReplaceFormulaReferences PurgeStr, "", True
    End If

'Fix Legacy Functions
FixLegacyFunctions

End Sub


'==============================================================================
' Name:     FixLegacyFunctions
' Author:   PS
' Desc:     Replaces old, legacy formulas with shiny new ones
' Args:     None
' Comments: (1) Use to fix legacy functions etc
'==============================================================================
Sub FixLegacyFunctions()

Application.ScreenUpdating = False

'MECH MODULE
ReplaceFormulaReferences "GetASHRAEDuct", "DuctAtten_ASHRAE", False

ReplaceFormulaReferences "GetASHRAEPlenumLoss", "PlenumLoss_ASHRAE", False

ReplaceFormulaReferences "GetASHRAEPlenumLoss_OneThirdOctave", _
    "PlenumLossOneThirdOctave_ASHRAE", False
    
ReplaceFormulaReferences "GetDuctBreakIn", "DuctBreakIn_NEBB", False

ReplaceFormulaReferences "GetDuctBreakout", "DuctBreakOut_NEBB", False

ReplaceFormulaReferences "GetDuctDirectivity", "DuctDirectivity_PGD", False

ReplaceFormulaReferences "GetElbowLoss", "ElbowLoss_ASHRAE", False

ReplaceFormulaReferences "GetElbowLossASHRAE", "ElbowLoss_ASHRAE", False

ReplaceFormulaReferences "GetElbowLossNEBB", "ElbowLoss_NEBB", False

ReplaceFormulaReferences "GetERL_ASHRAE", "ERL_ASHRAE", False

ReplaceFormulaReferences "GetERL_NEBB", "ERL_NEBB", False

ReplaceFormulaReferences "GetFlexDuct", "FlexDuctAtten_ASHRAE", False

ReplaceFormulaReferences "GetRegenNoise_ASHRAE", "RegenNoise_ASHRAE", False

ReplaceFormulaReferences "GetReynoldsDuct", "DuctAtten_Reynolds", False

ReplaceFormulaReferences "GetReynoldsDuctCircular", _
    "DuctAttenCircular_Reynolds", False

'NOISE MODULE
ReplaceFormulaReferences "GetRoomLoss", "RoomLossTypical", False

ReplaceFormulaReferences "GetRoomLossRT", "RoomLossTypicalRT", False

'BASICS MODULE
ReplaceFormulaReferences "GetSpeedOfSound", "SpeedOfSound", False

ReplaceFormulaReferences "GetWavelength", "Wavelength", False

Application.ScreenUpdating = True

End Sub

'==============================================================================
' Name:     ReplaceFormulaReferences
' Author:   PS
' Desc:     Replaces all references in formulas
' Args:     FindStr - The part of the formula to be found
'           ReplaceStr - The new formula
'           ThisSheetOnly - set to TRUE for the current sheet only, othewise
'           code will loop through all sheets
' Comments: (1)
'==============================================================================
Sub ReplaceFormulaReferences(FindStr As String, ReplaceStr As String, _
Optional ThisSheetOnly As Boolean)

Application.StatusBar = "REPLACING: " & FindStr & "      WITH: " & ReplaceStr

Dim sh As Integer
Dim ReturnSheet As String 'sheet to return to when it's all done

ReturnSheet = ActiveSheet.Name

    If ThisSheetOnly = True Then 'current sheet only, no loops
    Cells.Replace What:=FindStr, Replacement:=ReplaceStr, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
    Else
        For sh = 1 To ActiveWorkbook.Sheets.Count
            If Sheets(sh).Type = xlWorksheet Then 'not for chart sheet types
            Sheets(sh).Activate
            Cells.Replace What:=FindStr, Replacement:=ReplaceStr, LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, _
                SearchFormat:=False, ReplaceFormat:=False
            End If
        Next sh
    End If
    
'go back to sheet where you started
Sheets(ReturnSheet).Activate

Application.StatusBar = False

End Sub













'**************
'Code Graveyard
'**************

''LEGACY CODE FOR SCANNING
'Sub ScanProjectInfoDirectory()
'Dim splitDir() As String
'Dim searchPath As String
'Dim testPath As String
'Dim foundProjectDirectory As Boolean
'Dim searchLevel As Integer
'Dim checkExists As String
'
'foundProjectDirectory = False
'searchLevel = 0
'    While foundProjectDirectory = False And searchLevel <= 4 'max 4 searchlevels
'        testPath = ""
'        splitDir = Split(ActiveWorkbook.path, "\", Len(ActiveWorkbook.path), vbTextCompare)
'
'            For i = 0 To UBound(splitDir) - searchLevel
'            testPath = testPath & "\" & splitDir(i)
'            Next i
'
'        If Len(testPath) = 0 Then End
'
'        testPath = Right(testPath, Len(testPath) - 1) & "\" & "ProjectInfo.txt"
'        'Debug.Print testPath
'        checkExists = Dir(testPath)
'
'            If Len(checkExists) > 0 Then
'            foundProjectDirectory = True
'            PROJECTINFODIRECTORY = testPath
'            End If
'        searchLevel = searchLevel + 1
'    Wend
'End Sub


''Legacy code for text files
'Sub GetProjectInfo()
'
'On Error GoTo closefile
'Dim ReadStr() As String
'Dim SplitHeader() As String
'Dim splitData() As String
'Dim jobNoCol As Integer
'Dim jobNameCol As Integer
'
''Get Array from text
'Close #1
'
'Open PROJECTINFODIRECTORY For Input As #1  'public
'
'Application.StatusBar = "Opening file: " & PROJECTINFODIRECTORY
'
''header is line 0
'ReDim Preserve ReadStr(0)
'Line Input #1, ReadStr(0)
''Debug.Print ReadStr(0)
'SplitHeader = Split(ReadStr(0), ";", Len(ReadStr(0)), vbTextCompare)
'    For C = 0 To UBound(SplitHeader)
'        If SplitHeader(C) = "Job number*" Then
'        jobNoCol = C
'        End If
'
'        If SplitHeader(C) = "Job name*" Then
'        jobNameCol = C
'        End If
'    Next C
''data is line 1
'ReDim Preserve ReadStr(1)
'Line Input #1, ReadStr(1)
''Debug.Print ReadStr(1)
'
'splitData = Split(ReadStr(1), ";", Len(ReadStr(1)), vbTextCompare)
'PROJECTNO = splitData(jobNoCol)
'PROJECTNAME = splitData(jobNameCol)
'
'closefile:
'Close #1
'
'Application.StatusBar = False
'
'End Sub
