Attribute VB_Name = "SheetTools"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================

'Options Anlaysis
Public RngVar1 As String
Public RngVar2 As String
Public TargetRng As String

Public ResultRng As String
Public CreateHeading As Boolean


'==============================================================================
' Name:     TrimSheetName
' Author:   PS
' Desc:     Trims leading and trailing characters from sheet name
' Args:     InputStr - name of sheet from RefEdit box
' Comments: (1)
'==============================================================================
Function TrimSheetName(inputStr As String) As String
    If inputStr <> "" Then
        If Left(inputStr, 1) = "'" Then 'trim apostrophes and !
        TrimSheetName = Mid(inputStr, 2, Len(inputStr) - 3)
        Else 'trim "!"
        TrimSheetName = Left(inputStr, Len(inputStr) - 1)
        End If
    End If
End Function

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
'Description
Cells(3, 3).Value = _
    "=MID(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))+1,255)"

End Sub

'==============================================================================
' Name:     FillHeaderBlockAll
' Author:   PS
' Desc:     Loops through all sheets and fills in headers
' Args:     None
' Comments: (1)
'==============================================================================
Sub FillHeaderBlockAll()

Dim StartSheet As String

    If T_SheetType = "NR1L" Or T_SheetType = "R2R" Or T_SheetType = "N1L" Or _
        T_SheetType = "BA" Then '<-TODO make this more flexible
    msg = MsgBox("Header block not supported for Sheet Type '" & _
        T_SheetType & "'", vbOKOnly, "Error - header block")
    End
    End If

StartSheet = ActiveSheet.Name

'Project No and Name
GetProjectInfoHTML

    For sh = 1 To ActiveWorkbook.Sheets.Count

    Sheets(sh).Activate
        
        'check for typecode, which means there's a header block
        If NamedRangeExists("TYPECODE", True) Then
        'todo: check for header block in sheet?

        Cells(1, 3).Value = PROJECTNO
        Cells(2, 3).Value = PROJECTNAME
        'Date
        Cells(1, 10).Value = Now
        'Engineer
            If ENGINEER = "" Then Update_ENGINEER
        Cells(2, 11).Value = ENGINEER
        'Description
        Cells(3, 3).Value = _
            "=MID(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))+1,255)"
        End If

    Next sh

'go back to where you started
Sheets(StartSheet).Activate

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
StrUserName = Application.userName
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

On Error GoTo errCatch

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
            
            'old paths split on the characters PS e.g. \\corp.pbwan.net\ANZ\Projects\PS116962_State_Basketball\4_WIP\Doc_Disc\AC_Acoustics
            SplitPS = Split(testPath, "PS", Len(testPath), vbTextCompare)
            'new BD planner paths are different e.g. U:\ProjectsAU\200xxx\200095_Gilbert_and_Tobin_a
            SplitBDP = Split(testPath, "xxx\", Len(testPath), vbTextCompare)
            
            'Check which file path type it is
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
                'Debug.Print "Checking path: " & HTMLFilePath
    
                Application.StatusBar = "Scanning: " & testPath
                checkExists = Dir(HTMLFilePath)
                
                    'if HTML file was found, stores in public Variable
                    If Len(checkExists) > 0 Then
                    Application.StatusBar = "Project HTML file found!"
                    foundProjectDirectory = True
                    PROJECTINFODIRECTORY = testPath & "\" & checkExists
                    'Debug.Print "****PATH FOUND****"
                    End If
                    
                ElseIf UBound(SplitBDP) > 0 Then
                ProjNoExtract = Left(SplitBDP(1), 6)
                HTMLFilePath = Right(testPath, Len(testPath)) & "\*" & _
                    ProjNoExtract & "*.html"
                Application.StatusBar = "Scanning: " & testPath
                Debug.Print HTMLFilePath
                checkExists = Dir(HTMLFilePath)
                
                    'if HTML file was found, stores in public Variable
                    If Len(checkExists) > 0 Then
                    Application.StatusBar = "Project HTML file found!"
                    foundProjectDirectory = True
                    PROJECTINFODIRECTORY = testPath & "\" & checkExists
                    'Debug.Print "****PATH FOUND****"
                    End If
                    
                End If
            End If
        searchLevel = searchLevel + 1
    Wend

Application.StatusBar = False
Exit Sub
errCatch:
    If Err.Number = 52 Then
    msg = MsgBox("HTML file not found", vbOKOnly, _
        "Project Info Error")
    Else
    msg = MsgBox("Error " & Err.Number & chr(10) & Err.Description, vbOKOnly, _
        "Project Info Error")
    End If
Application.StatusBar = False
End Sub

'==============================================================================
' Name:     FixReferences
' Author:   PS
' Desc:     Fixes formula references to other copies of Trace
' Args:     Mode - String set by Ribbon control to loop through all sheets
'           within a workbook
' Comments: (1) Changed to make default run on this sheet only, and not replace legacy functions
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
    If Mode = "FixReferencesAll" Then
    ReplaceFormulaReferences PurgeStr, "", False
    'Fix Legacy Functions
    FixLegacyFunctions False
    ElseIf Mode = "FixReferencesDefault" Then
    ReplaceFormulaReferences PurgeStr, "", True
    Else 'current sheet only, no loops, fix
    ReplaceFormulaReferences PurgeStr, "", True
    FixLegacyFunctions True
    End If



End Sub


'==============================================================================
' Name:     FixLegacyFunctions
' Author:   PS
' Desc:     Replaces old, legacy formulas with shiny new ones
' Args:     None
' Comments: (1) Use to fix legacy functions etc
'==============================================================================
Sub FixLegacyFunctions(Optional ThisSheetOnly As Boolean)

Application.ScreenUpdating = False

'MECH MODULE
ReplaceFormulaReferences "GetASHRAEDuct", "DuctAtten_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetASHRAEPlenumLoss", "PlenumLoss_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetASHRAEPlenumLoss_OneThirdOctave", _
    "PlenumLossOneThirdOctave_ASHRAE", ThisSheetOnly
    
ReplaceFormulaReferences "GetDuctBreakIn", "DuctBreakIn_NEBB", ThisSheetOnly

ReplaceFormulaReferences "GetDuctBreakout", "DuctBreakOut_NEBB", ThisSheetOnly

ReplaceFormulaReferences "GetDuctDirectivity", "DuctDirectivity_PGD", ThisSheetOnly

ReplaceFormulaReferences "GetElbowLoss", "ElbowLoss_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetElbowLossASHRAE", "ElbowLoss_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetElbowLossNEBB", "ElbowLoss_NEBB", ThisSheetOnly

ReplaceFormulaReferences "GetERL", "ERL_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetERL_ASHRAE", "ERL_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetERL_NEBB", "ERL_NEBB", ThisSheetOnly

ReplaceFormulaReferences "GetFlexDuct", "FlexDuctAtten_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetRegenNoise_ASHRAE", "RegenNoise_ASHRAE", ThisSheetOnly

ReplaceFormulaReferences "GetReynoldsDuct", "DuctAtten_Reynolds", ThisSheetOnly

ReplaceFormulaReferences "GetReynoldsDuctCircular", _
    "DuctAttenCircular_Reynolds", ThisSheetOnly

'NOISE MODULE
ReplaceFormulaReferences "GetRoomLoss", "RoomLossTypical", ThisSheetOnly

ReplaceFormulaReferences "GetRoomLossRT", "RoomLossTypicalRT", ThisSheetOnly

'BASICS MODULE
ReplaceFormulaReferences "GetSpeedOfSound", "SpeedOfSound", ThisSheetOnly

ReplaceFormulaReferences "GetWavelength", "Wavelength", ThisSheetOnly

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
' Comments: (1) Will now turn off calculation during update
'==============================================================================
Sub ReplaceFormulaReferences(FindStr As String, ReplaceStr As String, _
Optional ThisSheetOnly As Boolean)

Application.StatusBar = "REPLACING: " & FindStr & "      WITH: " & ReplaceStr
Application.Calculation = xlCalculationManual
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
Application.Calculation = xlAutomatic
Application.StatusBar = False

End Sub



'==============================================================================
' Name:     OptionsAnalysis
' Author:   PS
' Desc:     Calculates all combinations of validations
' Args:     None
' Comments: (1)
'==============================================================================
Sub OptionsAnalysis()
Dim Var1Selector As String
Dim Var2Selector As String
Dim Var1Options() As Variant 'array
Dim Var2Options() As Variant 'array
Dim ResultSheet As String
Dim TargetSheet As String
Dim CalcSheet As String
Dim ResultSheetType As String
Dim TargetSheetType As String
Dim CalcSheetType As String
Dim ResultRow As Integer
Dim Var1Row As String
Dim Var2Row As String
Dim TargetRow As Integer
Dim ResultAddr As String
Dim Res_StartCol As Integer
Dim Res_EndCol As Integer
Dim Tar_StartCol As Integer
Dim Tar_EndCol As Integer
Dim AnalysisInputsString As String
Dim PreviousInputs() As String
Dim CheckRng As Range

'<-TODO: Title row
'<-TODO: check number of rows
frmOptionsAnalysis.RefTargetRng = "'" & ActiveSheet.Name & "'!" & Selection.Address
'check for previous run
Set CheckRng = Cells(Selection.Row, T_Description)
    If Not CheckRng.Comment Is Nothing Then
    PreviousInputs = Split(CheckRng.Comment.text, ",")
        If UBound(PreviousInputs) = 2 Then
            With frmOptionsAnalysis
            .RefVar1Rng = PreviousInputs(0)
            .RefVar2Rng = PreviousInputs(1)
            .RefResult = PreviousInputs(2)
            End With
        End If
    End If
        
frmOptionsAnalysis.Show

    If btnOkPressed = False Then End

'set addresses and rows
Var1Row = ExtractAddressElement(RngVar1, 2)
Var2Row = ExtractAddressElement(RngVar2, 2)
Var1Selector = Cells(CInt(Var1Row), T_Description).Address
Var2Selector = Cells(CInt(Var2Row), T_Description).Address
ResultRow = ExtractAddressElement(ResultRng, 2)
TargetRow = ExtractAddressElement(TargetRng, 2)
InitialRow = TargetRow

'set sheet names
TargetSheet = ExtractAddressElement(TargetRng, 0)
TargetSheet = TrimSheetName(TargetSheet)
CalcSheet = ExtractAddressElement(RngVar1, 0)
CalcSheet = TrimSheetName(CalcSheet)
ResultSheet = ExtractAddressElement(ResultRng, 0)
ResultSheet = TrimSheetName(ResultSheet)
'set sheet types
ResultSheetType = Sheets(ResultSheet).Range("TYPECODE").Value
CalcSheetType = Sheets(CalcSheet).Range("TYPECODE").Value
TargetSheetType = Sheets(TargetSheet).Range("TYPECODE").Value
'Store inputs for future use
AnalysisInputsString = RngVar1 & "," & RngVar2 & "," & ResultRng

Sheets(CalcSheet).Activate

    'get lists of options - variable 1
    If HasDataValidation(Range(Var1Selector)) Then
    RngVar1 = Range(Var1Selector).Validation.Formula1
    Var1Options = Range(RngVar1).Value2
    Else 'just the one option
    ReDim Var1Options(1, 1)
    Var1Options(1, 1) = Range(Var1Selector).Value
    End If

    'get lists of options - variable 2
    If HasDataValidation(Range(Var2Selector)) Then
    RngVar2 = Range(Var2Selector).Validation.Formula1
    Var2Options = Range(RngVar2).Value
    Else 'just one option
    ReDim Var2Options(1, 1)
    Var2Options(1, 1) = Range(Var2Selector).Value
    End If

    'loop through each source
    For s = 1 To UBound(Var1Options)
    
        'loop through each attenuator
        For A = 1 To UBound(Var2Options)
        
        'Debug.Print Var1Options(S, 1) & " // " & Var2Options(a, 1)
        
            If Var1Options(s, 1) <> "" And Var2Options(A, 1) <> "" Then 'skip blank entries
            'set description (and thereby values)
            Range(Var1Selector).Value = Var1Options(s, 1)
            Range(Var2Selector).Value = Var2Options(A, 1)
            
            'write to output
            Sheets(TargetSheet).Cells(TargetRow, T_Description).Value = _
                Var1Options(s, 1) & " // " & Var2Options(A, 1)
            
            'set ranges for Results / Target sheets
            Res_StartCol = GetSheetTypeColumns(ResultSheetType, "LossGainStart")
            Res_EndCol = GetSheetTypeColumns(ResultSheetType, "LossGainEnd")
            Tar_StartCol = GetSheetTypeColumns(TargetSheetType, "LossGainStart")
            Tar_EndCol = GetSheetTypeColumns(TargetSheetType, "LossGainEnd")
            
            'results
            ResultsAddr = Range(Cells(ResultRow, Res_StartCol), Cells(ResultRow, Res_EndCol)).Address
            WriteAddr = Range(Cells(TargetRow, Tar_StartCol), Cells(TargetRow, Tar_EndCol)).Address '<--TODO: make input variable
            Sheets(TargetSheet).Range(WriteAddr).Value = Range(ResultsAddr).Value
            TargetRow = TargetRow + 1
            End If
        
        Next A
        
    Next s
    
Sheets(TargetSheet).Activate


InsertComment AnalysisInputsString, T_Description, False

    'colours are nice
    If ApplyHeatMap = True Then
    Range(Cells(InitialRow, T_LossGainStart), _
        Cells(TargetRow - 1, T_LossGainStart)).Select
    HeatMap (True)
    End If
    
End Sub



'==============================================================================
' Name:     ApplyFreqValidation
' Author:   PS
' Desc:     Applies validation to frequency headers, calls conditional
'           formatting
' Args:     None
' Comments: (1)
'==============================================================================
Sub ApplyFreqValidation()

Dim Col As Integer
Dim ValidationString As String
Dim StartAddr As String

StartAddr = Selection.Address 'save where the cursor is

Cells(T_FreqRow, T_LossGainStart).Select

    For Col = T_LossGainStart To T_LossGainEnd
        If HasDataValidation(Cells(T_FreqRow, Col)) = False Then
        ValidationString = Cells(T_FreqRow, Col).Value & "," & _
            Cells(T_FreqRow, Col).Value & "*"
        SetDataValidation Col, ValidationString
        End If
    Next Col
    
ApplyFreqConditionalFormat T_LossGainStart, T_LossGainEnd

    If T_SheetType = "MECH" Then
        For Col = T_RegenStart To T_RegenEnd
            If HasDataValidation(Cells(T_FreqRow, Col)) = False Then
            ValidationString = Cells(T_FreqRow, Col).Value & "," & _
                Cells(T_FreqRow, Col).Value & "*"
            SetDataValidation Col, ValidationString
            End If
        Next Col
    ApplyFreqConditionalFormat T_RegenStart, T_RegenEnd
    End If

Range(StartAddr).Select

End Sub


'==============================================================================
' Name:     ApplyFreqConditionalFormat
' Author:   PS
' Desc:     Applies conditinoal formatting to frequency headers
' Args:     FirstCol - first column
'           LastCol - final column
' Comments: (1) Set as T_Lossgain or T_Regen variables
'==============================================================================
Sub ApplyFreqConditionalFormat(FirstCol As Integer, LastCol As Integer)

Dim FirstBand As String
Dim rngBands As Range

'conditional formatting
Set rngBands = Range(Cells(T_FreqRow, FirstCol), Cells(T_FreqRow, LastCol))
rngBands.FormatConditions.Delete

FirstBand = Cells(T_FreqRow, FirstCol).Address(False, False)

'apply new formatting
rngBands.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=RIGHT(" & FirstBand & ",1)=""*"""
rngBands.FormatConditions(rngBands.FormatConditions.Count).SetFirstPriority

    With rngBands.FormatConditions(1).Font
        .Bold = True
        .Italic = True
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    
End Sub


'==============================================================================
' Name:     ScheduleBuilder
' Author:   PS
' Desc:     Makes an attenuator schedule out of all sheets
' Args:     None
' Comments: (1)
'==============================================================================
Sub ScheduleBuilder()
'put current range in form
frmScheduleBuilder.RefTargetRng.Value = "'" & ActiveSheet.Name & "'!" & _
    Selection.Address
'show form, the rest of the code is in there
frmScheduleBuilder.Show

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
