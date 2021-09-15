Attribute VB_Name = "VARS"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
'variables for centrally accessed text files
Public ROOTPATH As String
Public TEMPLATELOCATION As String
Public STANDARDCALCLOCATION As String
Public FIELDSHEETLOCATION As String
Public EQUIPMENTSHEETLOCATION As String
Public ASHRAE_DUCT As String
Public ASHRAE_FLEX As String
Public ASHRAE_REGEN As String
Public PROJECTINFODIRECTORY As String
Public FANTECH_SILENCERS As String
Public FANTECH_DUCTS As String
Public ACOUSTIC_LOUVRES As String
Public DUCT_DIRLOSS As String
Public SRL_LG_DIRECTIVITY As String
Public SRL_DUCTS As String
Public TXT_RAW As String
Public TXT_HEAD As String
'variables for header block
Public ENGINEER As String
Public PROJECTNO As String
Public PROJECTNAME As String
'variables for column numbers
Public T_Description As Integer
Public T_LossGainStart As Integer
Public T_LossGainEnd As Integer
Public T_RegenStart As Integer
Public T_RegenEnd As Integer
Public T_ParamStart As Integer
Public T_ParamEnd As Integer
Public T_Comment As Integer
'variables for row numbers
Public T_FreqRow As Integer
'variables for Range addresses
Public T_ParamRng(3) As String
Public T_FreqStartRng As String
'variables for other controls
Public T_BandType As String
Public T_SheetType As String
'variables for form control
Public btnOkPressed As Boolean


'==============================================================================
'MASTER SET OF RANGES FOR SHEET TYPES
'note that -1 means no such column
'   STRUCTURE:
'   Array(Description, LossGainStart, LossGainEnd, RegenStart, RegenEnd,
'   ParamStart,ParamEnd,Comment,FreqRow)
'==============================================================================
Public Function OCT_cols() As Variant
OCT_cols = Array(2, 5, 13, 5, 13, 14, 15, 16, 6) 'lin and A
End Function
Public Function TO_cols() As Variant
TO_cols = Array(2, 5, 25, 5, 13, 26, 27, 28, 6) 'lin and A
End Function
Public Function MECH_cols() As Variant
MECH_cols = Array(2, 9, 17, 20, 28, 3, 6, -1, 6)
End Function
Public Function CVT_cols() As Variant
CVT_cols = Array(2, 5, 31, 34, 42, 43, 46, 47, 6)
End Function
Public Function LF_TO_cols() As Variant
LF_TO_cols = Array(2, 5, 31, -1, -1, 32, 33, 34, 6)
End Function
Public Function LF_OCT_cols() As Variant
LF_OCT_cols = Array(2, 5, 14, -1, -1, 15, 16, 17, 6)
End Function

'==============================================================================
' Name:     CurrentSheetColumns
' Author:   PS
' Desc:     Sets central variables to the correct column for that SheetType
' Args:     None
' Comments: (1) Add more columns here
'==============================================================================
Public Function CurrentSheetColumns(Optional InputSheetType As String) As Variant

Dim SheetType As String

    If InputSheetType = "" Then
    SheetType = T_SheetType
    Else
    SheetType = InputSheetType
    End If

    If Left(SheetType, 3) = "OCT" Then 'OCT OR OCTA
    CurrentSheetColumns = OCT_cols
    ElseIf Left(SheetType, 2) = "TO" Then 'TO OR TOA
    CurrentSheetColumns = TO_cols
    ElseIf SheetType = "LF" Then
    CurrentSheetColumns = LF_Cols
    ElseIf SheetType = "MECH" Then
    CurrentSheetColumns = MECH_cols
    ElseIf SheetType = "CVT" Then
    CurrentSheetColumns = CVT_cols
    ElseIf SheetType = "LF_TO" Then
    CurrentSheetColumns = LF_TO_cols
    ElseIf SheetType = "LF_OCT" Then
    CurrentSheetColumns = LF_OCT_cols
    '<---------------------------------TODO: exception for standard calc sheets
    Else
    ErrorTypeCode
    End If
End Function

'==============================================================================
' Name:     CurrentSheetBands
' Author:   PS
' Desc:     Sets central variables to the correct band types
' Args:     None
' Comments: (1)
'==============================================================================
Public Function CurrentSheetBands() As Variant
    If Left(T_SheetType, 3) = "OCT" Then 'OCT OR OCTA
    CurrentSheetBands = "oct"
    ElseIf Left(T_SheetType, 2) = "TO" Then 'TO OR TOA
    CurrentSheetBands = "to"
    ElseIf T_SheetType = "LF" Then
    CurrentSheetBands = "to"
    ElseIf T_SheetType = "MECH" Then
    CurrentSheetBands = "oct"
    ElseIf T_SheetType = "CVT" Then
    CurrentSheetBands = "cvt"
    ElseIf T_SheetType = "LF_TO" Then
    CurrentSheetBands = "to"
    ElseIf T_SheetType = "LF_OCT" Then
    CurrentSheetBands = "oct"
    Else
    ErrorTypeCode
    End If
End Function

'==============================================================================
' Name:     GetSettings
' Author:   PS
' Desc:     Sets central variables for control of elements
' Args:     None
' Comments: (1) Points to the path where Trace is installed. Sometimes this
'           doesn't work, but I can't figure out why.
'==============================================================================
Public Sub GetSettings()

On Error Resume Next

    'catches the error where excel doesn't know how to do its job
    If Application.AddIns.Count = 0 Then
    'hard coded location of AddIn as a fallback
    ROOTPATH = "U:\SectionData\Property\Specialist Services\Acoustics\1 - Technical Library\Excel Add-in\Trace"
    Else
    ROOTPATH = Application.AddIns("Trace").Path
    End If

'Debug.Print RootPath

TEMPLATELOCATION = ROOTPATH & "\Template Sheets"
STANDARDCALCLOCATION = ROOTPATH & "\Standard Calc Sheets"
FIELDSHEETLOCATION = ROOTPATH & "\Field Sheets"
EQUIPMENTSHEETLOCATION = ROOTPATH & "\Equipment Import Sheets"
ASHRAE_DUCT = ROOTPATH & "\DATA\ASHRAE_DUCTS.txt"
ASHRAE_FLEX = ROOTPATH & "\DATA\ASHRAE_FLEX.txt"
ASHRAE_REGEN = ROOTPATH & "\DATA\ASHRAE_REGEN.txt"
FANTECH_SILENCERS = ROOTPATH & "\DATA\Silencers.txt"
FANTECH_DUCTS = ROOTPATH & "\DATA\FANTECH_DUCTS.txt"
ACOUSTIC_LOUVRES = ROOTPATH & "\DATA\Louvres.txt"
DUCT_DIRLOSS = ROOTPATH & "\DATA\DuctDir.txt"
SRL_LG_DIRECTIVITY = ROOTPATH & "\DATA\SRL_LouvreGrilleDirectivity.txt"
SRL_DUCTS = ROOTPATH & "\DATA\SRL_Ducts.txt"

TestLocation TEMPLATELOCATION, vbDirectory
TestLocation STANDARDCALCLOCATION, vbDirectory
TestLocation FIELDSHEETLOCATION, vbDirectory
TestLocation EQUIPMENTSHEETLOCATION, vbDirectory
TestLocation (ASHRAE_DUCT)
TestLocation (ASHRAE_FLEX)
TestLocation (ASHRAE_REGEN)
TestLocation (FANTECH_SILENCERS)
TestLocation (FANTECH_DUCTS)
TestLocation (ACOUSTIC_LOUVRES)
TestLocation (DUCT_DIRLOSS)
TestLocation (SRL_LG_DIRECTIVITY)
TestLocation (SRL_DUCTS)

'SQLite 3DLL
'SQLite3Initialize (ROOTPATH & "\DATA")

End Sub

'==============================================================================
' Name:     GetSheetTypeColumns
' Author:   PS
' Desc:     Returns the columns for the requested type for the input SheetType.
'           Arrays are structured as:
'           Array(Description,LossStart,LossEnd,GainStart,GainEnd,ParamStart,
'           ParamEnd,Comment,FreqRow)
' Args:     SheetType, ColumnType (denoted by strings)
' Comments: (1) Supports all sheet types. Note that negative numbers are errors
'==============================================================================
Function GetSheetTypeColumns(SheetType As String, ColumnType As String)

Dim i As Integer
Dim SheetCols() As Variant
SheetCols = CurrentSheetColumns(SheetType)

    Select Case ColumnType
    Case Is = "Description"
    i = 0
    Case Is = "LossGainStart"
    i = 1
    Case Is = "LossGainEnd"
    i = 2
    Case Is = "RegenStart"
    i = 3
    Case Is = "RegenEnd"
    i = 4
    Case Is = "ParamStart"
    i = 5
    Case Is = "ParamEnd"
    i = 6
    Case Is = "Comment"
    i = 7
    Case Is = "FreqRow"
    i = 8
    End Select
    
GetSheetTypeColumns = SheetCols(i)

End Function

'==============================================================================
' Name:     ColNum2Str
' Author:   PS
' Desc:     Converts column numbers to strings
' Args:     ColNo, the column number
' Comments: (1) neat
'==============================================================================
Function ColNum2Str(ColNo As Integer) As String
Dim vArr 'variable array to hold split
vArr = Split(Cells(1, ColNo).Address(True, False), "$")
ColNum2Str = vArr(0)
End Function

'==============================================================================
' Name:     TestLocation
' Author:   PS
' Desc:     Tests if a reference text file exists.
' Args:     PathStr, the path to be tested and SearchType (defaults to file,
'           but can be set as vbDirectory
' Comments: (1) All text files are tested during GetSettings(), with a warning
'           coming up if the file isn't found.
'==============================================================================
Function TestLocation(PathStr As String, Optional SearchType)

If IsMissing(SearchType) Then SearchType = vbNormal

    If Dir(PathStr, SearchType) = "" Then
    TestLocation = False
        If SearchType = vbDirectory Then
        msg = MsgBox("Directory '" & PathStr & " not found!", _
            vbOKOnly, "Trace Error - Missing data file!")
        Else
        msg = MsgBox("File '" & PathStr & " not found!", _
            vbOKOnly, "Trace Error - Missing data file!")
        End If
        
'    '***********
'    End '<-Escape!
'    '***********
    
    Else
    TestLocation = True
'        If SearchType = vbDirectory Then
'        Debug.Print "Directory Found!    "; PathStr
'        Else
'        Debug.Print "File Found!         "; PathStr
'        End If
    End If
    
End Function

'==============================================================================
' Name:     ScreenInput
' Author:   PS
' Desc:     Checks for numeric inputs and converts to Double if required
' Args:     X, Variant
' Comments: (1) Used mostly in forms
'==============================================================================
Function ScreenInput(x As Variant, Optional Mode As String) '<-------TODO Can we use this elsewhere?
    If IsNumeric(x) Then
    ScreenInput = CDbl(x)
    Else
    ScreenInput = 0
    End If
End Function


'==============================================================================
' Name:     SetSheetTypeControls
' Author:   PS
' Desc:     Sets the column numbers for the 8 defined column ranges:
'           T_Description, T_LossGainStart, T_LossGainEnd, T_RegenStart,
'           T_RegenEnd, T_ParamStart, T_ParamEnd, T_Comment
'           AND
'           Address strings for other ranges
' Args:     None
' Comments: (1) Called whenever a function is input into a sheet
'           (2) includes CheckTemplateRow
'==============================================================================
Sub SetSheetTypeControls()
Dim SheetCols() As Variant

    If IsNamedRange("TYPECODE") Then
    T_SheetType = ActiveSheet.Range("TYPECODE").Value
    Else
    ErrorTypeCode
    End If
    
'if you're setting columns for function import, check the row number as well
CheckTemplateRow (Selection.Row)
SheetCols = CurrentSheetColumns

'set public variables for columns and rows
T_Description = SheetCols(0)
T_LossGainStart = SheetCols(1)
T_LossGainEnd = SheetCols(2)
T_RegenStart = SheetCols(3)
T_RegenEnd = SheetCols(4)
T_ParamStart = SheetCols(5)
T_ParamEnd = SheetCols(6)
T_Comment = SheetCols(7)
T_FreqRow = SheetCols(8)
'Central addressses, note: absolute row
T_FreqStartRng = Cells(T_FreqRow, T_LossGainStart).Address(True, False)
    If T_ParamStart > 0 Then
        For i = LBound(T_ParamRng) To UBound(T_ParamRng)
        T_ParamRng(i) = Cells(Selection.Row, T_ParamStart + i). _
            Address(False, True) 'absolute column
        Next
    End If
'band type: oct/to/cvt
T_BandType = CurrentSheetBands
End Sub

'==============================================================================
' Name:     CheckTemplateRow
' Author:   PS
' Desc:     checks if the user is trying to put functions in header/default rows
' Args:     rw - Row Number
' Comments: (1) 'Checks that user isn't in header rows.
'           These rows are protected by this function.
'           None shall Pass.
'==============================================================================
Public Sub CheckTemplateRow(Rw As Integer)
Dim MoveToCalcRow As Long
Dim FirstRow As Integer
    
    'set first row number
    If T_SheetType = "MECH" And Rw > 38 Then
    '********************
    End 'stop everything!
    '********************
    Else
    FirstRow = 7
    End If

    'check if it's too far up
    If Rw <= FirstRow Then
    MoveToCalcRow = MsgBox("Looks like your cursor is in the header block. " & _
        chr(10) & "Do you want to move down to the first calculation row?", _
        vbYesNo, "Down down....")
        
        If MoveToCalcRow = vbYes Then
        Cells(8, Selection.Column).Select
        Else
        '********************
        End 'stop everything!
        '********************
        End If

    End If
    
End Sub


'==============================================================================
' Name:     FindFrequencyBand
' Author:   PS
' Desc:     Returns the column number of SearchBand
' Args:     SearchBand - the band to be matched
' Comments: (1) Useful in Rw and STC
'==============================================================================
Public Function FindFrequencyBand(SearchBand As String)
Dim i As Integer
i = -1
    For i = T_LossGainStart To T_LossGainEnd
        If Cells(T_FreqRow, i).Value = SearchBand Then
        FindFrequencyBand = i
        Exit Function
        End If
    Next i
End Function











'BLANK HEADER BLOCK:


'==============================================================================
' Name:
' Author:
' Desc:
' Args:
' Comments: (1)
'==============================================================================
