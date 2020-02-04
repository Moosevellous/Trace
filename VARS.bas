Attribute VB_Name = "VARS"
''''''''''''''''''''''
'GLOBAL VARIABLES'''''
''''''''''''''''''''''
Public TEMPLATELOCATION As String
Public STANDARDCALCLOCATION As String
Public FIELDSHEETLOCATION As String
Public EQUIPMENTSHEETLOCATION As String
Public ASHRAE_DUCT As String
Public ASHRAE_FLEX As String
Public ASHRAE_REGEN As String
Public ENGINEER As String
Public PROJECTNO As String
Public PROJECTNAME As String
Public PROJECTINFODIRECTORY As String
Public FANTECH_SILENCERS As String
Public FANTECH_DUCTS As String
Public ACOUSTIC_LOUVRES As String
Public DUCT_DIRLOSS As String
Public colourUSERINPUT As Long
Public colourFINALRESULT  As Long


''''''''''''''''''''''
'END GLOBAL VARIABLES
''''''''''''''''''''''

'sub to initialise
Public Sub GetSettings()
Dim RootPath  As String

On Error Resume Next

'Debug.Print "No of addins: " & Application.AddIns.Count
'    For Each ad In Application.AddIns
'    Debug.Print ad.Name
'        If ad.Name = "NoiseCalc.xlam" Then
'        RootPath = Application.AddIns("NoiseCalc").Path
'        End If
'    Next ad

    If Application.AddIns.Count = 0 Then 'catches the error where excel doesn't know how to do its job
    RootPath = "U:\SectionData\Property\Specialist Services\Acoustics\1 - Technical Library\Excel Add-in\Trace" 'hard coded location of AddIn as a fallback
    Else
    RootPath = Application.AddIns("Trace").Path
    End If

'Debug.Print RootPath

TEMPLATELOCATION = RootPath & "\Template Sheets\Blank Calculation Sheet.xlsm"
STANDARDCALCLOCATION = RootPath & "\Standard Calc Sheets"
FIELDSHEETLOCATION = RootPath & "\Field Sheets"
EQUIPMENTSHEETLOCATION = RootPath & "\Equipment Import Sheets"
ASHRAE_DUCT = RootPath & "\DATA\ASHRAE_DUCTS.txt"
ASHRAE_FLEX = RootPath & "\DATA\ASHRAE_FLEX.txt"
ASHRAE_REGEN = RootPath & "\DATA\ASHRAE_REGEN.txt"
FANTECH_SILENCERS = RootPath & "\DATA\Silencers.txt"
FANTECH_DUCTS = RootPath & "\DATA\FANTECH_DUCTS.txt"
ACOUSTIC_LOUVRES = RootPath & "\DATA\Louvres.txt"
DUCT_DIRLOSS = RootPath & "\DATA\DuctDir.txt"

TestLocation (TEMPLATELOCATION)
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

'Colours
colourUSERINPUT = RGB(254, 253, 195)
colourFINALRESULT = RGB(146, 205, 220)
End Sub

Function GetSheetTypeRanges(SheetType As String, ColumnType As String)

'Arrays of values for different sheet types
'Arrays are structured (Description,LossStart,LossEnd,GainStart,GainEnd,ParamStart,ParamEnd,Comment) <-------7 elements
'Negative numbers are errors

Dim i As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MASTER TABLE OF RANGES FOR SHEET TYPES
'note that -2 means no such column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
OCT_cols = Array(2, 5, 13, -1, -1, 14, 15, 16)
TO_cols = Array(2, 5, 25, -1, -1, 26, 27, 28)
MECH_cols = Array(2, 9, 17, 20, 28, 3, 6, -1)
CVT_cols = Array(2, 5, 31, 33, 41, -1, -1, 32)
'CONVERT SHEET?
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Left(SheetType, 3) = "OCT" Then 'OCT OR OCTA
    SheetTypeArray = OCT_cols
    ElseIf Left(SheetType, 2) = "TO" Then 'TO OR TOA
    SheetTypeArray = TO_cols
    ElseIf SheetType = "LF" Then
    SheetTypeArray = LF_Cols
    ElseIf SheetType = "MECH" Then
    SheetTypeArray = MECH_cols
    Else
    ErrorTypeCode
    End
    End If

    Select Case ColumnType
    Case Is = "Description"
    i = 0
    Case Is = "DataStart"
    i = 1
    Case Is = "DataEnd"
    i = 2
    Case Is = "LossStart" 'same as DataStart, but for the sake of flexible naming.....
    i = 3
    Case Is = "LossEnd" 'as above
    i = 2
    Case Is = "GainStart"
    i = 3
    Case Is = "GainEnd"
    i = 4
    Case Is = "ParamStart"
    i = 5
    Case Is = "ParamEnd"
    i = 6
    Case Is = "Comment"
    i = 7
    End Select

GetSheetTypeRanges = SheetTypeArray(i)
    
End Function

Function TestLocation(PathStr As String, Optional SearchType)

If IsMissing(SearchType) Then SearchType = vbNormal

    If Dir(PathStr, SearchType) = "" Then
    TestLocation = False
        If SearchType = vbDirectory Then
        msg = MsgBox("Directory '" & PathStr & " not found!", vbOKOnly, "Trace Error - Missing data file!")
        Else
        msg = MsgBox("File '" & PathStr & " not found!", vbOKOnly, "Trace Error - Missing data file!")
        End If
        
    '***********
    End
    '***********
    
    Else
    TestLocation = True
'        If SearchType = vbDirectory Then
'        Debug.Print "Directory Found!    "; PathStr
'        Else
'        Debug.Print "File Found!         "; PathStr
'        End If
    End If
    
End Function



Function ScreenInput(X As Variant) 'to screen data input into forms and functions
    If IsNumeric(X) Then
    ScreenInput = CDbl(X)
    Else
    ScreenInput = 0
    End If
End Function

