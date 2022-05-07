Attribute VB_Name = "Basics"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Dim ConditionValue As Variant 'why is this dim and not public?
Dim SearchString As String
Public BasicFunctionType As String
Public RangeSelection As String
Public Range2Selection As String
Public ApplyToSheetType As Boolean 'if true, will apply style to sheet type
Public BasicsApplyStyle As String 'sets style
'Frequency bands
Public FBC_bandwidth As Integer
Public FBC_mode As String
Public FBC_baseTen As Boolean
'Mass-air-mass
Public MAM_M1 As Double
Public MAM_M2 As Double
Public MAM_Width As Double
Public MAM_Description As String
'Room Corrections
Public DistanceFromSource As Double
Public RoomVolume As Double

'==============================================================================
' Name:     freqStr2Num
' Author:   PS
' Desc:     Converts octave band strings to values
' Args:     fstr, the frequency band centre frequency as a string
' Comments: (1) Used almost everywhere, because '2k' beats writing '2000' etc
'==============================================================================
Function freqStr2Num(fStr As String) As Double

If Right(fStr, 1) = "*" Then 'trim stars
    'trim
    fStr = Left(fStr, Len(fStr) - 1)
End If

    Select Case fStr
    Case Is = "1"
    freqStr2Num = 1
    Case Is = "1.25"
    freqStr2Num = 1.25
    Case Is = "1.6"
    freqStr2Num = 1.6
    Case Is = "2"
    freqStr2Num = 2
    Case Is = "2.5"
    freqStr2Num = 2.5
    Case Is = "3.15"
    freqStr2Num = 3.15
    Case Is = "4"
    freqStr2Num = 4
    Case Is = "5"
    freqStr2Num = 5
    Case Is = "6.3"
    freqStr2Num = 6.3
    Case Is = "8"
    freqStr2Num = 8
    Case Is = "10"
    freqStr2Num = 10
    Case Is = "12.5"
    freqStr2Num = 12.5
    Case Is = "16"
    freqStr2Num = 16
    Case Is = "20"
    freqStr2Num = 20
    Case Is = "25"
    freqStr2Num = 25
    Case Is = "31.5"
    freqStr2Num = 31.5
    Case Is = "40"
    freqStr2Num = 40
    Case Is = "50"
    freqStr2Num = 50
    Case Is = "63"
    freqStr2Num = 63
    Case Is = "80"
    freqStr2Num = 80
    Case Is = "100"
    freqStr2Num = 100
    Case Is = "125"
    freqStr2Num = 125
    Case Is = "160"
    freqStr2Num = 160
    Case Is = "200"
    freqStr2Num = 200
    Case Is = "250"
    freqStr2Num = 250
    Case Is = "315"
    freqStr2Num = 315
    Case Is = "400"
    freqStr2Num = 400
    Case Is = "500"
    freqStr2Num = 500
    Case Is = "630"
    freqStr2Num = 630
    Case Is = "800"
    freqStr2Num = 800
    Case Is = "1k"
    freqStr2Num = 1000
    Case Is = "1000"
    freqStr2Num = 1000
    Case Is = "1.25k"
    freqStr2Num = 1250
    Case Is = "1250"
    freqStr2Num = 1250
    Case Is = "1.6k"
    freqStr2Num = 1600
    Case Is = "1600"
    freqStr2Num = 1600
    Case Is = "2k"
    freqStr2Num = 2000
    Case Is = "2000"
    freqStr2Num = 2000
    Case Is = "2.5k"
    freqStr2Num = 2500
    Case Is = "2500"
    freqStr2Num = 2500
    Case Is = "3.15k"
    freqStr2Num = 3150
    Case Is = "3150"
    freqStr2Num = 3150
    Case Is = "4k"
    freqStr2Num = 4000
    Case Is = "4000"
    freqStr2Num = 4000
    Case Is = "5k"
    freqStr2Num = 5000
    Case Is = "5000"
    freqStr2Num = 5000
    Case Is = "6.3k"
    freqStr2Num = 6300
    Case Is = "6300"
    freqStr2Num = 6300
    Case Is = "8k"
    freqStr2Num = 8000
    Case Is = "8000"
    freqStr2Num = 8000
    Case Is = "10k"
    freqStr2Num = 10000
    Case Is = "10000"
    freqStr2Num = 10000
    Case Is = "12.5k"
    freqStr2Num = 12500
    Case Is = "12500"
    freqStr2Num = 12500
    Case Is = "16k"
    freqStr2Num = 16000
    Case Is = "16000"
    freqStr2Num = 16000
    Case Is = "20k"
    freqStr2Num = 20000
    Case Is = "20000"
    freqStr2Num = 20000
    Case Else 'catch the exception
    freqStr2Num = 0
    End Select
    
End Function

'==============================================================================
' Name:     GetArrayIndex_OCT
' Author:   PS
' Desc:     Returns an array index of octave bands, starting from 63Hz
' Args:     fStr - octave band centre frequency
'           OffsetBands - positive value=index up for a given band
' Comments: (1) Used a lot in ISO9613 but also elsewhere
'==============================================================================
Function GetArrayIndex_OCT(fStr As String, Optional OffsetBands As Integer)
Dim freq As Double

freq = freqStr2Num(fStr) 'converts to Double
    Select Case freq
    Case Is = 16
    GetArrayIndex_OCT = -2 + OffsetBands
    Case Is = 31.5
    GetArrayIndex_OCT = -1 + OffsetBands
    Case Is = 63 'default
    GetArrayIndex_OCT = 0 + OffsetBands
    Case Is = 125
    GetArrayIndex_OCT = 1 + OffsetBands
    Case Is = 250
    GetArrayIndex_OCT = 2 + OffsetBands
    Case Is = 500
    GetArrayIndex_OCT = 3 + OffsetBands
    Case Is = 1000
    GetArrayIndex_OCT = 4 + OffsetBands
    Case Is = 2000
    GetArrayIndex_OCT = 5 + OffsetBands
    Case Is = 4000
    GetArrayIndex_OCT = 6 + OffsetBands
    Case Is = 8000
    GetArrayIndex_OCT = 7 + OffsetBands
    Case Else
    GetArrayIndex_OCT = 999 'for catching errors
    End Select
    
End Function

'==============================================================================
' Name:     GetArrayIndex_TO
' Author:   PS
' Desc:     Returns an array index of one-third octave bands, starting from 50Hz
' Args:     fStr - octave band centre frequency
'           OffsetBands - positive value=index up for a given band
' Comments: (1) TODO: update for string inputs for consistency?
'==============================================================================
Function GetArrayIndex_TO(f As Double, Optional OffsetBands As Integer)

    Select Case f
    Case Is = 1
    GetArrayIndex_TO = -17 + OffsetBands
    Case Is = 1.25
    GetArrayIndex_TO = -16 + OffsetBands
    Case Is = 1.6
    GetArrayIndex_TO = -15 + OffsetBands
    Case Is = 2
    GetArrayIndex_TO = -14 + OffsetBands
    Case Is = 2.5
    GetArrayIndex_TO = -13 + OffsetBands
    Case Is = 3.15
    GetArrayIndex_TO = -12 + OffsetBands
    Case Is = 4
    GetArrayIndex_TO = -11 + OffsetBands
    Case Is = 5
    GetArrayIndex_TO = -10 + OffsetBands
    Case Is = 6.3
    GetArrayIndex_TO = -9 + OffsetBands
    Case Is = 8
    GetArrayIndex_TO = -8 + OffsetBands
    Case Is = 10
    GetArrayIndex_TO = -7 + OffsetBands
    Case Is = 12.5
    GetArrayIndex_TO = -6 + OffsetBands
    Case Is = 16
    GetArrayIndex_TO = -5 + OffsetBands
    Case Is = 20
    GetArrayIndex_TO = -4 + OffsetBands
    Case Is = 25
    GetArrayIndex_TO = -3 + OffsetBands
    Case Is = 31.5
    GetArrayIndex_TO = -2 + OffsetBands
    Case Is = 40
    GetArrayIndex_TO = -1 + OffsetBands
    Case Is = 50
    GetArrayIndex_TO = 0 + OffsetBands 'DEFAULT
    Case Is = 63
    GetArrayIndex_TO = 1 + OffsetBands
    Case Is = 80
    GetArrayIndex_TO = 2 + OffsetBands
    Case Is = 100
    GetArrayIndex_TO = 3 + OffsetBands
    Case Is = 125
    GetArrayIndex_TO = 4 + OffsetBands
    Case Is = 160
    GetArrayIndex_TO = 5 + OffsetBands
    Case Is = 200
    GetArrayIndex_TO = 6 + OffsetBands
    Case Is = 250
    GetArrayIndex_TO = 7 + OffsetBands
    Case Is = 315
    GetArrayIndex_TO = 8 + OffsetBands
    Case Is = 400
    GetArrayIndex_TO = 9 + OffsetBands
    Case Is = 500
    GetArrayIndex_TO = 10 + OffsetBands
    Case Is = 630
    GetArrayIndex_TO = 11 + OffsetBands
    Case Is = 800
    GetArrayIndex_TO = 12 + OffsetBands
    Case Is = 1000
    GetArrayIndex_TO = 13 + OffsetBands
    Case Is = 1250
    GetArrayIndex_TO = 14 + OffsetBands
    Case Is = 1600
    GetArrayIndex_TO = 15 + OffsetBands
    Case Is = 2000
    GetArrayIndex_TO = 16 + OffsetBands
    Case Is = 2500
    GetArrayIndex_TO = 17 + OffsetBands
    Case Is = 3150
    GetArrayIndex_TO = 18 + OffsetBands
    Case Is = 4000
    GetArrayIndex_TO = 19 + OffsetBands
    Case Is = 5000
    GetArrayIndex_TO = 20 + OffsetBands
    Case Is = 6300
    GetArrayIndex_TO = 21 + OffsetBands
    Case Is = 8000
    GetArrayIndex_TO = 22 + OffsetBands
    Case Is = 10000
    GetArrayIndex_TO = 23 + OffsetBands
    Case Is = 12500
    GetArrayIndex_TO = 23 + OffsetBands
    Case Is = 16000
    GetArrayIndex_TO = 24 + OffsetBands
    Case Is = 20000
    GetArrayIndex_TO = 25 + OffsetBands
    Case Else
    GetArrayIndex_TO = -1
    End Select
End Function

'==============================================================================
' Name:     MapOneThird2Oct
' Author:   PS
' Desc:     Returns index of octave bands, based on groupings of one-third bands
' Args:     f_input, one third octave band centre frequency
' Comments: (1) Set to 50Hz, could make this more flexible?
'==============================================================================
Function MapOneThird2Oct(f_input As Double)
'map a 1/3 octave centre frequency to the relevant 1/1 octave band centre frequency
'OR get column index of octave band centre frequencies
    Select Case f_input
    Case Is = 50
    MapOneThird2Oct = 0
    Case Is = 63
    MapOneThird2Oct = 0
    Case Is = 80
    MapOneThird2Oct = 0
    Case Is = 100
    MapOneThird2Oct = 1
    Case Is = 125
    MapOneThird2Oct = 1
    Case Is = 160
    MapOneThird2Oct = 1
    Case Is = 200
    MapOneThird2Oct = 2
    Case Is = 250
    MapOneThird2Oct = 2
    Case Is = 315
    MapOneThird2Oct = 2
    Case Is = 400
    MapOneThird2Oct = 3
    Case Is = 500
    MapOneThird2Oct = 3
    Case Is = 630
    MapOneThird2Oct = 3
    Case Is = 800
    MapOneThird2Oct = 4
    Case Is = 1000
    MapOneThird2Oct = 4
    Case Is = 1250
    MapOneThird2Oct = 4
    Case Is = 1600
    MapOneThird2Oct = 5
    Case Is = 2000
    MapOneThird2Oct = 5
    Case Is = 2500
    MapOneThird2Oct = 5
    Case Is = 3150
    MapOneThird2Oct = 6
    Case Is = 4000
    MapOneThird2Oct = 6
    Case Is = 5000
    MapOneThird2Oct = 6
    Case Else 'catch array errors with this line
    MapOneThird2Oct = -1
    End Select
End Function

'==============================================================================
' Name:     SPLSUM
' Author:   PS
' Desc:     Logarithmically adds all cells in the input range rng1
' Args:     rng1, an array of cells
' Comments: (1) Built to be flexible and useful.
'==============================================================================
Public Function SPLSUM(ParamArray Rng1() As Variant) As Variant
On Error Resume Next

Dim c As Range
Dim i As Long

SPLSUM = -99
    For i = LBound(Rng1) To UBound(Rng1)
    'Debug.Print TypeName(rng1(i))
        If TypeName(Rng1(i)) = "Double" Then
            If Rng1(i) > 0 Then 'negative values are ignored
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + (10 ^ (Rng1(i) / 10)))
            End If
        ElseIf TypeName(Rng1(i)) = "Range" Then
            For Each c In Rng1(i).Cells
                If c.Value <> Empty And IsNumeric(c.Value) Then
                SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                    (10 ^ (SPLSUM / 10)) + (10 ^ (c.Value / 10)))
                End If
            Next c
        End If
    Next i

    'catch exceptions
    If SPLSUM < 0 Then SPLSUM = Empty

End Function

'==============================================================================
' Name:     SPLAV
' Author:   PS
' Desc:     Logarithmically averages all cells in the input range rng1
' Args:     rng1, an array of cells
' Comments: (1) Built to be flexible and useful.
'==============================================================================
Public Function SPLAV(ParamArray Rng1() As Variant) As Variant
On Error Resume Next

Dim c As Range
Dim i As Long
Dim n As Integer 'number of values
SPLAV = -99
n = 0
    For i = LBound(Rng1) To UBound(Rng1)
    'Debug.Print TypeName(Rng1(i))
        If TypeName(Rng1(i)) = "Double" Then
            If Rng1(i) > 0 Then 'negative values are ignored
            SPLAV = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLAV / 10)) + (10 ^ (Rng1(i) / 10)))
            n = n + 1
            End If
        ElseIf TypeName(Rng1(i)) = "Range" Then
            For Each c In Rng1(i).Cells
                If c.Value <> Empty And IsNumeric(c.Value) Then
                SPLAV = 10 * Application.WorksheetFunction.Log10( _
                    (10 ^ (SPLAV / 10)) + (10 ^ (c.Value / 10)))
                n = n + 1
                End If
            Next c
        End If
    Next i

'Average +10log(1/n) in log domain
SPLAV = SPLAV + 10 * Application.WorksheetFunction.Log10(1 / n)

    'catch exceptions
    If SPLAV < 0 Then SPLAV = Empty

End Function

'==============================================================================
' Name:     SPLMINUS
' Author:   PS
' Desc:     Logarithmically subtraces one cell from another
' Args:     SPLtotal, SPL2 (the one to be subtracted)
' Comments: (1) One line macros ftw!
'==============================================================================
Public Function SPLMINUS(SPLtotal As Double, SPL2 As Double) As Variant
On Error Resume Next
SPLMINUS = 10 * Application.WorksheetFunction.Log10( _
    (10 ^ (SPLtotal / 10)) - (10 ^ (SPL2 / 10)))
    
    'catch exceptions
    If SPLMINUS < 0 Then SPLMINUS = Empty
End Function

'==============================================================================
' Name:     TL_ThirdsToOctave
' Author:   PS
' Desc:     Converts transmission losses from 1/3 octave to 1/1 octave bands
' Args:     rngInput, the three cell array of TLs
' Comments: (1) Assumes the values come in as a horizontal array of cells
'           (2) can cope with negative inputs, and returns negative values
'           back, as per the Trace convention
'==============================================================================
Public Function TL_ThirdsToOctave(rngInput As Range) As Variant

Dim isNegative As Boolean
Dim TL1 As Single, TL2 As Single, TL3 As Single 'values for each band

'Debug.Print TypeName(rngInput.Cells(1, 1).Value)

TL1 = rngInput.Cells(1, 1).Value
TL2 = rngInput.Cells(1, 2).Value
TL3 = rngInput.Cells(1, 3).Value

    'switch for pos/neg sign (losses should be positive)
    If TL1 < 0 And TL2 < 0 And TL3 < 0 Then isNegative = True

    'flip signs
    If TL1 < 0 Then TL1 = TL1 * -1
    If TL2 < 0 Then TL2 = TL2 * -1
    If TL3 < 0 Then TL3 = TL3 * -1
    

    If isNegative Then 'return result as negative
    TL_ThirdsToOctave = 10 * Application.WorksheetFunction.Log10((1 / 3) * _
        ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    Else 'return result as positive
    TL_ThirdsToOctave = -10 * Application.WorksheetFunction.Log10((1 / 3) * _
        ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    End If

End Function

'==============================================================================
' Name:     CompositeTL
' Author:   PS
' Desc:     Calculates the composite transmission loss given input TLs and
'           areas of each element
' Args:     TL_Range (TL spectrum), AreaRange (Surface areas of each element)
' Comments: (1) Function Incomplete?
'==============================================================================
Public Function CompositeTL(TL_Range As Range, AreaRange As Range) As Variant

Dim TotalArea As Double
Dim TotalWeightedTL As Double
Dim Switch As Integer

TotalArea = 0
TotalWeightedTL = 0
i = 1

    If TL_Range(1) < 0 Then 'TL is negative values
    Switch = 1
    Else
    Switch = -1
    End If
    
    'calculate weighted TLs
    For Each a In AreaRange
    TotalArea = TotalArea + a
    TotalWeightedTL = TotalWeightedTL + a * (10 ^ (Switch * TL_Range(i) / 10))
    i = i + 1
    Next a

'take log of the whole thing
CompositeTL = (Switch * -1) * 10 * Application.WorksheetFunction.Log _
    (TotalArea / TotalWeightedTL)

End Function

'==============================================================================
' Name:     SPLSUMIF
' Author:   PS
' Desc:     Logarithmically adds values, if a criterion is met
' Args:     SumRange (values to be added), Condition (the type of criterion to be
'           evaluated), and ConditionRange (values to be evaluated, if not the
'           sum range itself)
' Comments: (1) Currently supports > >= < <= and =, no wildcard matching
'           (2) Now includes Match for wildcard searches
'==============================================================================
Function SPLSUMIF(SumRange As Range, Condition As String, _
    Optional ConditionRange As Variant)

Dim IfRange As Range
Dim SheetNm As String 'name of sheet

    'Check which Range will be evaluating the IF function
    If IsMissing(ConditionRange) Then
    Set IfRange = SumRange
    Else
    Set IfRange = ConditionRange
    End If

ConditionType = FindConditionType(Condition)
    
'initialise function
SPLSUMIF = -99
SheetNm = IfRange.Worksheet.Name

    For Each c In IfRange.Cells
    
'    Debug.Print "row: "; C.Row; "column: "; C.Column
'    Debug.Print "Condition test: "; ConditionType; " "; C.Value
'    Debug.Print "Cell value: "; Sheets(SheetNm).Cells(C.Row, C.Column).Value
'    Debug.Print ""
    
    Rw = IfRange.Row - SumRange.Row
    clmn = IfRange.Column - SumRange.Column
    
        Select Case ConditionType
        Case Is = "GreaterThan"
            If c.Value > ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "GreaterThanEqualTo"
            If c.Value >= ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "LessThan"
            If c.Value < ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "LessThanEqualTo"
            If c.Value <= ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "Equals"
            If c.Value = ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "Match"
            If InStr(1, c.Value, ConditionValue, vbTextCompare) > 0 Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUMIF / 10)) + _
                (10 ^ (Sheets(SheetNm).Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            End If
        Case Is = "" 'no condtion type
            SPLSUMIF = -99
        End Select
    Next c

End Function

'==============================================================================
' Name:     SPLAVIF
' Author:   PS
' Desc:     Logarithmically averages values, if a criterion is met
' Args:     SumRange (values to be added), Condition (the type of criterion to be
'           evaluated), and ConditionRange (values to be evaluated, if not the
'           sum range itself)
' Comments: (1) Currently supports > >= < <= and =, no wildcard matching
'           (2) Now includes Match for wildcard searches
'==============================================================================
Function SPLAVIF(SumRange As Range, ConditionStr As String, _
    Optional ConditionRange As Variant)

Dim IfRange As Range
Dim numVals As Integer
Dim ConditionType As String
Dim SPLSUM As Single
Dim SheetNm As String 'name of sheet

    'Check which Range will be evaluating the IF function
    If IsMissing(ConditionRange) Then
    Set IfRange = SumRange
    Else
    Set IfRange = ConditionRange
    End If

ConditionType = FindConditionType(ConditionStr)
    
'initialise function
SPLSUM = -99
SPLAVIF = -99
numVals = 0

    For Each c In IfRange.Cells
    
'    Debug.Print "row: "; C.Row; "column: "; C.Column
'    Debug.Print "Condition test: "; ConditionType; " "; C.Value
'    Debug.Print "Cell value: "; SumRange(C.Row, C.Column).Value
'    Debug.Print ""
    
    Rw = IfRange.Row - SumRange.Row
    clmn = IfRange.Column - SumRange.Column
    
        Select Case ConditionType
        Case Is = "GreaterThan"
            If c.Value > ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "GreaterThanEqualTo"
            If c.Value >= ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "LessThan"
            If c.Value < ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "LessThanEqualTo"
            If c.Value <= ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "Equals"
            If c.Value = ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "Match"
            If InStr(1, c.Value, ConditionValue, vbTextCompare) > 0 Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10( _
                (10 ^ (SPLSUM / 10)) + _
                (10 ^ (Cells(c.Row - Rw, c.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "" 'no condtion type
            SPLSUM = -99
        End Select
    Next c
    
'Debug.Print numVals; "Values:"
'subtract 10log(n) to average the result
SPLAVIF = SPLSUM - (10 * Application.WorksheetFunction.Log(numVals))

End Function

'==============================================================================
' Name:     FindConditionType
' Author:   PS
' Desc:     Filters conditions for SPLSUMIF and SPLAVIF
' Args:     inputFormula
' Comments: (1) Currently supports > >= < <= and =, no wildcard matching
'           (2) Defaults to "=" if no character
'           (3) Removed variable TypeFound - not required
'==============================================================================
Function FindConditionType(inputFormula As String)
  
    If Left(inputFormula, 2) = ">=" Then
    FindConditionType = "GreaterThanEqualTo"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 2))
    ElseIf Left(inputFormula, 2) = "<=" Then
    FindConditionType = "LessThanEqualTo"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 2))
    ElseIf Left(inputFormula, 1) = "=" Then
    FindConditionType = "Equals"
    ConditionValue = Right(inputFormula, Len(inputFormula) - 1)
        'for numbers as filters
        If IsNumeric(ConditionValue) Then
        ConditionValue = CSng(ConditionValue)
        End If
    ElseIf Left(inputFormula, 1) = "<" Then
    FindConditionType = "LessThan"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 1))
    ElseIf Left(inputFormula, 1) = ">" Then
    FindConditionType = "GreaterThan"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 1))
    ElseIf Left(inputFormula, 1) = "*" And Right(inputFormula, 1) = "*" Then
    FindConditionType = "Match"
    ConditionValue = Mid(inputFormula, 2, Len(inputFormula) - 2)
    Else 'default to equals
    FindConditionType = "Equals"
    End If
    
End Function

'==============================================================================
' Name:     FitzroyRT
' Author:   PS
' Desc:     Calculates the reverberation time according to Fitzroy's method
' Args:     Dimensions of room (x/y/z) in metres
'           Si - surface area of each element in m^2
'           Direction (assignment of each surface)
'           alpha_i - absorption value, alpha of that element
' Comments: (1) Watch out for the averaging formula, not always defined in the
'           textbooks.
'           (2) Includes an error catch so you never try and evaluate Ln(0)
'==============================================================================
Public Function FitzroyRT(x As Long, y As Long, z As Long, S_i As Range, _
    Direction As Range, alpha_i As Range)
'a_x is alpha-bar x, ie the average absorption for surfaces in the x-direction
Dim a_x As Single
Dim a_y As Single
Dim a_z As Single
Dim Sx_total As Single
Dim Sy_total As Single
Dim Sz_total As Single
Dim S_total As Single
Dim Volume As Single

If S_i.Count <> alpha_i.Count Then
FitzroyRT = vbError
End If

'average the total absorption in each direction
    For elem = 1 To S_i.Count
'    Debug.Print Direction(elem); "    "; alpha_i(elem); "    "; S_i(elem)
        If S_i(elem) > 0 Then 'ignore areas of 0 or negative values
            Select Case Direction(elem)
            
            Case Is = "X"
            a_x = a_x + (S_i(elem) * alpha_i(elem))
            Sx_total = Sx_total + S_i(elem)
            Case Is = "x"
            a_x = a_x + (S_i(elem) * alpha_i(elem))
            Sx_total = Sx_total + S_i(elem)
            Case Is = "Y"
            a_y = a_y + (S_i(elem) * alpha_i(elem))
            Sy_total = Sy_total + S_i(elem)
            Case Is = "y"
            a_y = a_y + (S_i(elem) * alpha_i(elem))
            Sy_total = Sy_total + S_i(elem)
            Case Is = "Z"
            a_z = a_z + (S_i(elem) * alpha_i(elem))
            Sz_total = Sz_total + S_i(elem)
            Case Is = "z"
            a_z = a_z + (S_i(elem) * alpha_i(elem))
            Sz_total = Sz_total + S_i(elem)
            End Select
        End If
    Next elem

S_total = Sx_total + Sy_total + Sz_total
a_x = a_x / Sx_total
a_y = a_y / Sy_total
a_z = a_z / Sz_total

'catch error when alphaBar=1 and ln(0)=ERROR
If a_x = 1 Then a_x = 0.99999
If a_y = 1 Then a_y = 0.99999
If a_z = 1 Then a_z = 0.99999

Volume = x * y * z

'Debug.Print "ax:"; a_x; "   ay:"; a_y; "   az"; a_z

FitzroyRT = (0.161 * Volume / S_total ^ 2) * _
    (((-Sx_total / Application.WorksheetFunction.Ln(1 - a_x)) + _
    (-Sy_total / Application.WorksheetFunction.Ln(1 - a_y)) + _
    (-Sz_total / Application.WorksheetFunction.Ln(1 - a_z))))

End Function

'==============================================================================
' Name:     SpeedOfSound
' Author:   PS
' Desc:     Returns speed of sound in air
' Args:     Temperature, and optional switch for degrees Kelvin (assumed false)
' Comments: (1) Perhaps there's a more advanced method, but this'll do for now
'==============================================================================
Function SpeedOfSound(temp As Long, Optional IsKelvin As Boolean)
    If IsKelvin = False Then 'convert to kelvin, not hobbs
    temp = temp + 273.15
    End If
SpeedOfSound = (1.4 * 287.1848 * temp) ^ 0.5 'square root of Gamma * R * Temp for air
End Function

'==============================================================================
' Name:     Wavelength
' Author:   PS
' Desc:     Returns wavelength for an input frequency and speed of sound
' Args:     fstr, SoundSpeed
' Comments: (1) Simple, yeah?
'==============================================================================
Function Wavelength(fStr As String, SoundSpeed As Long)
f = freqStr2Num(fStr)
Wavelength = SoundSpeed / f
End Function

'==============================================================================
' Name:     GetBandwidthIndex
' Author:   PS
' Desc:     Returns bandwidth index according to ANSI S1.11
' Args:     f - one third octave band nominal frequency
' Comments: (1) Simple, yeah?
'==============================================================================
Function GetBandwidthIndex(f As Double)
FrequencyArray = Array(1, 1.25, 1.6, 2, 2.5, 3.15, 4, 5, 6.3, 8, 10, 12.5, 16, 20, _
25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, _
1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000)
    For i = LBound(FrequencyArray) To UBound(FrequencyArray) + 1
        If FrequencyArray(i) = f Then
        GetBandwidthIndex = i
        Exit Function
        End If
    Next i
End Function

'==============================================================================
' Name:     FrequencyBandCutoff
' Author:   PS
' Desc:     Cutoff frequencies for band filters, as defined in ANSI S1.11:
'           Specification for Octave, Half-Octave, and Third Octave Band
'           Filter Sets
' Args:     fStr - centre frequency of band, Hz
'           Mode - "upper" or "lower"
'           Bandwidth - 1 and 3 for oct and 1/3 oct (defaults as 1/3 oct)
'           baseTen - boolean (defaults to TRUE)
' Comments: (1) Also known as the crossover frequency
'==============================================================================
Function FrequencyBandCutoff(fStr As String, Mode As String, _
Optional bandwidth As Integer, Optional baseTen As Boolean)

Dim G As Double 'gain
Dim f As Double 'frequency
Dim fr As Double 'reference frequency
Dim b As Integer 'bandwidth denominator
Dim x As Double


f = freqStr2Num(fStr)
fr = 1000
b = GetBandwidthIndex(f)

    If bandwidth = Empty Then bandwidth = 3 'default to one thirds
    
    'catch optional variable, default to Base 10
    If IsEmpty(baseTen) Then
    baseTen = True
    End If
    
    'set value of G for different modes
    If baseTen = True Then
    G = 10 ^ (3 / 10)
    Else 'baseten=false
    G = 2
    End If
    
    If b Mod 2 = 1 Then 'odd bandwidth number
    'If bandwidth Mod 2 = 1 Then 'odd
    x = Round(bandwidth * Application.WorksheetFunction.Log(f / fr) / _
        Application.WorksheetFunction.Log(G), 1)
    fm = fr * G ^ (x / bandwidth)
    Else 'even
    x = (2 * bandwidth * Application.WorksheetFunction.Log(f / fr) / _
        Application.WorksheetFunction.Log(G) - 1) / 2
    fm = fr * G ^ ((2 * x + 1) / (2 * bandwidth))
    End If

    'select mode: upper/lower
    If Mode = "upper" Then
    FrequencyBandCutoff = fm * G ^ (1 / (2 * bandwidth))
    ElseIf Mode = "lower" Then
    FrequencyBandCutoff = fm * G ^ (-1 / (2 * bandwidth))
    Else
    FrequencyBandCutoff = 0
    End If

End Function

'==============================================================================
' Name:     ExtractAddressElement
' Author:   PS
' Desc:     Get the first and last rows of  arange that's been input, assumes
'           the string contains relative references, and therefore '$' signs
' Args:     AddressStr (String of a range), elemNo (which element number to extract
' Comments: (1) Used for form frmBasic. A little hacky but it works.
'           (2) Renamed from ExtractRefElement. Now used in Options Analysis too.
'==============================================================================
Function ExtractAddressElement(AddressStr As String, elemNo As Integer)
Dim SplitStr() As String
Dim CheckRow As String
SplitStr = Split(AddressStr, "$", Len(AddressStr), vbTextCompare)
    If elemNo <= UBound(SplitStr) Then
    CheckRow = SplitStr(elemNo)
        'catch trailing colon character
        If Right(CheckRow, 1) = ":" Then CheckRow = Left(CheckRow, Len(CheckRow) - 1)
    ExtractAddressElement = CheckRow
    End If
End Function

'==============================================================================
' Name:     MassAirMass
' Author:   PS
' Desc:
' Args:     m1 - mass of first element in kg/m2
'           m2 - mass of second element in kg/m2
'           CavitySpace - distance between the leaves, in mm
' Comments: (1)
'==============================================================================
Function MassAirMass(m1 As Double, m2 As Double, CavitySpace As Double, Optional vAirTemp As Variant)
Dim a As Double
Dim AirTemp As Long
a = 1 / (2 * Application.WorksheetFunction.Pi())
rho = 1.225 'constant for now
    If IsMissing(vAirTemp) Or vAirTemp = "" Then
    AirTemp = 20 'default to 20 degrees celsius
    Else
    AirTemp = CLng(vAirTemp)
    End If
c = SpeedOfSound(AirTemp, False)
D = CavitySpace / 1000 'convert to metres
    If m1 > 0 And m2 > 0 Then
    MassAirMass = a * ((rho * (c ^ 2) * (m1 + m2)) / (D * m1 * m2)) ^ (1 / 2)
    Else
    MassAirMass = 0
    End If
End Function


'==============================================================================
' Name:     RoomCorrection_Schultz
' Author:   JCD
' Desc:     Returns the Schultz Room Correction at the specified frequency for
'           a given length, width, height, distance
' Args:     l = room length in meters
'           w = room width in meters
'           h = room height in meters
'           r = distance from source in meters
'           fStr = frequency in Hz
'
' Comments: (1) Source: Schultz, ASHRAE Transactions 1983, 91(1), pp 124-153.
'           (2) Assumes a rectilinear room
'==============================================================================
Function RoomCorrection_Schultz(Length As Double, Width As Double, Height As Double, _
    DistanceFromSource As Double, fStr As String)
    
Dim Volume As Double ' room volume
Dim f As Double ' frequency

Volume = Length * Width * Height
f = freqStr2Num(fStr)
    
    'guard clause
    If DistanceFromSource <= 0 Then
    Exit Function
    End If

RoomCorrection_Schultz = -10 * Application.WorksheetFunction.Log10(DistanceFromSource) _
    - 5 * Application.WorksheetFunction.Log10(Volume) _
    - 3 * Application.WorksheetFunction.Log10(f) + 12
End Function

'==============================================================================
' Name:     RoomCorrection_Plantroom
' Author:   PS
' Desc:     Returns a correction to calculate Lp from Lw, based on RT of the room
' Args:     l = room length in meters
'           w = room width in meters
'           h = room height in meters
'           RT = Reverberation time of the room
'
' Comments: (1) Assumes a rectilinear room
'==============================================================================
Function RoomCorrection_Plantroom(Length As Double, Width As Double, _
Height As Double, RT As Double)
    
Dim Volume As Double ' room volume

Volume = Length * Width * Height

RoomCorrection_Plantroom = -10 * Application.WorksheetFunction.Log10(Volume) + _
    10 * Application.WorksheetFunction.Log10(RT) + 14
End Function

'==============================================================================
' Name:     RoughRT
' Author:   PS
' Desc:     Parallel box method, integrated into area correction
' Args:
' Comments: (1) Source for these RTs?
'==============================================================================
Function RoughRT(fStr As String, RT_at500Hz As Double)
'Dim Masonry_RT(8) As Double
'Dim SubstantalLF_RT(8) As Double

'                  63  125  250 500  1k  2k   4k   8k
Masonry_RT = Array(1#, 1.2, 1.1, 1#, 1#, 0.9, 0.7, 0.5)
SubstantalLF_RT = Array(0.7, 0.8, 0.9, 1#, 1#, 0.9, 0.7, 0.5)

i = GetArrayIndex_OCT(fStr)

RoughRT = Masonry_RT(i) * RT_at500Hz

End Function

'==============================================================================
' Name:     MassLaw
' Author:   PS
' Desc:     Calculates mass law
' Args:     SurfaceDensity
' Comments: (1)
'==============================================================================
Function MassLaw(fStr As String, SurfaceDensity As Double)
freq = freqStr2Num(fStr)
MassLaw = 20 * Application.WorksheetFunction.Log10(freq * SurfaceDensity) - 48
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'==============================================================================
' Name:     GetFrequencyRange
' Author:   PS
' Desc:     Returns selected ID of dropdown control
' Args:     control - from dropdown on ribbon
'           freqID - ID string of selected item
' Comments: (1)
'==============================================================================
Sub GetFrequencyRange(control As IRibbonControl, ByRef freqID As Variant)
    If IsEmpty(T_FreqRange) Then
    T_FreqRange = "FreqRange0" 'default value
    End If
freqID = T_FreqRange
End Sub

'==============================================================================
' Name:     SetFrequencyRange
' Author:   PS
' Desc:     Set start and end frequencies for working range
' Args:     control - from dropdown on ribbon
'           id - string of freq type
'           index - item number, from 0
' Comments: (1) calls SwitchFrequencyColumns
'==============================================================================
Sub SetFrequencyRange(control As IRibbonControl, id As String, index As Integer)
'Dim ribUI As IRibbonUI
'Debug.Print control.id; vbTab; control.Tag

'Debug.Print id; index

SetSheetTypeControls

T_FreqRange = id
WorkingFreqRanges id

    'check for validation, apply if needed
    If HasDataValidation(Cells(T_FreqRow, T_LossGainStart)) = False Then
    ApplyFreqValidation
    End If
    
SwitchFrequencyColumns T_LossGainStart, T_LossGainEnd
    If T_SheetType = "MECH" Then
    SwitchFrequencyColumns T_RegenStart, T_RegenEnd
    End If
    
End Sub

'==============================================================================
' Name:     SwitchFrequencyColumns
' Author:   PS
' Desc:     Set start and end frequencies for working range
' Args:     ColStart - First column
'           ColEnd - Last column
' Comments: (1)
'==============================================================================
Sub SwitchFrequencyColumns(ColStart As Integer, ColEnd As Integer)

Dim SwitchCol As Integer
Dim OptionStr
Dim ValidationForm As String

    'loop through octave bands and set columns on or off
    For SwitchCol = ColStart To ColEnd
    'convert to number
    f = freqStr2Num(Cells(T_FreqRow, SwitchCol).Value)
    ValidationForm = Cells(T_FreqRow, SwitchCol).Validation.Formula1
    OptionStr = Split(ValidationForm, ",")
        If f < T_FreqStart Or f > T_FreqEnd Then 'exclude
        Cells(T_FreqRow, SwitchCol) = OptionStr(1)
        Else 'include
        Cells(T_FreqRow, SwitchCol) = OptionStr(0)
        End If
    Next SwitchCol
    
End Sub


'==============================================================================
' Name:     WorkingFreqRanges
' Author:   PS
' Desc:     Maps range options depending on SheetType, and sets dropdown options
' Args:     FreqRange as string
' Comments: (1)
'==============================================================================
Sub WorkingFreqRanges(FreqRange As String)
    
Select Case FreqRange
    Case Is = "FreqRange0" 'full spectrum
    T_FreqStart = freqStr2Num(Cells(T_FreqRow, T_LossGainStart).Value)
    T_FreqEnd = freqStr2Num(Cells(T_FreqRow, T_LossGainEnd).Value)
    Case Is = "FreqRange1"
    T_FreqStart = 63
    T_FreqEnd = 8000
    Case Is = "FreqRange2"
    T_FreqStart = 63
    T_FreqEnd = 4000
    Case Is = "FreqRange3"
    T_FreqStart = 125
    T_FreqEnd = 8000
    Case Is = "FreqRange4"
    T_FreqStart = 125
    T_FreqEnd = 4000
    Case Is = "FreqRangeCustom"
    T_FreqStart = freqStr2Num(Cells(T_FreqRow, Selection.Column).Value)
    T_FreqEnd = freqStr2Num(Cells(T_FreqRow, _
        Selection.Column + Selection.Columns.Count - 1).Value)
End Select

End Sub

'==============================================================================
' Name:     InsertBasicFunction
' Author:   PS
' Desc:     Inserts function, based on the user inputs in frmBasic
' Args:     functionName (depending on which function gets selected from the
'           ribbon.
' Comments: (1) currently supported functions: SUM, SPLSUM, SPLAV, SPLMINUS,
'           SPLSUMIF, SPLAVIF
'==============================================================================
Sub InsertBasicFunction(functionName As String)
Dim FirstRow As String
Dim LastRow As String
Dim FirstRow2 As String
Dim LastRow2 As String
Dim ColumnLetter As String
Dim NeedsTwoRanges As Boolean
Dim MarkerSymbol As String

    Select Case functionName
    Case Is = "SUM"
    frmBasicFunctions.optSum.Value = True
    MarkRowAs ("SUM")
    Case Is = "SPLSUM"
    frmBasicFunctions.optSPLSUM.Value = True
    MarkRowAs ("SUM")
    Case Is = "SPLAV"
    frmBasicFunctions.optSPLAV.Value = True
    MarkRowAs ("AV")
    Case Is = "SPLMINUS"
    frmBasicFunctions.optSPLMINUS.Value = True
    'TODO: minus
    Case Is = "SPLSUMIF"
    frmBasicFunctions.optSPLSUMIF.Value = True
    MarkRowAs ("SUM")
    Case Is = "SPLAVIF"
    frmBasicFunctions.optSPLAVIF.Value = True
    MarkRowAs ("AV")
    Case Is = "AV"
    frmBasicFunctions.optAverage.Value = True
    MarkRowAs ("AV")
    End Select

frmBasicFunctions.chkApplyToSheetType.Caption = "Apply for Sheet Type: " & _
    T_SheetType

frmBasicFunctions.Show

If btnOkPressed = False Then End
    
    'check for a secondary range, which is needed for some functions
    If functionName = "SPLMINUS" Or _
        functionName = "SPLSUMIF" Or _
        functionName = "SPLAVIF" Then
        
    NeedsTwoRanges = True
    
        If Range2Selection = "" Then
        msg = MsgBox("Error - you must select a secondary Range", _
            vbOKOnly, "Two is better than one.")
        End  'if no ranges selected then skip
        Else 'get rows for the other range
        FirstRow2 = ExtractAddressElement(Range2Selection, 2)
        LastRow2 = ExtractAddressElement(Range2Selection, 4)
        End If
        
    End If
   
'set description
SetDescription BasicFunctionType
Cells(Selection.Row, 1).Value = MarkerSymbol 'symbol in first column

    'build formula
    If ApplyToSheetType = True Then
    FirstRow = ExtractAddressElement(RangeSelection, 2)
    LastRow = ExtractAddressElement(RangeSelection, 4)
    ColumnLetter = ColNum2Str(T_LossGainStart)
        If NeedsTwoRanges = True Then
        'note, only single line inputs for functions with two ranges
        BuildFormula "" & BasicFunctionType & _
            "(" & ColumnLetter & FirstRow & "," & _
            ColumnLetter & FirstRow2 & ")"
        Else
        BuildFormula "" & BasicFunctionType & _
            "(" & ColumnLetter & FirstRow & ":" & ColumnLetter & LastRow & ")"
        End If

    Else
    BuildFormula "" & BasicFunctionType & _
        "(" & RangeSelection & ")"
    End If


    'apply style
    If BasicsApplyStyle <> "" Then
    SetTraceStyle BasicsApplyStyle
    End If

End Sub

'==============================================================================
' Name:     BandCutoff
' Author:   PS
' Desc:     Inserts band cutoff formula for SheetType
' Args:     None
' Comments: (1) includes code for setting up frmFrequencyBandCutoff
'==============================================================================
Sub BandCutoff()
    'set default values in the form, based on the Sheet Type
    If T_SheetType = "TO" Or T_SheetType = "TOA" Or T_SheetType = "LF_TO" Then
    frmFrequencyBandCutoff.optBand3 = True
    Else
    frmFrequencyBandCutoff.optBand1 = True
    End If
frmFrequencyBandCutoff.Show
If btnOkPressed = False Then End
SetDescription "Frequency Band Cutoff (" & FBC_mode & ", Hz)"
BuildFormula "FrequencyBandCutoff(" & _
    T_FreqStartRng & ",""" & FBC_mode & """," & FBC_bandwidth & "," & _
    FBC_baseTen & ")"
End Sub

'==============================================================================
' Name:     PutWavelength
' Author:   PS
' Desc:     Inserts formula for wavelength & speed of sound (in parameter col)
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutWavelength()

SetDescription "Wavelength (m)"

Cells(Selection.Row, T_ParamStart + 1).Value = 20 'default to 20 degrees celcius
Cells(Selection.Row, T_ParamStart).Value = "=SpeedOfSound(" & T_ParamRng(1) & ")"
BuildFormula "Wavelength(" & T_FreqStartRng & "," _
    & T_ParamRng(0) & ")"

'Formatting
Range(Cells(Selection.Row, T_LossGainStart), _
    Cells(Selection.Row, T_LossGainEnd)).NumberFormat = "0.00"
SetUnits "mps", T_ParamStart
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """""0""°C """
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutSpeedofSound
' Author:   PS
' Desc:     Inserts formula for wavelength & speed of sound (in parameter col)
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutSpeedOfSound()

SetDescription "Speed of Sound"

Cells(Selection.Row, T_ParamStart + 1).Value = 20 'default to 20 degrees celcius
Cells(Selection.Row, T_ParamStart).Value = "=SpeedOfSound(" & T_ParamRng(1) & ")"

'Formatting
SetUnits "mps", T_ParamStart
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """""0""°C """
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutMassAirMass
' Author:   PS
' Desc:     Inserts formula for mass-air-mass
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutMassAirMass()

SetDescription "Mass-Air-Mass"

frmMAM.Show
If btnOkPressed = False Then End

Cells(Selection.Row, T_ParamStart).Value = MAM_Width
Cells(Selection.Row, T_ParamStart + 1).Value = "=MassAirMass(" & MAM_M1 & "," & _
    MAM_M2 & "," & Cells(Selection.Row, T_ParamStart).Address(False, False) & ",20)"
InsertComment MAM_Description, T_ParamStart + 1

'Formatting
SetUnits "mm", T_ParamStart
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = "0.0 ""Hz"""
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutRCSchultz
' Author:   PS
' Desc:     Inserts formula for the Schultz approximation
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRCSchultz()
SetDescription "Room Loss (Schultz)"

'=RoomCorrection_Schultz(Length,Width,Height,DistanceFromSource,fStr)

frmSchultz.Show
If btnOkPressed = False Then End

Cells(Selection.Row, T_ParamStart).Value = DistanceFromSource
BuildFormula "RoomCorrection_Schultz(" & roomL & "," & _
    roomW & "," & roomH & "," & T_ParamRng(0) & "," & T_FreqStartRng & ")"
InsertComment "Distance to source, m", T_ParamStart, False

'Formatting
ParameterMerge Selection.Row
SetUnits "m", T_ParamStart
SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutRCPlantroom
' Author:   PS
' Desc:     Inserts formula for the Plantroom Room Correction approximation
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRCPlantroom()
SetDescription "Room Loss (Plantrooms)"

ParameterMerge Selection.Row
Cells(Selection.Row, T_ParamStart).Value = 36
SetUnits "m3", T_ParamStart
SetTraceStyle "Input", True

BuildFormula "-10*log(" & T_ParamRng(0) & ")+10*log(" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & ")+14"
InsertComment "Room volume, m" & chr(179), T_ParamStart, False

End Sub
