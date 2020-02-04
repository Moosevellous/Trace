Attribute VB_Name = "Basics"
Dim ConditionValue As Variant
Public BasicFunctionType As String
Public RangeSelection As String
Public Range2Selection As String
Public FitToSheetType As Boolean
Public BasicsApplyStyle As String

Function freqStr2Num(fstr) As Double
    Select Case fstr
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
    Case Else
    freqStr2Num = 0
    End Select
End Function


Function GetOctaveColumnIndex(freq)
    Select Case freq
    Case Is = "63"
    GetOctaveColumnIndex = 0
    Case Is = "125"
    GetOctaveColumnIndex = 1
    Case Is = "250"
    GetOctaveColumnIndex = 2
    Case Is = "500"
    GetOctaveColumnIndex = 3
    Case Is = "1k"
    GetOctaveColumnIndex = 4
    Case Is = "2k"
    GetOctaveColumnIndex = 5
    Case Is = "4k"
    GetOctaveColumnIndex = 6
    Case Is = "8k"
    GetOctaveColumnIndex = 7
    Case Is = 1000
    GetOctaveColumnIndex = 4
    Case Is = 2000
    GetOctaveColumnIndex = 5
    Case Is = 4000
    GetOctaveColumnIndex = 6
    Case Is = 8000
    GetOctaveColumnIndex = 7
    Case Else
    GetOctaveColumnIndex = 999 'for catching errors
    End Select
End Function

Public Function SPLSUM(ParamArray rng1() As Variant) As Variant
On Error Resume Next

Dim C As Range
Dim i As Long

SPLSUM = -99
For i = LBound(rng1) To UBound(rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(rng1(i)) = "Double" Then
        If rng1(i) > 0 Then 'negative values are ignored
        SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (rng1(i) / 10)))
        End If
    ElseIf TypeName(rng1(i)) = "Range" Then
        For Each C In rng1(i).Cells
            If C.Value <> Empty Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (C.Value / 10)))
            End If
        Next C
    End If
Next i

''check for negative
'If SPLSUM < -5 Then
'SPLSUM = ""
'End If

End Function

Public Function SPLAV(ParamArray rng1() As Variant) As Variant
On Error Resume Next

Dim C As Range
Dim i As Long
Dim N As Integer
SPLAV = -99
N = 0
For i = LBound(rng1) To UBound(rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(rng1(i)) = "Double" Then
        If rng1(i) > 0 Then 'negative values are ignored
        SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (rng1(i) / 10)))
        N = N + 1
        End If
    ElseIf TypeName(rng1(i)) = "Range" Then
        For Each C In rng1(i).Cells
            If C.Value <> Empty And IsNumeric(C.Value) Then
            SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (C.Value / 10)))
            N = N + 1
            End If
        Next C
    End If
Next i

'Average +10log(1/n) in log domain
SPLAV = SPLAV + 10 * Application.WorksheetFunction.Log10(1 / N)

''check for negative
'If SPLSUM < -5 Then
'SPLSUM = ""
'End If

End Function

Public Function SPLMINUS(SPLtotal As Double, SPL2 As Double) As Variant
On Error Resume Next

SPLMINUS = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLtotal / 10)) - (10 ^ (SPL2 / 10)))

End Function

Public Function TL_ThirdsToOctave(rngInput As Range) As Variant

Dim isNegative As Boolean
Dim TL1 As Single, TL2 As Single, TL3 As Single

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
    TL_ThirdsToOctave = 10 * Application.WorksheetFunction.Log10((1 / 3) * ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    Else 'return result as positive
    TL_ThirdsToOctave = -10 * Application.WorksheetFunction.Log10((1 / 3) * ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    End If

End Function

Public Function CompositeTL(TL_Range As Range, AreaRange As Range) As Variant

Dim TotalArea As Double
Dim TL_Calc_Range() As Double

'original formula
'=-10*LOG(SUM($N$42:$O$45)/(($N$42*(10^(F$42/10)))+($N$43*(10^(F$43/10)))+($N$44*(10^(F$44/10)))+($N$45*(10^(F$45/10)))))

'Debug.Print TypeName(TL_Range)
TotalArea = 0

    For a = 1 To AreaRange.Rows.Count '<- TODO for each???
        If AreaRange(a, 1).Value <> "" Then
        TotalArea = TotalArea + AreaRange(a, 1)
        End If
    Next a

    For i = 1 To TL_Range.Rows.Count
    'Debug.Print TL_Range(i, 1).Text
    ReDim Preserve TL_Calc_Range(i)
        If TL_Range(i, 1).Text = "" Then
        TL_Calc_Range(i) = -99 'how low can you go!
        Else
        TL_Calc_Range(i) = TL_Range(i, 1) 'how low can you go!
        End If
    Next i

WeightedSum = 0
    For elem = 0 To UBound(TL_Calc_Range)
    'add 'em up!
    Next elem
    
CompositeTL = -1 * 10 * Application.WorksheetFunction.Log(TotalArea / (Area1 * 10 ^ (TL1 / 10) + Area2 * 10 ^ (TL2 / 10) + Area3 * 10 ^ (TL3 / 10) + Area3 * 10 ^ (TL3 / 10) + _
Area4 * 10 ^ (TL4 / 10) + Area5 * 10 ^ (TL5 / 10) + Area6 * 10 ^ (TL6 / 10) + Area7 * 10 ^ (TL7 / 10) + Area8 * 10 ^ (TL8 / 10)))

End Function

Function SPLSUMIF(SumRange As Range, Condition As String, Optional ConditionRange As Variant)

Dim IfRange As Range
Dim TypeFound As Boolean

    'Check which Range will be evaluating the IF function
    If IsMissing(ConditionRange) Then
    Set IfRange = SumRange
    Else
    Set IfRange = ConditionRange
    End If

ConditionType = FindConditionType(Condition)
If ConditionType = "" Then TypeFound = False
    
'initialise function
SPLSUMIF = -99

    For Each C In IfRange.Cells
    
'    Debug.Print "row: "; C.Row; "column: "; C.Column
'    Debug.Print "Condition test: "; ConditionType; " "; C.Value
'    Debug.Print "Cell value: "; SumRange(C.Row, C.Column).Value
'    Debug.Print ""
    
    rw = IfRange.Row - SumRange.Row
    clmn = IfRange.Column - SumRange.Column
    
        Select Case ConditionType
        Case Is = "GreaterThan"
            If C.Value > ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUMIF / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            End If
        Case Is = "GreaterThanEqualTo"
            If C.Value >= ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUMIF / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            End If
        Case Is = "LessThan"
            If C.Value < ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUMIF / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            End If
        Case Is = "LessThanEqualTo"
            If C.Value <= ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUMIF / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            End If
        Case Is = "Equals"
            If C.Value = ConditionValue Then
            SPLSUMIF = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUMIF / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            End If
        Case Is = "" 'no condtion type
            SPLSUMIF = -99
        End Select
    Next C

End Function

Function SPLAVIF(SumRange As Range, ConditionStr As String, Optional ConditionRange As Variant)

Dim IfRange As Range
Dim TypeFound As Boolean
Dim numVals As Integer
Dim ConditionType As String
Dim SPLSUM As Single

    'Check which Range will be evaluating the IF function
    If IsMissing(ConditionRange) Then
    Set IfRange = SumRange
    Else
    Set IfRange = ConditionRange
    End If

ConditionType = FindConditionType(ConditionStr)
If ConditionType = "" Then TypeFound = False
    
'initialise function
SPLSUM = -99
SPLAVIF = -99
numVals = 0

    For Each C In IfRange.Cells
    
'    Debug.Print "row: "; C.Row; "column: "; C.Column
'    Debug.Print "Condition test: "; ConditionType; " "; C.Value
'    Debug.Print "Cell value: "; SumRange(C.Row, C.Column).Value
'    Debug.Print ""
    
    rw = IfRange.Row - SumRange.Row
    clmn = IfRange.Column - SumRange.Column
    
        Select Case ConditionType
        Case Is = "GreaterThan"
            If C.Value > ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "GreaterThanEqualTo"
            If C.Value >= ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "LessThan"
            If C.Value < ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "LessThanEqualTo"
            If C.Value <= ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "Equals"
            If C.Value = ConditionValue Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Cells(C.Row - rw, C.Column - clmn).Value / 10)))
            numVals = numVals + 1
            End If
        Case Is = "" 'no condtion type
            SPLSUM = -99
        End Select
    Next C
    
'Debug.Print numVals; "Values:"
SPLAVIF = SPLSUM - (10 * Application.WorksheetFunction.Log(numVals)) '-10log(n) to average

End Function

Function FindConditionType(inputFormula As String)
  
    If Left(inputFormula, 2) = ">=" Then
    FindConditionType = "GreaterThanEqualTo"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 2))
    TypeFound = True
    ElseIf Left(inputFormula, 2) = "<=" Then
    FindConditionType = "LessThanEqualTo"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 2))
    TypeFound = True
    ElseIf Left(inputFormula, 1) = "=" Then
    FindConditionType = "Equals"
    ConditionValue = Right(inputFormula, Len(inputFormula) - 1)
    TypeFound = True
    ElseIf Left(inputFormula, 1) = "<" Then
    FindConditionType = "LessThan"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 1))
    TypeFound = True
    ElseIf Left(inputFormula, 1) = ">" Then
    FindConditionType = "GreaterThan"
    ConditionValue = CSng(Right(inputFormula, Len(inputFormula) - 1))
    TypeFound = True
    Else 'catch all, no equals
    FindConditionType = ""
    TypeFound = False
    End If
    
End Function


Public Function FitzroyRT(X As Long, Y As Long, Z As Long, S_i As Range, Direction As Range, alpha_i As Range)

Dim a_x As Single 'a_x is alpha-bar x, ie the averabe absorption for surfaces in the x direction
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

Volume = X * Y * Z

'Debug.Print "ax:"; a_x; "   ay:"; a_y; "   az"; a_z

FitzroyRT = (0.161 * Volume / S_total ^ 2) * _
(((-Sx_total / Application.WorksheetFunction.Ln(1 - a_x)) + _
(-Sy_total / Application.WorksheetFunction.Ln(1 - a_y)) + _
(-Sz_total / Application.WorksheetFunction.Ln(1 - a_z))))

End Function

Function GetSpeedOfSound(temp As Long, Optional IsKelvin As Boolean)
    If IsKelvin = False Then 'convert to kelvin, not hobbs
    temp = temp + 273.15
    End If
GetSpeedOfSound = (1.4 * 287.1848 * temp) ^ 0.5 'square root of Gamma * R * Temp for air
End Function

Function GetWavelength(fstr As String, SoundSpeed As Long)
f = freqStr2Num(fstr)
GetWavelength = SoundSpeed / f
End Function

Function FrequencyBandCutoff(freq As String, Mode As String, Optional bandwidth As Double, Optional baseTen As Boolean)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'As Specified in:
'ANSI S1.11: Specification for Octave, Half-Octave, and Third Octave Band Filter Sets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim G As Double
Dim f As Single

f = freqStr2Num(freq)

    If bandwidth = Empty Then bandwidth = 3 'default to one thirds?
    
    If baseTen = Empty Then 'catch optional variable
    baseTen = True
    Else
    baseTen = False
    End If

    If baseTen = True Then
    G = 10 ^ (3 / 10)
    Else
    G = 2
    End If

fr = 1000
    
    If bandwidth Mod 2 = 1 Then 'odd
    X = Round(bandwidth * Application.WorksheetFunction.Log(f / fr) / Application.WorksheetFunction.Log(G), 1)
    fm = fr * G ^ (X / bandwidth)
    Else 'even
    X = (2 * bandwidth * Application.WorksheetFunction.Log(f / fr) / Application.WorksheetFunction.Log(G) - 1) / 2
    fm = fr * G ^ ((2 * X + 1) / (2 * bandwidth))
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


Function RangeAddressSheetExtents(AddrStr As String, SheetType As String) As String()
Dim Addresses()  As String

splitStr = Split(AddrStr, ",", Len(AddrStr), vbTextCompare)



End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub InsertBasicFunction(SheetType As String, functionName As String)
Dim FirstRow As String
Dim LastRow As String

    Select Case functionName
    Case Is = "SUM"
    frmBasicFunctions.optSum.Value = True
    Case Is = "SPLSUM"
    frmBasicFunctions.optSPLSUM.Value = True
    Case Is = "SPLAV"
    frmBasicFunctions.optSPLAV.Value = True
    Case Is = "SPLMINUS"
    frmBasicFunctions.optSPLMINUS.Value = True
    Case Is = "SPLSUMIF"
    frmBasicFunctions.optSPLSUMIF.Value = True
    Case Is = "SPLAVIF"
    frmBasicFunctions.optSPLAVIF.Value = True
    End Select

frmBasicFunctions.chkApplyToSheetType.Caption = "Apply for Sheet Type: " & SheetType

frmBasicFunctions.Show

If btnOkPressed = False Then End

    If functionName = "SPLMINUS" Or functionName = "SPLSUMIF" Or functionName = "SPLAVIF" Then 'no need to check
        If Range2Selection = "" Then
        msg = MsgBox("Error - you must select a secondary Range", vbOKOnly, "Two is better than one.")
        End  'if no ranges selected then skip
        End If
    End If

    If FitToSheetType = True Then
    FirstRow = ExtractRefElement(RangeSelection, 2)
    LastRow = ExtractRefElement(RangeSelection, 4)
    End If

    If Left(SheetType, 3) = "OCT" Then
        If FitToSheetType = True Then
        Cells(Selection.Row, 5).Value = "=" & BasicFunctionType & "(E" & FirstRow & "E" & LastRow & ")"
        Else
        Cells(Selection.Row, 5).Value = "=" & BasicFunctionType & "(" & RangeSelection & ")"
        End If
        'Cells(Selection.Row, 5).Value = "=" & BasicFunctionType & "(E$6,N" & Selection.Row & ")"
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=" & BasicFunctionType & "(E$6,Z" & Selection.Row & ")"
    Cells(Selection.Row, 5).Value = "=" & BasicFunctionType & "(" & RangeSelection & ")"
    Else
    ErrorOctOnly
    End If

Cells(Selection.Row, 2).Value = functionName

ExtendFunction (SheetType)

    'apply style
    If BasicsApplyStyle <> "" Then
    ApplyTraceStyle "Trace " & BasicsApplyStyle, SheetType, Selection.Row
    End If

End Sub


Sub BandCutoff(SheetType As String)

End Sub

Function ExtractRefElement(AddressStr As String, elemNo As Integer)
Dim splitStr() As String
splitStr = Split(AddressStr, "$", Len(AddressStr), vbTextCompare)
    If elemNo <= UBound(splitStr) Then
    ExtractRefElement = splitStr(elemNo)
    End If
End Function

Sub Wavelength(SheetType As String)
Dim Col As Integer
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "Wavelength"
Cells(Selection.Row, 3).Value = ""
Cells(Selection.Row, 4).Value = ""

    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 14).Value = "=GetSpeedOfSound($O" & Selection.Row & ")"
    Cells(Selection.Row, 5).Value = "=GetWavelength(E$6,$N" & Selection.Row & ")"
    ParamCol1 = 14
    ParamCol2 = 15
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)).NumberFormat = "0.00"
    
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 26).Value = "=GetSpeedOfSound($AA" & Selection.Row & ")"
    Cells(Selection.Row, 5).Value = "=GetWavelength(E$6,$Z" & Selection.Row & ")"
    ParamCol1 = 26
    ParamCol2 = 27
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 25)).NumberFormat = "0.00"
    
    ElseIf Left(SheetType, 5) = "LF_TO" Then
    Cells(Selection.Row, 32).Value = "=GetSpeedOfSound($AG" & Selection.Row & ")"
    Cells(Selection.Row, 5).Value = "=GetWavelength(E$6,$AF" & Selection.Row & ")"
    ParamCol1 = 32
    ParamCol2 = 33
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 31)).NumberFormat = "0.0"
    
    Else
    ErrorTypeCode
    End If

Cells(Selection.Row, ParamCol2).Value = 20 'default to 20 degrees celcius

ExtendFunction (SheetType)
'Formatting
Cells(Selection.Row, ParamCol1).NumberFormat = """""#""m/s"";""""# ""m/s"""
Cells(Selection.Row, ParamCol2).NumberFormat = """""0""°C """
fmtUserInput SheetType, True

End Sub



Sub SpeedOfSound(SheetType As String)

End Sub

