Attribute VB_Name = "Curves"
Dim RowToCheck As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     AWeightCorrections
' Author:   PS
' Desc:     Returns the A-weighting corrections in one-thrd octave bands
' Args:     fStr (frequency band)
' Comments: (1) From 10Hz to 20kHz
'==============================================================================
Function AWeightCorrections(fstr As String)
Dim dBAAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fstr)

ArrayIndex = 999 'for error catching

'                     Corrections for 1/3 octave band centre frequencies
'                     10     12.5   16     20     25     31.5   40     50     63
'   80     100    125    160    200    250   315   400   500   630   800 1000
'   1250 1600 2000 2500 3150 4000 5000 6300 8000  10000 12500 16000 20000)
dBAAdjustment = Array(-70.4, -63.4, -56.7, -50.5, -44.7, -39.4, -34.6, -30.2, -26.2, _
    -22.5, -19.1, -16.1, -13.4, -10.9, -8.6, -6.6, -4.8, -3.2, -1.9, -0.8, 0#, 0.6, 1#, _
    1.2, 1.3, 1.2, 1#, 0.5, -0.1, -1.1, -2.5, -4.3, -6.6, -9.3)
freqTitles = Array(10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, _
    80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, _
    1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000)
    
'alternative methods!!!!????
'http://www.beis.de/Elektronik/AudioMeasure/WeightingFilters.html
'=10*LOG(((35041384000000000*f^8)/((20.598997^2+f^2)^2*(107.65265^2+f^2)*(737.86223^2+f^2)*(12194.217^2+f^2)^2)))
    
    For i = LBound(freqTitles) To UBound(freqTitles)
        If freq = freqTitles(i) Then
        ArrayIndex = i
        found = True
        End If
    Next i
    
    If ArrayIndex <> 999 Then 'error
    AWeightCorrections = dBAAdjustment(ArrayIndex)
    Else
    AWeightCorrections = "-"
    End If
    
End Function

'==============================================================================
' Name:     CWeightCorrections
' Author:   PS
' Desc:     Returns the C-weighting corrections in one-thrd octave bands
' Args:     fStr (frequency band)
' Comments: (1) From 10Hz to 20kHz
'==============================================================================
Function CWeightCorrections(fstr As String)
Dim dBCAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fstr)

ArrayIndex = 999 'for error catching

'                     Corrections for 1/3 octave band centre frequencies
'                     10     12.5   16     20     25     31.5   40     50     63
'   80     100    125    160    200    250   315   400   500   630   800 1000
'   1250 1600 2000 2500 3150 4000 5000 6300 8000  10000 12500 16000 20000)
dBCAdjustment = Array(-14.3, -11.2, -8.5, -6.2, -4.4, -3.1, -2#, -1.3, -0.8, -0.5, _
    -0.3, -0.2, -0.1, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    -0.1, -0.2, -0.3, -0.5, -0.8, -1.3, -2#, -3#, -4.4, -6.2, -8.5, -11.2)
freqTitles = Array(10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, _
    80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, _
    1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000)


    For i = LBound(freqTitles) To UBound(freqTitles)
        If freq = freqTitles(i) Then
        ArrayIndex = i
        End If
    Next i
    
    If ArrayIndex <> 999 Then 'error
    CWeightCorrections = dBCAdjustment(ArrayIndex)
    Else
    CWeightCorrections = "-"
    End If
    
End Function

'==============================================================================
' Name:     GWeightCorrections
' Author:   PS
' Desc:     Returns the G-weighting corrections in one-thrd octave bands
' Args:     fStr (frequency band)
' Comments: (1) From 0.25Hz to 100Hz
'           (2) Reference to ISO 7196-1995
'           (3) Updated freqTitles array, was missing one value (!)
'==============================================================================
Function GWeightCorrections(fstr As String)
Dim dBGAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fstr)

ArrayIndex = 999 'for error catching

'                     Corrections for 1/3 octave band centre frequencies
'                    0.25 0.315  0.4    0.5    0.63   0.8   1.0  1.25   1.63    2
'   2.5   3.15  4.0  5.0 6.3  8  10 12.5 16 20  25  31.5  40   50   63   80   100
dBGAdjustment = Array(-88, -80, -72.1, -64.3, -56.6, -49.5, -43, -37.5, -32.6, -28.3, _
    -24.1, -20, -16, -12, -8, -4, 0, 4, 7.7, 9, 3.7, -4, -12, -20, -28, -36, -44)
freqTitles = Array(0.25, 0.315, 0.4, 0.5, 0.63, 0.8, 1#, 1.25, 1.6, 2, _
    2.5, 3.15, 4, 5, 6.3, 8, 10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100)

    For i = LBound(freqTitles) To UBound(freqTitles)
        If freq = freqTitles(i) Then
        ArrayIndex = i
        End If
    Next i
    
    If ArrayIndex <> 999 Then 'error
    GWeightCorrections = dBGAdjustment(ArrayIndex)
    Else
    GWeightCorrections = "-"
    End If
    
End Function

'==============================================================================
' Name:     NRcurve
' Author:   PS
' Desc:     Returns the values of the NR curve, at CurveNo
' Args:     CurveNo (name/number of the curve) fStr (frequency band)
' Comments: (1) Based on AS1469 - Acoustics - Methods for the determination of
'           noise rating numbers
'==============================================================================
Function NRcurve(CurveNo As Integer, fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim Ifreq As Integer
Dim freq As Double
freq = freqStr2Num(fstr)

    If freq < 31 Or freq > 8000 Then
    NRcurve = "-"
    Exit Function
    End If

'coefficients from Table 1 of AS1469
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)
'''''''''''''''''''''''''''''''''
Ifreq = GetArrayIndex_OCT(fstr, 1)

NRcurve = A_f(Ifreq) + (B_f(Ifreq) * CurveNo)
End Function

'==============================================================================
' Name:     PNCcurve
' Author:   PS
' Desc:     Returns the value of the PNC curve at that frequency band
' Args:     fStr, the frequency band centre frequency as a string, CurveNo the
'           index number of the curve
' Comments: (1) Based on someone else's code I think?
'==============================================================================
Function PNCcurve(CurveNo As Integer, fstr As String)

Dim DataTable(0 To 10, 0 To 8) As Double
Dim IStart As Integer
Dim Ifreq As Integer
Dim freq As Double

freq = freqStr2Num(fstr) 'for checking only

    If freq < 31 Or freq > 8000 Then
    PNCcurve = "-"
    Exit Function
    End If

'define curves
'bands       31.5 63  125 250 500 1k  2k  4k 8k
PNC15 = Array(58, 43, 35, 28, 21, 15, 10, 8, 8)
PNC20 = Array(59, 46, 39, 32, 26, 20, 15, 13, 13)
PNC25 = Array(60, 49, 43, 37, 31, 25, 20, 18, 18)
PNC30 = Array(61, 52, 46, 41, 35, 30, 25, 23, 23)
PNC35 = Array(62, 55, 50, 45, 40, 35, 30, 28, 28)
PNC40 = Array(64, 59, 54, 50, 45, 40, 36, 33, 33)
PNC45 = Array(67, 63, 58, 54, 50, 45, 41, 38, 38)
PNC50 = Array(70, 66, 62, 58, 54, 50, 46, 43, 43)
PNC55 = Array(73, 70, 66, 62, 59, 55, 51, 48, 48)
PNC60 = Array(76, 73, 69, 66, 63, 59, 56, 53, 53)
PNC65 = Array(79, 76, 73, 70, 67, 64, 61, 58, 58)

    For i = 0 To 8
    DataTable(0, i) = PNC15(i)
    DataTable(1, i) = PNC20(i)
    DataTable(2, i) = PNC25(i)
    DataTable(3, i) = PNC30(i)
    DataTable(4, i) = PNC35(i)
    DataTable(5, i) = PNC40(i)
    DataTable(6, i) = PNC45(i)
    DataTable(7, i) = PNC50(i)
    DataTable(8, i) = PNC55(i)
    DataTable(9, i) = PNC60(i)
    DataTable(10, i) = PNC65(i)
    Next i

'get array index, from 31.5Hz band
Ifreq = GetArrayIndex_OCT(fstr, 1)
    
    'select row of Data
    Select Case CurveNo
    Case Is = 15
    DataRow = 0
    Case Is = 20
    DataRow = 1
    Case Is = 25
    DataRow = 2
    Case Is = 30
    DataRow = 3
    Case Is = 35
    DataRow = 4
    Case Is = 40
    DataRow = 5
    Case Is = 45
    DataRow = 6
    Case Is = 50
    DataRow = 7
    Case Is = 55
    DataRow = 8
    Case Is = 60
    DataRow = 9
    Case Is = 65
    DataRow = 10
    End Select

PNCcurve = DataTable(DataRow, Ifreq)

End Function

'==============================================================================
' Name:     NR_rate
' Author:   PS
' Desc:     Rates an input spectrum against NR curves and returns the rating
' Args:     DataTable the spectrum to be rated, assumed to start at 31.5Hz band
'           fStr, the starting frequency
' Comments: (1) Based on AS1469 - Acoustics - Methods for the determination of
'           noise rating numbers
'==============================================================================
Function NR_rate(DataTable As Variant, Optional fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim NR_f, NR As Double
Dim NRTemp As Double
Dim IStart, Col As Integer

    If DataTable.Rows.Count <> 1 Then
        NR_rate = "ERROR!"
        Exit Function
    End If
NRTemp = 0

'coefficients from Table 1 of AS1469
'bands      31.5  63    125 250 500  1k 2k    4k    8k
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
'bands      31.5   63    125   250   500   1k  2k     4k     8k
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)

    'if no frequency input, assume data starts at 31.5Hz
    If fstr = "" Then fstr = "31.5"
    
'get array index, from 31.5Hz band
IStart = GetArrayIndex_OCT(fstr, 1)
    
    'Debug.Print DataTable.Columns.Count
    For Col = 1 To DataTable.Columns.Count
        If IsNumeric(DataTable(1, Col)) Then
            NR_f = (DataTable(1, Col) - A_f(IStart + Col - 1)) / _
                B_f(IStart + Col - 1)  'get the NR for that octave band
            If NR_f > NR Then 'if greater than highest NR found so far
                NR = NR_f
            End If
        End If
    Next Col
    
    If NR > 100 Then
        NR_rate = "OVER 100!"
        Exit Function
    End If
NR_rate = WorksheetFunction.RoundUp(NR, 0)
End Function

'==============================================================================
' Name:     NCcurve
' Author:   PS
' Desc:     Returns the value of the NC curve at input freuqency. Calls
'           InterpolateNCcurve is curve is not defined in ANSI S12.2
' Args:     CurveNo, the index number of the NC curve, fStr, the frequency band
'           as a string
' Comments: (1) Based on ANSI S12.2 2008
'==============================================================================
Function NCcurve(CurveNo As Integer, fstr As String)
Dim Ifreq As Integer
Dim freq As Integer

freq = freqStr2Num(fstr)

    If freq < 16 Or freq > 8000 Then 'catch frequencies out of range
    NCcurve = "-"
    Exit Function
    End If
    
'Define NC curves
'According to ANSI S12.2 2008
'            Octave band centre frequencies
'            16  31  63  125 250 500 1k  2k  4k  8k
NC70 = Array(90, 90, 84, 79, 75, 72, 71, 70, 68, 68)
NC65 = Array(90, 88, 80, 75, 71, 68, 65, 64, 63, 62)
NC60 = Array(90, 85, 77, 71, 66, 63, 60, 59, 58, 57)
NC55 = Array(89, 82, 74, 67, 62, 58, 56, 54, 53, 52)
NC50 = Array(87, 79, 71, 64, 58, 54, 51, 49, 48, 47)
NC45 = Array(85, 76, 67, 60, 54, 49, 46, 44, 43, 42)
NC40 = Array(84, 74, 64, 56, 50, 44, 41, 39, 38, 37)
NC35 = Array(82, 71, 60, 52, 45, 40, 36, 34, 33, 32)
NC30 = Array(81, 68, 57, 48, 41, 35, 32, 29, 28, 27)
NC25 = Array(80, 65, 54, 44, 37, 31, 27, 24, 22, 22)
NC20 = Array(79, 63, 50, 40, 33, 26, 22, 20, 17, 16)
NC15 = Array(78, 61, 47, 36, 28, 22, 18, 14, 12, 11)

'get array index, from 16Hz band
Ifreq = GetArrayIndex_OCT(fstr, 2)
    
    If CurveNo Mod 5 = 0 Then 'simply return the defined value
        Select Case CurveNo
        Case Is = 15
        ChosenCurve = NC15
        Case Is = 20
        ChosenCurve = NC20
        Case Is = 25
        ChosenCurve = NC25
        Case Is = 30
        ChosenCurve = NC30
        Case Is = 35
        ChosenCurve = NC35
        Case Is = 40
        ChosenCurve = NC40
        Case Is = 45
        ChosenCurve = NC45
        Case Is = 50
        ChosenCurve = NC50
        Case Is = 55
        ChosenCurve = NC55
        Case Is = 60
        ChosenCurve = NC60
        Case Is = 65
        ChosenCurve = NC65
        Case Is = 70
        ChosenCurve = NC70
        End Select
    NCcurve = ChosenCurve(Ifreq)
    Else 'interpolate between the curves
    NCcurve = InterpolateNCcurve(CurveNo, fstr)
    End If
    
End Function

'==============================================================================
' Name:     InterpolateNCcurve
' Author:   PS
' Desc:     Returns the value of the interpolated NC curve at input freuqency
' Args:     CurveNo, the index number of the NC curve,
'           fStr, the frequency band as a string
' Comments: (1) Based on ANSI S12.2 2008
'==============================================================================
Function InterpolateNCcurve(CurveNo As Integer, fstr As String)

Dim freq As Integer
Dim Remainder As Integer
Dim UpperCurveValue As Single
Dim LowerCurveValue As Single
Dim UpperCurve As Integer
Dim LowerCurve As Integer

freq = freqStr2Num(fstr)

Remainder = CurveNo Mod 5
'x values
UpperCurve = CurveNo + (5 - Remainder)
LowerCurve = CurveNo - Remainder
'y values
UpperCurveValue = NCcurve(UpperCurve, fstr)
LowerCurveValue = NCcurve(LowerCurve, fstr)

'interpolate linearly
m = (UpperCurveValue - LowerCurveValue) / (UpperCurve - LowerCurve)
InterpolateNCcurve = LowerCurveValue + (m * (CurveNo - LowerCurve))

End Function

'==============================================================================
' Name:     NCrate
' Author:   PS
' Desc:     Rates the input spectrum against the NC curve
' Args:     DataTable - the spectrum to be rated, assumed to start at 16Hz band
'           Optional fStr, the starting frequency band
' Comments: (1) Calls NCcurve for values
'==============================================================================
Function NCrate(DataTable As Variant, Optional fstr As String)

Dim NC As Double
Dim freq As Double
Dim IStart As Integer
Dim Col As Integer
Dim NCtemp As Integer
   
   If DataTable.Rows.Count <> 1 Then
        NCrate = "ERROR!"
        Exit Function
    End If

'bands
octaveBands = Array(16, 31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000)

    If fstr = "" Then
    freq = 16 'if no frequency input, assume data starts at 16Hz octave band
    Else
    freq = freqStr2Num(fstr)
    End If

'get array index, from 16Hz band
IStart = GetArrayIndex_OCT(fstr, 2)

    NCtemp = 15
    found = False
    SumExceedances = 0
    While found = False
    'Debug.Print "Checking NC"; i
    test_freq = octaveBands(IStart)
        For Col = 1 To DataTable.Columns.Count 'all input value
        test_freq = octaveBands(IStart + Col - 1) 'DataTable is indexed from 1, not 0
            If IsNumeric(DataTable(1, Col)) Then
            'get value of curve at that band
            NC_curve_value = NCcurve(NCtemp, CStr(test_freq))
            'Debug.Print DataTable(1, Col + 1).Value; "    NCvalue: "; NC_curve_value
                If DataTable(1, Col).Value > NC_curve_value Then
                SumExceedances = SumExceedances + (DataTable(1, Col) - NC_curve_value)
                End If
            End If
        Next Col
    
        'catch error
        If NCtemp > 70 Then
        found = True
        errnc = True
        ElseIf SumExceedances = 0 Then
        found = True
        NCrate = NCtemp
        End If
    
    NCtemp = NCtemp + 1
    SumExceedances = 0
    Wend
    
    If errnc = True Then
    NCrate = "ERROR"
    End If

End Function

'==============================================================================
' Name:     RwCurve
' Author:   PS
' Desc:     Returns the value of the RwCurve at frequency fStr
' Args:     CurveNo (index at 500Hz band), fStr(frequency band), Mode (can
'           optionally set to 'oct' mode)
' Comments: (1) Based on ISO717.1 - Acoustics Rating of Sound Insulation in
'           Buildings and of Building Elements  - Part 1: Airborne Sound
'           Insulation
'==============================================================================
Function RwCurve(CurveNo As Variant, fstr As String, Optional Mode As String)
Dim freq As Double
'''''''''''''''''''''''''''''''
'REFERENCE CURVES FROM ISO717.1
'band          125 250 500 1k  2k
Rw_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz, Rw52 curve
'band           100 125 160 200 250 315 400 500 630 800  1k 1.2k 1.6k 2k 2.5k 3.15k
Rw_ThOct = Array(33, 36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56) 'From 100 Hz to 3150 Hz, Rw52 curve
''''''''''''''''''''''''''''''''

freq = freqStr2Num(fstr)

    If freq < 100 Or freq > 3150 Then 'catch out of range errors
    RwCurve = "-"
    Exit Function
    End If
    
    IStart = 999 'for error checking
    
    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    'get array index, from 31.5Hz band
    IStart = GetArrayIndex_OCT(fstr, -1)
    Else 'one-third octave bands
    'get array index, from 100Hz band
    IStart = GetArrayIndex_TO(freq, -3)
    End If
    
    If IStart = 999 Then ' no matching band
        RwCurve = "-"
        Exit Function
    End If
    
    'build formula
    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    RwCurve = Rw_Oct(IStart) + CurveNo - 52
    Else
    RwCurve = Rw_ThOct(IStart) + CurveNo - 52
    End If

End Function


'==============================================================================
' Name:     RwRate
' Author:   PS
' Desc:     Rates the input data against the Rw curves and returns the single
'           value rating.
' Args:     DataTable (spectrum of TLs), Mode (can optionally set to 'oct' mode)
' Comments: (1) Based on ISO717.1 - Acoustics Rating of Sound Insulation in
'           Buildings and of Building Elements  - Part 1: Airborne Sound
'           Insulation
'           (2) Assumed to start at 100Hz (third oct) or 125Hz (oct) band
'           (3) Now references the function RwCurve to centralise
'           (4) To allow negative Rw spectra, there's now a multiplier of -1
'               when calculating deficiencies
'==============================================================================
Function RwRate(DataTable As Variant, Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Double 'empty array for deficiences
Dim Multiplier As Double
Dim MaxDeficiencies As Double
Dim fstr As String

'Legacy baseline curves for reference, I guess
'Rw10 curves
''band          100 125 160 200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
'Rw_ThOct = Array(-9, -6, -3, 0, 3, 6, 9, 10, 11, 12, 13, 14, 14, 14, 14, 14)
''band         125 250 500 1k  2k
'Rw_Oct = Array(-6, 3, 10, 13, 14)

SumDeficiencies = 0

    'check for negative values
    If DataTable(1) < 0 Then
    Multiplier = -1
    Else
    Multiplier = 1
    End If

    'set variables for modes and build array of frequencies
    If Mode = "oct" Then
    MaxDeficiencies = 10#
    bands = Array("125", "250", "500", "1k", "2k")
    Else
    MaxDeficiencies = 32#
    bands = Array("100", "125", "160", "200", "250", "315", "400", "500", _
        "630", "800", "1k", "1.25k", "1.6k", "2k", "2.5k", "3.15k")
    End If

'build curve from Rw5 (500Hz band)
CurveIndex = 5
ReDim Rw_curve(UBound(bands))
    For j = LBound(bands) To UBound(bands)
    fstr = bands(j)
    Rw_curve(j) = RwCurve(CurveIndex, fstr, Mode)
    Next j

    'Evaluate each Rw Curve
    While SumDeficiencies <= MaxDeficiencies
    
        'index Rw_curve +1dB
        For y = LBound(bands) To UBound(bands)
        Rw_curve(y) = Rw_curve(y) + 1
        Next y
        
        CurveIndex = CurveIndex + 1
    
    SumDeficiencies = 0 'reset at each evaluation
        
        'Loop over all values and add those below the curve to the sum of deficiencies
        For x = LBound(bands) To UBound(bands)
        CheckDef = Rw_curve(x) - DataTable(x + 1) * Multiplier ' VBA and it's stupid 1 indexing
            If CheckDef > 0 Then 'only positive values are deficient
            Deficiencies(x) = CheckDef
            Else
            Deficiencies(x) = 0
            End If
        SumDeficiencies = SumDeficiencies + Deficiencies(x)
        Next x
   'Debug.Print "Rw = " & CurveIndex
   'Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
    Wend

RwRate = CurveIndex - 1

'don't return anything less than Rw5
If RwRate <= 5 Then RwRate = "-"

End Function

'==============================================================================
' Name:     CtrRate
' Author:   PS
' Desc:     Rates the input data against the Ctr curves and returns the single
'           value rating.
' Args:     DataTable (spectrum of TLs), rw (Rating curve from RwRate,
'           Mode (can optionally set to 'oct' mode)
' Comments: (1) Based on ISO717.1 - Acoustics Rating of Sound Insulation in
'           Buildings and of Building Elements  - Part 1: Airborne Sound
'           Insulation
'==============================================================================
Function CtrRate(DataTable As Variant, rw As Integer, Optional Mode As String)
Dim i As Integer
Dim PartialSum As Double
Dim A As Double
'curves from ISO717.1
'band             100  125  160  200  250  315  400  500  630  800 1k  1.2k 1.6k 2k 2.5k 3.15k
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15) 'as per ISO717-1
'band                    50   63    80  100  125  160  200  250  315  400  500  630  800  1k  1.2k 1.6k 2k 2.5k 3.15k 4k  5k
Ctr_Oct_50to5000 = Array(-25, -23, -21, -20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15, -16, -18)
'band           125  250  500 1k  2k
Ctr_Oct = Array(-14, -10, -7, -4, -6) 'as per ISO717-1
'band                     63  125  250  500 1k  2k  4k
Ctr_Oct_50to5000 = Array(-18, -14, -10, -7, -4, -6 - 11) 'as per ISO717-1

'initialise
PartialSum = 0

'<-Todo: update for extended frequency range
'  50  63  80 ... 4k 5k
' -25 -23 -21 ...-16 -18

If IsNumeric(DataTable(0)) = False Then
Exit Function
End If

    'Octave Band mode
    If Mode = "oct" Or Mode = "Oct" Or Mode = "OCT" Then
        For i = LBound(Ctr_Oct) To UBound(Ctr_Oct)
        PartialSum = PartialSum + (10 ^ ((Ctr_Oct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    Else 'One third octave band mode
        For i = LBound(Ctr_ThOct) To UBound(Ctr_ThOct)
        PartialSum = PartialSum + (10 ^ ((Ctr_ThOct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    End If
    
A = Round(-10 * Application.WorksheetFunction.Log10(PartialSum), 0)
CtrRate = A - rw
End Function

'==============================================================================
' Name:     CRate
' Author:   PS
' Desc:     Rates the input data against the C curves and returns the single
'           value rating.
' Args:     DataTable (spectrum of TLs), rw (Rating curve from RwRate,
'           Mode (can optionally set to 'oct' mode)
' Comments: (1) Based on ISO717.1 - Acoustics Rating of Sound Insulation in
'           Buildings and of Building Elements  - Part 1: Airborne Sound
'           Insulation
'==============================================================================
Function CRate(DataTable As Variant, rw As Integer, Optional Mode As String)
Dim i As Integer
Dim PartialSum As Double
Dim A As Double
'curves from ISO717.1
'band             100  125  160  200  250  315  400  500  630  800 1k  1.2k 1.6k 2k 2.5k 3.15k
C_ThOct = Array(-29, -26, -23, -21, -19, -17, -15, -13, -12, -11, -10, -9, -9, -9, -9, -9) 'as per ISO717-1
'band                    50  63    80   100  125  160  200  250  315  400  500  630  800 1k  1.2k 1.6k 2k 2.5k 3.15k
C_ThOct_50to3150 = Array(-40, -36, -33, -29, -26, -23, -21, -19, -17, -15, -13, -12, -11, -10, -9, -9, -9, -9, -9)
'band                    50  63    80   100  125  160  200  250  315  400  500  630  800  1k  1.2k 1.6k  2k  2.5k 3.15k 4k   5k
C_ThOct_50to5000 = Array(-41, -37, -34, -30, -27, -24, -22, -20, -18, -16, -14, -13, -12, -11, -10, -10, -10, -10, -10, -10, -10)
'band           125  250  500 1k  2k
C_Oct = Array(-21, -14, -8, -5, -4) 'as per ISO717-1
'band                   63  125  250  500 1k  2k
C_Oct_50to3150 = Array(-31, -21, -14, -8, -5, -4)
'band                   63  125  250  500 1k  2k  4k
C_Oct_50to5000 = Array(-32, -22, -15, -9, -6, -5, -5)

'initialise
PartialSum = 0

'<-Todo: update method for extended frequency range
'NOTE: completely different spectrum!??!!??!

    'Octave Band mode
    If Mode = "oct" Or Mode = "Oct" Or Mode = "OCT" Then
        For i = LBound(C_Oct) To UBound(C_Oct)
        PartialSum = PartialSum + (10 ^ ((C_Oct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    Else 'One third octave band mode
        For i = LBound(C_ThOct) To UBound(C_ThOct)
        PartialSum = PartialSum + (10 ^ ((C_ThOct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    End If
    
A = Round(-10 * Application.WorksheetFunction.Log10(PartialSum), 0)
CRate = A - rw
End Function

'==============================================================================
' Name:     STCRate
' Author:   PS
' Desc:     Rates the input TL spectrum against the STC curves, returns the
'           rating
' Args:     DataTable (input TL spectrum)
' Comments: (1) Based on ASTM E413-16 Classification for Rating Sound Insulation
'==============================================================================
Function STCRate(DataTable As Variant)

Dim MaxDeficiency As Long
Dim SumDeficiencies As Long
Dim Deficiencies(16) As Long

'REFERENCE CURVES from Table 1 of ASTM E413-16
'band           125  160  200  250  315 400 500 600 1k 1.2k 1.6k 2k 2.5k 3.15k 4k
STC_ThOct = Array(-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4) 'STC0
'TODO: STC octave band mode
CurveIndex = STC_ThOct(6) '500 Hz band

    While SumDeficiencies <= 32 And MaxDeficiency <= 8

    'index STC curves
        For y = LBound(STC_ThOct) To UBound(STC_ThOct)
        STC_ThOct(y) = STC_ThOct(y) + 1
        Next y

    CurveIndex = CurveIndex + 1

    SumDeficiencies = 0 'reset at each evaluation
    MaxDeficiency = 0

        For x = LBound(STC_ThOct) To UBound(STC_ThOct)
        CheckDef = STC_ThOct(x) - Round(DataTable(x + 1), 0)
            If CheckDef > 0 Then 'only positive values are deficient
            Deficiencies(x) = CheckDef
            Else
            Deficiencies(x) = 0
            End If
        SumDeficiencies = SumDeficiencies + Deficiencies(x)
        Next x
    MaxDeficiency = Application.WorksheetFunction.Max(Deficiencies)
'        Debug.Print "STC = " & CurveIndex
'        Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
'        Debug.Print "Max Deficiency= " & MaxDeficiency
'        Debug.Print "                      "
    Wend

STCRate = CurveIndex - 1

End Function

'==============================================================================
' Name:     STCCurve
' Author:   PS
' Desc:     Returns the value of the STC Curve at frequency fStr
' Args:     CurveNo (index at 500Hz band), fStr(frequency band)
' Comments: (1) Based on ASTM E413-16 Classification for Rating Sound Insulation
'==============================================================================
Function STCCurve(CurveNo As Variant, fstr As String) 'Optional Mode As String)
Dim freq As Double
Dim IStart As Integer

'REFERENCE CURVES from Table 1 of ASTM E413-16
'STC_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz <---There is no octave version?
'band           125  160  200  250  315 400 500 600 1k 1.2k 1.6k 2k 2.5k 3.15k 4k
STC_ThOct = Array(-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4) 'STC0

freq = freqStr2Num(fstr)

    If freq < 125 Or freq > 4000 Then 'catch out of range errors
    STCCurve = "-"
    Exit Function
    End If

IStart = 999 'for error checking

    'get array index, from 125Hz band
    IStart = GetArrayIndex_TO(freq, -4)

    If IStart = 999 Then ' no matching band
        STCCurve = "-"
        Exit Function
    End If

    STCCurve = STC_ThOct(IStart) + CurveNo

End Function

'==============================================================================
' Name:     LnwCurve
' Author:   PS
' Desc:     Returns the value of the Lnw Curve at frequency fStr
' Args:     CurveNo (index at 500Hz band)
'           fStr(frequency band)
'           Mode - Set to "oct" for octave band
' Comments: (1) Based on ISO717-2 Acoustics � Rating of sound insulation
'           in buildings and of building elements - Part 2:  Impact sound
'           insulation
'==============================================================================
Function LnwCurve(CurveNo As Variant, fstr As String, Optional Mode As String)
Dim freq As Double

'REFERENCE CURVE Lnw60
'bands          125 250 500 1k  2k
Lnw_Oct = Array(67, 67, 65, 62, 49)
'bands          100  125 160  200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
Lnw_ThOct = Array(62, 62, 62, 62, 62, 62, 61, 60, 59, 58, 57, 54, 51, 48, 45, 42)

freq = freqStr2Num(fstr)

    If freq < 100 Or freq > 3150 Then 'catch out of range errors
    LnwCurve = "-"
    Exit Function
    End If

IStart = 999 'for error checking

    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    'get array index, from 125Hz band
    IStart = GetArrayIndex_OCT(fstr, -1)
    Else 'one-third octave bands
    'get array index, from 100Hz band
    IStart = GetArrayIndex_TO(freq, -3)
    End If
        
    If IStart = 999 Then ' no matching band
        LnwCurve = "-"
        Exit Function
    End If

    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    LnwCurve = Lnw_Oct(IStart) + CurveNo - 60
    Else
    LnwCurve = Lnw_ThOct(IStart) + CurveNo - 60
    End If

End Function

'==============================================================================
' Name:     LnwRate
' Author:   PS
' Desc:     Rates the input spectrum against the Lnw curves and returns the
'           rating
' Args:     DataTable(spectrum of values), Mode (optional string for
'           octave-band mode
' Comments: (1) Based on ISO717-2 Acoustics � Rating of sound insulation
'           in buildings and of building elements - Part 2:  Impact sound
'           insulation
'==============================================================================
Function LnwRate(DataTable As Variant, Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Double
Dim x, y As Integer

'REFERENCE CURVE Lnw88
'bands          125 250 500 1k  2k
Lnw_Oct = Array(90, 90, 88, 85, 72)
'bands          100  125 160  200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
Lnw_ThOct = Array(90, 90, 90, 90, 90, 90, 89, 88, 87, 86, 85, 82, 79, 76, 73, 70)

SumDeficiencies = 0
    
    If Mode = "oct" Then 'octave band mode
        While SumDeficiencies <= 10
            'move curve down by one
            For y = LBound(Lnw_Oct) To UBound(Lnw_Oct)
            Lnw_Oct(y) = Lnw_Oct(y) - 1
            Next y
            
            CurveIndex = CurveIndex - 1
        
        SumDeficiencies = 0 'reset at each evaluation
            
            For x = LBound(Lnw_Oct) To UBound(Lnw_Oct)
            CheckDef = DataTable(x + 1) - Lnw_ThOct(x) ' VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are deficient
                Deficiencies(x) = CheckDef
                Else
                Deficiencies(x) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(x)
            Next x
    '    Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
    '    Debug.Print "Rw = " & CurveIndex
        Wend
        
    Else 'one-third octaves mode
        
        While SumDeficiencies <= 32
        
            'index Lnw Curve
            For y = LBound(Lnw_ThOct) To UBound(Lnw_ThOct)
            Lnw_ThOct(y) = Lnw_ThOct(y) - 1
            Next y
            
        CurveIndex = Lnw_ThOct(7) '500 Hz band (zero index)
        'Debug.Print "Lnw: " & CurveIndex
        
        SumDeficiencies = 0 'reset at each evaluation
    
            For x = LBound(Lnw_ThOct) To UBound(Lnw_ThOct)
            CheckDef = DataTable(x + 1) - Lnw_ThOct(x) 'VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are 'deficient' i.e. too loud
                'Debug.Print CheckDef
                Deficiencies(x) = CheckDef
                Else
                Deficiencies(x) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(x)
            Next x
        'Debug.Print "Deficiencies: " & SumDeficiencies
        Wend
        
    End If
    
LnwRate = CurveIndex + 1

End Function

'==============================================================================
' Name:     CiRate
' Author:   PS
' Desc:     Rates the input spectrum against the Ci-correction
' Args:     DataTable(spectrum of values), Lnw (single number rating)
' Comments: (1) Based on ISO717-2 Acoustics � Rating of sound insulation
'           in buildings and of building elements - Part 2:  Impact sound
'           insulation - Annex A
'==============================================================================
Function CiRate(DataTable As Variant, Lnw As Integer, Optional Mode As String)
Dim LnSum As Double
Dim PartialSum As Single

    'check for correct number of values, as per Annex C
    'note that count and ubound are different
    If Mode = "oct" Then
        If DataTable.Count <> 5 Then
        CiRate = "Error"
        Exit Function
        End If
    Else
        If DataTable.Count <> 16 Then
        CiRate = "Error"
        Exit Function
        End If
    End If
    
    'check for numbers
    If IsNumeric(DataTable(1)) = True Then
    LnSum = SPLSUM(DataTable)
    'Debug.Print "LnSum:"; LnSum; "- 15 -"; Lnw
    CiRate = Round(LnSum, 0) - 15 - Lnw 'from A.2.1
    Else
    CiRate = "Error"
    End If
    
End Function

'==============================================================================
' Name:     IICRate
' Author:   PS
' Desc:     Determines the IIC rating of a spectrum, input from 100Hz to 3150Hz
'           one-third-octave bands
' Args:     DataTable - Cells of measured SPL from 100Hz one-third octave band
' Comments: (1) First cut
'==============================================================================
Function IICRate(DataTable As Variant)
Dim CurveIndex As Integer
Dim SumExceedances As Long
Dim Exceedances(16) As Long
Dim MaxExceedance As Long
Dim CheckDef As Long

'Reference curve, from ASTM E989-06, from 100Hz
'bands         100 125 160 200250315400500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
IIC_ThOct = Array(2, 2, 2, 2, 2, 2, 1, 0, -1, -2, -3, -6, -9, -12, -15, -18)

SumExceedances = 99999
MaxExceedance = 99999

    While SumExceedances > 32 Or MaxExceedance > 8
    
        'index Lnw Curve, one up from previous
        For y = LBound(IIC_ThOct) To UBound(IIC_ThOct)
        IIC_ThOct(y) = IIC_ThOct(y) + 1
        Next y
        
    CurveIndex = IIC_ThOct(7) '500 Hz band (zero index)
    
    SumExceedances = 0 'reset at each evaluation
    'MaxExceedance = 0
        For x = LBound(IIC_ThOct) To UBound(IIC_ThOct)
        CheckDef = Round(DataTable(x + 1), 0) - IIC_ThOct(x) 'round values
            If CheckDef > 0 Then 'only positive values are 'deficient' / too loud
            'Debug.Print CheckDef
            Exceedances(x) = CheckDef
            Else
            Exceedances(x) = 0
            End If
        SumExceedances = SumExceedances + Exceedances(x)
        MaxExceedance = Application.WorksheetFunction.Max(Exceedances)
        Next x
    'Debug.Print "Exceedances: " & SumExceedances; "Max: " & MaxExceedance
    Wend
IICRate = 110 - CurveIndex
End Function

'==============================================================================
' Name:     IICCurve
' Author:   PS
' Desc:     Returns the value of the IIC curve at frequency band fStr
' Args:     CurveNo, fStr
' Comments: (1) rough as guts!
'==============================================================================
Function IICCurve(CurveNo As Variant, fstr As String) 'Optional Mode As String)
Dim freq As Double

'Reference curve
'bands         100 125 160 200250315400500630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
IIC_ThOct = Array(2, 2, 2, 2, 2, 2, 1, 0, -1, -2, -3, -6, -9, -12, -15, -18) 'IIC0


freq = freqStr2Num(fstr)

    If freq < 100 Or freq > 3150 Then 'catch out of range errors
    IICCurve = "-"
    Exit Function
    End If
    
IStart = 999 'for error checking
    
'get array index, from 100Hz band
IStart = GetArrayIndex_TO(freq, -3)
        
    If IStart = 999 Then 'no matching band
        IIC_ThOct = "-"
        Exit Function
    End If
        
IICCurve = IIC_ThOct(IStart) + (110 - CurveNo)
End Function

'==============================================================================
' Name:     RNCcurve
' Author:   PS
' Desc:     Returns the value of the RNC curve at frequency band fStr
' Args:     CurveNo, fStr
' Comments: (1) rough as guts!
'==============================================================================
Function RNCcurve(CurveNo As Integer, fstr As String) '<------TODO check this function
Dim OctaveBandIndex As Integer
'Table 5 of ANSI S12.2
'coefficients for 16,31.5,63,125,250,500,1000,2000,4000,8000 bands
bands = Array(16, 31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000)
LevelRanges = Array(81, 76, 71, 66) 'first 4 bands
K1 = Array(64.3333, 51, 37.6667, 24.3333, 11, 6, 2, -2, -6, -10)
K1alt = Array(31, 26, 21, 16, 11, 6, 2, -2, -6, -10)
K2 = Array(3, 2, 1.5, 1.2, 1, 1, 1, 1, 1, 1)
K2alt = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)

freq = freqStr2Num(fstr)

    For i = LBound(bands) To UBound(bands)
        If freq = bands(i) Then
        OctaveBandIndex = i
        End If
    Next i
    
RNCcurve = (CurveNo / K2(OctaveBandIndex)) + K1(OctaveBandIndex)
    
    If OctaveBandIndex <= 3 Then 'first 4 bands
        If RNCcurve > LevelRanges(OctaveBandIndex) Then
        RNCcurve = (CurveNo / K2alt(OctaveBandIndex)) + K1alt(OctaveBandIndex)
        End If
    End If
    
End Function

'==============================================================================
' Name:     RNCrate
' Author:   PS
' Desc:     Determines theRNC rating of a spectrum, input from 16Hz to 8000Hz
'           octave bands
' Args:     DataTable - Cells of measured SPL
' Comments: (1) First cut
'==============================================================================
Function RNCrate(DataTable As Variant) '<------TODO check this function
Dim OctaveBandIndex As Integer
'Table 5 of ANSI S12.2
'coefficients for 16,31.5,63,125,250,500,1000,2000,4000,8000 bands
bands = Array(16, 31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000)
LevelRanges = Array(81, 76, 71, 66) 'first 4 bands
K1 = Array(64.3333, 51, 37.6667, 24.3333, 11, 6, 2, -2, -6, -10)
K1alt = Array(31, 26, 21, 16, 11, 6, 2, -2, -6, -10)
K2 = Array(3, 2, 1.5, 1.2, 1, 1, 1, 1, 1, 1)
K2alt = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)


'    For i = LBound(bands) To UBound(bands)
'        If f = bands(i) Then
'        OctaveBandIndex = i
'        End If
'    Next i
RNCrate = 0

    For i = LBound(bands) To UBound(bands)
    RNCrate_temp = (DataTable(i) - K1(i)) * K2(i)
        If OctaveBandIndex <= 3 Then 'first 4 bands
            If RNCrate > LevelRanges(OctaveBandIndex) Then
            RNCrate = (DataTable(i) - K1alt(i)) * K2alt(i)
            End If
        End If
        
        If RNCrate_temp > RNCrate Then
        RNCrate = RNCrate_temp
        End If
        
    Next i
    
RNCrate = Round(RNCrate, 0)
End Function


'==============================================================================
' Name:     RCcurve
' Author:   PS
' Desc:     Returns the value of the RC (MarkII) curve at frequency band fStr
' Args:     CurveNo, fStr
' Comments: (1) First cut
'==============================================================================
Function RCcurve(CurveStr As String, fstr As String)
Dim Delta As Integer
Dim i As Integer
Dim CurveNo As Integer

'Reference curve: RC50
'bands            16 31.5 63 125 250 500 *1k* 2k  4k  8k
BaseCurve = Array(75, 75, 70, 65, 60, 55, 50, 45, 40, 35) 'element 6 is 1k
If InStr(1, CurveStr, "(", vbTextCompare) > 0 Then
    CurveNo = CInt(Left(CurveStr, Len(CurveStr) - 4))
Else
    CurveNo = CInt(CurveStr)
End If

Delta = CurveNo - 50

'get array index, from 16Hz band
i = GetArrayIndex_OCT(fstr, 2)

If CurveNo < 30 And i <= 1 Then 'flatline below 31Hz
    RCcurve = 55
Else
    RCcurve = BaseCurve(i) + Delta
End If

End Function


'==============================================================================
' Name:     RCrate
' Author:   PS
' Desc:     Returns the value of the RC (MarkII) curve at frequency band fStr
' Args:     DataTable - table of noise values
'           FirstFreq - first frequency band of data, assumed to be 16Hz
' Comments: (1) Cut
'==============================================================================
Function RCrate(DataTable As Variant, Optional FirstFreq As String)
Dim arrDeviations(9) As Double '16 to 8k
'Dim arrBands(10) As String
Dim i As Integer
Dim PSIL As Integer
Dim CurrentOct As String
Dim inputRange As Range
'Dim rngLF, rngMF, rngHF As Range
Dim QAI As Double
'Dim arrBands() As String

'<- Todo: write this
arrBands = Array("16", "31.5", "63", "125", "250", "500", "1k", "2k", "4k", "8k")
'find array index of the 500Hz band

'average the 500Hz to 2kHz band to get the PSIL
PSIL = Round((DataTable(6) + DataTable(7) + DataTable(8)) / 3, 0)

'octBandIndex = GetArrayIndex_OCT(FirstFreq, 2) 'start at 16Hz=0
'Set rngLF = Range("A1").Resize(3, 1)
'Set rngMF = Range("A2").Resize(3, 1)
'Set rngHF = Range("A3").Resize(3, 1)

'calculate the deviations from the curve
For i = 1 To 9
CurrentOct = arrBands(i - 1)
    arrDeviations(i) = DataTable(i) - RCcurve(CStr(PSIL), CurrentOct)
Next i


LF = 10 * Application.WorksheetFunction.Log((10 ^ (arrDeviations(1) / 10)) + _
    (10 ^ (arrDeviations(2) / 10)) + (10 ^ (arrDeviations(3) / 10))) _
    - 10 * Application.WorksheetFunction.Log(3)
MF = 10 * Application.WorksheetFunction.Log((10 ^ (arrDeviations(4) / 10)) + _
    (10 ^ (arrDeviations(5) / 10)) + (10 ^ (arrDeviations(6) / 10))) _
    - 10 * Application.WorksheetFunction.Log(3)
HF = 10 * Application.WorksheetFunction.Log((10 ^ (arrDeviations(7) / 10)) + _
    (10 ^ (arrDeviations(8) / 10)) + (10 ^ (arrDeviations(9) / 10))) _
    - 10 * Application.WorksheetFunction.Log(3)
    
    
QAI = Application.WorksheetFunction.Max(LF, MF, HF) - Application.WorksheetFunction.Min(LF, MF, HF)

'Debug.Print LF; MF; HF
If QAI <= 5 Then
    RCrate = PSIL & "N"
ElseIf QAI > 5 Then

    If LF = Application.WorksheetFunction.Max(LF, MF, HF) Then
        RCrate = PSIL & "(LF)"
    ElseIf MF = Application.WorksheetFunction.Max(LF, MF, HF) Then
        RCrate = PSIL & "(MF)"
    ElseIf HF = Application.WorksheetFunction.Max(LF, MF, HF) Then
        RCrate = PSIL & "(HF)"
    Else
        RCrate = PSIL
    End If
    
Else
    RCrate = "-"
End If

End Function

'==============================================================================
' Name:     AlphaWRate
' Author:   PS
' Desc:     Calculates weighted absorption
' Args:     alphaTable - table of input values from 250Hz to 4kHz
' Comments: (1) As per ISO11654
'==============================================================================
Function AlphaWRate(inputTable As Variant)
Dim SumDeficiencies As Double
Dim CheckDef As Double
Dim RefCurve
Dim y As Integer
Dim alphaTable(5) As Double
'octave bands  250, 500, 1k, 2k, 4k
RefCurve = Array(0.8, 1#, 1#, 1#, 0.9) 'From ISO11654

'IMPORTANT: Arrays are indexed from zero, Variants are indexed from one

    'fix input values and put them into an array
    For y = 1 To 5
        If inputTable(y) > 1 Then 'max out alphaTable values at 1.0
        alphaTable(y - 1) = 1
        Else 'round to nearest 0.5
        alphaTable(y - 1) = Application.WorksheetFunction.MRound(inputTable(y), 0.05) 'zero
        End If
    Next y

SumDeficiencies = 0
    While SumDeficiencies <= 0.1
    'Debug.Print "alpha: " & RefCurve(1) '500Hz value
    SumDeficiencies = 0
        For i = LBound(RefCurve) To UBound(RefCurve)
        CheckDef = alphaTable(i) - RefCurve(i)
            If CheckDef > 0 Then
            SumDeficiencies = SumDeficiencies + CheckDef
            End If
        Next i
        
    'Debug.Print "Sum of deficiencies: " & SumDeficiencies
        'move up curve
        For i = LBound(RefCurve) To UBound(RefCurve)
        RefCurve(i) = RefCurve(i) - 0.05
        Next i
    
    Wend

'Curve number is the value at 500Hz
AlphaWRate = RefCurve(1) + 0.05
    
End Function


'==============================================================================
' Name:     AlphaWCurve
' Author:   PS
' Desc:     Returns value of Weighted alpha curve from index at 500Hz
' Args:
' Comments: (1) As per ISO11654
'==============================================================================
Function AlphaWCurve(CurveValue As Double, fstr As String)
Dim RefCurve
Dim freq As Integer
Dim CurveDifference As Double
Dim i As Integer

RefCurve = Array(0.8, 1#, 1#, 1#, 0.9) 'From ISO11654

freq = freqStr2Num(fstr)

    If freq < 250 Or freq > 4000 Then 'catch out of range errors
    AlphaWCurve = "-"
    Exit Function
    End If

i = GetArrayIndex_OCT(fstr, -2) 'from 250Hz onwards

CurveDifference = CurveValue - RefCurve(1) '500Hz value

AlphaWCurve = Application.WorksheetFunction.Max(RefCurve(i) + CurveDifference, 0)

End Function


'==============================================================================
' Name:     ISO_226
' Author:   PS & Jeff Tackett
' Desc:     Returns the equal loudness (phon) curves from ISO 226
' Args:     fStr - frequency band from 20Hz to 12.5kHz
'           phon - number of phon curve
' Comments: (1) From 10Hz to 20kHz
'           (2) From the internet:
'                   Tf   is the exponent for loudness perception,
'                   a_f  is the threshold of hearing, (a_f)
'                   Lu   is the magnitude of the linear transfer function normalized at 1kHz.

'==============================================================================
Function ISO_226(fstr As String, phon As Double)
Dim f As Double
Dim i As Integer
Dim a_f_i As Double
Dim Lu_i As Double
Dim Tf_i As Double
'
' Generates an Equal Loudness Contour as described in ISO 226
'
' Usage:  [SPL FREQ] = ISO226(PHON)
'
'         PHON is the phon value in dB SPL that you want the equal
'           loudness curve to represent. (1phon = 1dB @ 1kHz)
'         SPL is the Sound Pressure Level amplitude returned for
'           each of the 29 frequencies evaluated by ISO226.
'         FREQ is the returned vector of frequencies that ISO226
'           evaluates to generate the contour.
'
' Desc:   This function will return the equal loudness contour for
'         your desired phon level.  The frequencies evaulated in this
'         function only span from 20Hz - 12.5kHz, and only 29 selective
'         frequencies are covered.  This is the limitation of the ISO
'         standard.
'
'         In addition the valid phon range should be 0 - 90 dB SPL.
'         Values outside this range do not have experimental values
'         and their contours should be treated as inaccurate.
'
'         If more samples are required you should be able to easily
'         interpolate these values using spline().
'
' Author: Jeff Tackett 03/01/05

f = freqStr2Num(fstr)

'Error Trapping, level and frequency
If (phon < 0) Or (phon > 100) Or f < 20 Or f > 12500 Then
    'msg = MsgBox("Phon value out of bounds!")
    ISO_226 = "-"
    Exit Function
End If

'                /---------------------------------------\
'''''''''''''''''          TABLES FROM ISO226             '''''''''''''''''
'                \---------------------------------------/
'f = Array(20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, _
    1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500)

A_f = Array(0.532, 0.506, 0.48, 0.455, 0.432, 0.409, 0.387, 0.367, 0.349, 0.33, 0.315, _
    0.301, 0.288, 0.276, 0.267, 0.259, 0.253, 0.25, 0.246, 0.244, 0.243, 0.243, _
    0.243, 0.242, 0.242, 0.245, 0.254, 0.271, 0.301)

Lu = Array(-31.6, -27.2, -23#, -19.1, -15.9, -13#, -10.3, -8.1, -6.2, -4.5, -3.1, _
    -2#, -1.1, -0.4, 0#, 0.3, 0.5, 0#, -2.7, -4.1, -1#, 1.7, _
    2.5, 1.2, -2.1, -7.1, -11.2, -10.7, -3.1)

Tf = Array(78.5, 68.7, 59.5, 51.1, 44#, 37.5, 31.5, 26.5, 22.1, 17.9, 14.4, _
    11.4, 8.6, 6.2, 4.4, 3#, 2.2, 2.4, 3.5, 1.7, -1.3, -4.2, _
    -6#, -5.4, -1.5, 6#, 12.6, 13.9, 12.3)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'get index number based on frequency
i = GetArrayIndex_TO(f, 4)
'get values at index
a_f_i = A_f(i)
Lu_i = Lu(i)
Tf_i = Tf(i)

'Deriving sound pressure level from loudness level (iso226 sect 4.1)
Af = (0.00447 * (10 ^ (0.025 * phon) - 1.15)) + ((0.4 * (10 ^ (((Tf_i + Lu_i) / 10) - 9))) ^ a_f_i)

ISO_226 = (10 / a_f_i) * Application.WorksheetFunction.Log(Af) - Lu_i + 94

End Function


'==============================================================================
' Name:     ISO23591_RTtargets
' Author:   PS
' Desc:     Returns the constants for target range lines from the standard
' Args:     None
' Comments: (1)
'==============================================================================
Function ISO23591_RTtargets(RoomVolume As Double, RoomType As String, UseCase As String)
Dim A, B As Double
Dim i, ArrayIndex As Integer
'Dim MaxVolume(), MinVolume() As Double
'Dim A_Array(), B_Array() As Double

MinVolume = Array(35, 35, 50, 50, 35, 35, 200, 200, 1500, 1500, 600, 600)
MaxVolume = Array(1000, 1000, 6500, 6500, 3000, 3000, 6500, 6500, 6500, 6500, 6500, 6500)
A_Array = Array(0.325, 0.415, 0.415, 0.6, 0.55, 0.72, 0.345, 0.415, 0.525, 0.675, 0.6, 0.83)
B_Array = Array(0.335, 0.335, 0.335, 0.5, 0.45, 0.65, 0.335, 0.335, 0.425, 0.575, 0.5, 0.73)

UseArray = Array("Amplified music - lower limit", _
    "Amplified music - upper limit", _
    "Loud acoustic music - lower limit", _
    "Loud acoustic music - upper limit", _
    "Quiet acoustic music - lower limit", _
    "Quiet acoustic music - upper limit")

'match to element number
For i = LBound(UseArray) To UBound(UseArray)
    If UseArray(i) = UseCase Then ArrayIndex = i
Next i

'set array index based on matching names
If RoomType = "Rehearsal rooms" Then
    ArrayIndex = ArrayIndex + 0
ElseIf RoomType = "Rehearsal use of recital rooms" Then
    ArrayIndex = ArrayIndex + 6
Else
    ISO23591_RTtargets = "-"
    Exit Function
End If

If RoomVolume < MinVolume(ArrayIndex) Or RoomVolume > MaxVolume(ArrayIndex) Then
    ISO23591_RTtargets = "-"
    Exit Function
End If

'set constants
A = A_Array(ArrayIndex)
B = B_Array(ArrayIndex)

'return values
ISO23591_RTtargets = A * Application.WorksheetFunction.Log10(RoomVolume) - B

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     PutNR
' Author:   PS
' Desc:     Builds the NR rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'           (2) now does mech!
'==============================================================================
Sub PutNRInline()
PutNR True
End Sub

'==============================================================================
' Name:     PutNR
' Author:   PS
' Desc:     Builds the NR rating and curve formulas for the row above
' Args:     SkipReferenceCurve - for inline ratings
' Comments: (1) only enabled for octave band sheets - how to do mech?
'           (2) now does mech!
'==============================================================================
Sub PutNR(Optional SkipReferenceCurve As Boolean)
Dim RowToCheck As Integer
    If T_BandType <> "oct" Then ErrorOctOnly
    
'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If

'rate function
ParameterMerge Selection.Row
Range(T_ParamRng(0)) = "=NR_rate(" & Range(Cells(RowToCheck, T_LossGainStart), _
    Cells(RowToCheck, T_LossGainEnd)).Address(False, False) & "," & T_FreqStartRng & ")"
Range(T_ParamRng(0)).NumberFormat = """NR = ""0"

'Curve function
If SkipReferenceCurve = False And T_SheetType <> "MECH" Then 'show ref curve
    BuildFormula "NRcurve(" & T_ParamRng(0) & "," & T_FreqStartRng & ")"
    SetDescription "=CONCAT(""NR ""," & T_ParamRng(0) & ","" Curve"")", Selection.Row, True
    SetTraceStyle "Curve" 'style on 8ve band cols
    SetTraceStyle "Input", True 'style on parameter cols
End If

End Sub


'==============================================================================
' Name:     PutNCInline
' Author:   PS
' Desc:     Builds the NC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'           (2) now does mech!
'==============================================================================
Sub PutNCInline()
PutNC True
End Sub

'==============================================================================
' Name:     PutNC
' Author:   PS
' Desc:     Builds the NC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'==============================================================================
Sub PutNC(Optional SkipReferenceCurve As Boolean)
Dim RowToCheck As Integer

If T_BandType <> "oct" Then ErrorOctOnly

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If

'rating curve
ParameterMerge Selection.Row
Range(T_ParamRng(0)) = "=NCrate(" & Range(Cells(RowToCheck, T_LossGainStart), _
    Cells(RowToCheck, T_LossGainEnd)).Address(False, False) & "," & T_FreqStartRng & ")"
Range(T_ParamRng(0)).NumberFormat = """NC = ""0"

'Ref. Curve
If SkipReferenceCurve = False Then 'rate function and T_SheetType <> "MECH" Then
    BuildFormula "NCcurve(" & T_ParamRng(0) & "," & T_FreqStartRng & ")"
    SetDescription "=CONCAT(""NC ""," & T_ParamRng(0) & ","" Curve"")", Selection.Row, True
    SetTraceStyle "Curve" 'style on 8ve band cols
    SetTraceStyle "Input", True 'style on parameter cols
End If

End Sub

'==============================================================================
' Name:     PutPNC
' Author:   PS
' Desc:     Builds the PNC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'==============================================================================
Sub PutPNC()

BuildFormula "PNCcurve(" & Cells(Selection.Row, T_ParamStart).Address(False, True) & _
    "," & Cells(T_FreqRow, 5).Address(True, False) & ")"
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 40 '<default to 40
Cells(Selection.Row, T_ParamStart).NumberFormat = """PNC = ""0"

SetDescription "=CONCAT(""PNC ""," & T_ParamRng(0) & ","" Curve"")", Selection.Row, True

SetDataValidation T_ParamStart, "15,20,25,30,35,40,45,50,55,60,65,70"

SetTraceStyle "Curve" 'style on 8ve band cols
SetTraceStyle "Input", True 'style on parameter cols

End Sub

'==============================================================================
' Name:     PutRC
' Author:   PS
' Desc:     Builds the RC rating and curve formulas for the row above
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRC()

BuildFormula "RCcurve(" & Cells(Selection.Row, T_ParamStart).Address(False, True) & _
    "," & Cells(T_FreqRow, 5).Address(True, False) & ")"

'formatting and labels
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart).NumberFormat = """RC = ""0"
Cells(Selection.Row, T_ParamStart) = 50 '<default to 50 (ref curve)
SetDescription "=CONCAT(""RC ""," & T_ParamRng(0) & ","" Curve"")", Selection.Row, True

'TODO: rating function
'rate function
'Cells(Selection.Row, T_ParamStart) = "=RCrate(" & Range(Cells(Selection.Row - 1, T_LossGainStart), _
'    Cells(Selection.Row - 1, T_LossGainEnd)).Address(False, False) & "," & T_FreqStartRng & ")"

SetTraceStyle "Curve" 'style on 8ve band cols
SetTraceStyle "Input", True 'style on parameter cols

End Sub

'==============================================================================
' Name:     PutRw
' Author:   PS
' Desc:     Builds the Rw rating and curve formulas for the row above
' Args:     SkipReferenceCurve - set to true to skip the reference curve
' Comments: (1) Includes Ctr correction by default
'           (2) Updated to include description set from ParamCol
'           (3) updated to allow reference curve to be skipped
'           (4) only uses one BuildFormula
'==============================================================================
Sub PutRw(Optional SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer
Dim strBandMode As String 'only used when in octave band mode
Dim RowToCheck As Integer

SetDescription "=CONCAT(""Rw ""," & T_ParamRng(0) & ","" curve"")"

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If

    'set switches for octave band mode
If T_BandType = "oct" Then
    strBandMode = ",""oct"")" 'formula needs extra input
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
ElseIf T_BandType = "to" Then
    strBandMode = ")" 'nothing, no extra input
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
End If
    
    'check the frequency bands were found
    If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
    
If SkipReferenceCurve = False Then 'insert reference curve and style
    BuildFormula "RwCurve(" & T_ParamRng(0) & "," & T_FreqStartRng & strBandMode
    SetTraceStyle "Input", True
    SetTraceStyle "Curve" 'style on 8ve band cols
End If

'Rw Rate
Cells(Selection.Row, T_ParamStart).Value = "=RwRate(" & Range( _
    Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
    .Address(False, False) & strBandMode '125 hz to 2kHz
'Ctr Rate
Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & Range( _
    Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
    .Address(False, False) & "," & T_ParamRng(0) & strBandMode

'formatting
Cells(Selection.Row, T_ParamStart).NumberFormat = """Rw ""0"
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """Ctr"" 0;""Ctr -""0"

End Sub

'==============================================================================
' Name:     PutCtrInline
' Author:   PS
' Desc:     Rates the current row for the Ctr correction
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutCtrInline()
PutCtr True
End Sub

'==============================================================================
' Name:     PutCtr
' Author:   PS
' Desc:     Puts the Ctr correction into a row with an Rw rating in it
' Args:     SkipReferenceCurve - set to true to skip the reference curve
' Comments: (1) updated to allow reference curve to be skipped
'==============================================================================
Sub PutCtr(Optional SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''
'octave bands
''''''''''''''''''''''''''''''''''''''''''''''''''
If T_BandType = "oct" Then
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
    
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
        
    'Ctr Rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & Range( _
        Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
        .Address(False, False) & "," & T_ParamRng(0) & ",""oct"")"
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'one-third octave bands
    ''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf T_BandType = "to" Then
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
    
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
        
    'Ctr rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & Range( _
        Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
        .Address(False, False) & "," & T_ParamRng(0) & ")"
End If

'formatting
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """Ctr"" 0;""Ctr -""0"

SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutRwCtrRating
' Author:   PS
' Desc:     Rates current row with Rw and Ctr
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRwCtrInline()
PutRw True
PutCtr True
End Sub


'==============================================================================
' Name:     PutSTCInline
' Author:   PS
' Desc:     Rates current row with STC
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutSTCInline()
PutSTC True
End Sub

'==============================================================================
' Name:     PutC
' Author:   PS
' Desc:     Rates the current row for the C correction
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutCInline()
PutC True
End Sub

'==============================================================================
' Name:     PutC
' Author:   PS
' Desc:     Builds the 'C' correction to the right cells
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutC(SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''
'octave bands
''''''''''''''''''''''''''''''''''''''''''''''''''
If T_BandType = "oct" Then
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
    
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
    
    'Ctr Rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CRate(" & Range( _
        Cells(Selection.Row, StartBandCol), Cells(Selection.Row, EndBandCol)) _
        .Address(False, False) & "," & T_ParamRng(0) & ",""oct"")"
        
''''''''''''''''''''''''''''''''''''''''''''''''''
'one-third octave bands
''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf T_BandType = "to" Then
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
    
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
        
    'Ctr rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CRate(" & Range( _
        Cells(Selection.Row, StartBandCol), Cells(Selection.Row, EndBandCol)) _
        .Address(False, False) & "," & T_ParamRng(0) & ")"
End If

'formatting
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """C"" 0;""C -""0"

SetTraceStyle "Input", True

End Sub


'==============================================================================
' Name:     PutSTC
' Author:   PS
' Desc:     Builds the STC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for third-octave sheets
'           (2) Updated to include description set from ParamCol
'==============================================================================
Sub PutSTC(Optional SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer
Dim RowToCheck As Integer

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If

    If T_BandType = "oct" Then ErrorThirdOctOnly

StartBandCol = FindFrequencyBand("125")
EndBandCol = FindFrequencyBand("4k")

    If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing

ParameterMerge (Selection.Row)

'STC rate
Cells(Selection.Row, T_ParamStart).Value = "=STCRate(" & Range( _
    Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
    .Address(False, False) & ")" '125 hz to 4kHz
Cells(Selection.Row, T_ParamStart).NumberFormat = """STC""0"

If SkipReferenceCurve = False Then 'STC curve
    BuildFormula "STCCurve(" & T_ParamRng(0) & "," & T_FreqStartRng & ")"
    SetTraceStyle "Input", True
    SetTraceStyle "Curve", False
    SetDescription "=CONCAT(""STC ""," & T_ParamRng(0) & ","" curve"")"
End If

End Sub

'==============================================================================
' Name:     PutLnwInline
' Author:   PS
' Desc:     Rates current row with Lnw curve
' Args:     None
' Comments: (1) only enabled for third-octave sheets
'==============================================================================
Sub PutLnwInline()
PutLnw True
End Sub

'==============================================================================
' Name:     PutLnw
' Author:   PS
' Desc:     Builds the Lnw rating and curve formulas for the row above
' Args:     SkipReferenceCurve
' Comments: (1) only enabled for third-octave sheets
'==============================================================================
Sub PutLnw(Optional SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer
Dim RowToCheck As Integer

SetDescription "Lnw Curve"
ParameterMerge (Selection.Row)

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If
    
If T_BandType = "oct" Then 'octave band mode
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
    
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
        
    If SkipReferenceCurve = False Then 'Lnw Curve
        BuildFormula "LnwCurve(" & T_ParamRng(0) _
            & "," & T_FreqStartRng & ",""oct"")"
        SetTraceStyle "Input", True
        SetTraceStyle "Curve"
    End If
    
    'Lnw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=LnwRate(" & Range( _
        Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
        .Address(False, False) & ",""oct"")"
        
ElseIf T_BandType = "to" Then 'one third octave band mode
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
        
        If StartBandCol = 0 Or EndBandCol = 0 Then ErrorFrequencyBandMissing
        
    If SkipReferenceCurve = False Then 'Lnw Curve
        BuildFormula "LnwCurve(" & T_ParamRng(0) _
            & "," & T_FreqStartRng & ")"
        SetTraceStyle "Input", True
        SetTraceStyle "Curve"
    End If
    
    'Lnw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=LnwRate(" & Range( _
        Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
        .Address(False, False) & ")"
End If
    
'Formatting
Cells(Selection.Row, T_ParamStart).NumberFormat = """Lnw""0"
End Sub


'==============================================================================
' Name:     PutAWeight
' Author:   PS
' Desc:     Inserts the A-weighting curve
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutAweight()
Dim ApplyAsStatic As Long
Dim StartBandCol As Integer
Dim EndBandCol As Integer

ApplyAsStatic = MsgBox("Insert as static values? " & chr(10) & _
    "Note that 'No' will insert as formula", vbYesNoCancel, _
    "Formula / Static Values")
    
    'catch error
    If ApplyAsStatic = vbCancel Then End

'A weighting Curve
BuildFormula "AWeightCorrections(" & T_FreqStartRng & ")"
SetDescription "A Weighting Curve"

    If ApplyAsStatic = vbYes Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial Paste:=xlValues
    End If

SetTraceStyle "Curve"
End Sub

'==============================================================================
' Name:     PutCWeight
' Author:   PS
' Desc:     Inserts the C-weighting curve
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutCweight()
Dim ApplyAsStatic As Long
Dim StartBandCol As Integer
Dim EndBandCol As Integer

ApplyAsStatic = MsgBox("Insert as static values? " & chr(10) & _
    "Note that 'No' will insert as formula", vbYesNoCancel, _
    "Formula / Static Values")
    
    'catch error
    If ApplyAsStatic = vbCancel Then End

'A weighting Curve
BuildFormula "CWeightCorrections(" & T_FreqStartRng & ")"
SetDescription "C Weighting Curve"

    If ApplyAsStatic = vbYes Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial Paste:=xlValues
    End If

SetTraceStyle "Curve"
End Sub

'==============================================================================
' Name:     PutGWeight
' Author:   PS
' Desc:     Inserts the G-weighting curve (infrasound curve)
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutGweight()
Dim ApplyAsStatic As Long
Dim StartBandCol As Integer
Dim EndBandCol As Integer

ApplyAsStatic = MsgBox("Insert as static values? " & chr(10) & _
    "Note that 'No' will insert as formula", vbYesNoCancel, _
    "Formula / Static Values")
    
    'catch error
    If ApplyAsStatic = vbCancel Then End

'A weighting Curve
BuildFormula "GWeightCorrections(" & T_FreqStartRng & ")"
SetDescription "G Weighting Curve"
InsertComment "From ISO 7196-1995, defined at 0.25Hz to 100Hz", T_Description

    If ApplyAsStatic = vbYes Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial Paste:=xlValues
    End If

SetTraceStyle "Curve"
End Sub

'==============================================================================
' Name:     PutMassLaw
' Author:   PS
' Desc:     Mass law for transmission loss of walls
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutMassLaw()
ParameterMerge (Selection.Row)
BuildFormula "MassLaw(" & T_FreqStartRng & "," & T_ParamRng(0) & ")"
SetDescription "Mass Law"
Cells(Selection.Row, T_ParamStart) = 10.5 '<default to plasterboard
Cells(Selection.Row, T_ParamStart).NumberFormat = "0.0 ""kg/m" & chr(178) & """"
SetTraceStyle "Curve"
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutAlphaWInline
' Author:   PS
' Desc:     Rates current line for AlphaW
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutAlphaWInline()
PutAlphaW (True)
End Sub

'==============================================================================
' Name:     PutAlphaW
' Author:   PS
' Desc:     Inserts formula for weighted average absorption
' Args:     SkipReferenceCurve - doesn't rate the line
' Comments: (1)
'==============================================================================
Sub PutAlphaW(Optional SkipReferenceCurve As Boolean)
Dim StartBandCol As Integer
Dim EndBandCol As Integer
Dim RowToCheck As Integer

'set which row to do
If SkipReferenceCurve = True Then
    RowToCheck = Selection.Row
Else
    RowToCheck = Selection.Row - 1
End If

    If T_BandType <> "oct" Then ErrorOctOnly

StartBandCol = FindFrequencyBand("250")
EndBandCol = FindFrequencyBand("4k")

'rating
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart).Value = "=AlphaWRate(" & Range( _
    Cells(RowToCheck, StartBandCol), Cells(RowToCheck, EndBandCol)) _
    .Address(False, False) & ")"

'ref curve
If SkipReferenceCurve = False Then
    BuildFormula "AlphaWCurve(" & T_ParamRng(0) & "," & T_FreqStartRng & ")"
    SetTraceStyle "Curve"
    SetDescription "Weighted Alpha"
    SetTraceStyle "Input", True
End If

'formatting
Cells(Selection.Row, T_ParamStart).NumberFormat = """alpha_w=""0.0"
Range(Cells(Selection.Row, T_LossGainStart), Cells(Selection.Row, T_LossGainEnd)) _
    .NumberFormat = "0.0"
End Sub


'==============================================================================
' Name:     EqualLoudness
' Author:   PS
' Desc:     Inserts formula for equal loudness (phon) curve as per ISO-226
' Args:     None
' Comments: (1)
'==============================================================================
Sub EqualLoudness()
    If T_BandType <> "to" Then ErrorThirdOctOnly


BuildFormula "ISO_226(" & Cells(T_FreqRow, 5).Address(True, False) & "," & _
    T_ParamRng(0) & ")"
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 40 '<default to 40
Cells(Selection.Row, T_ParamStart).NumberFormat = "0" & """ phon"""

SetDescription "=concat(" & T_ParamRng(0) & ","" phon"")"

'parameter column styles
SetTraceStyle "Input", True

End Sub


'''''''''''''''''
'RC curve
'Eqn 4.45 of Biess and Hansen
'L_B=RC+ (5/0.3) * log(1000/f)


'Function SliceArray(inputArray() As Variant, startIndex As Integer, endIndex As Integer) As Variant()
'    Dim subsetArray() As Variant
'    Dim i As Integer
'
'    ' Resize and populate the subsetArray with the desired elements
'    ReDim subsetArray(startIndex To endIndex)
'    For i = startIndex To endIndex
'        subsetArray(i) = inputArray(i - 1) ' Adjust index since arrays are zero-based
'    Next i
'
'    ' Return the subsetArray
'    SliceArray = subsetArray
'End Function
