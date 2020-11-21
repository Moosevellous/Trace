Attribute VB_Name = "Curves"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     AWeightCorrections
' Author:   PS
' Desc:     Returns the A-weighting corrections in one-thrd octave bands
' Args:     fstr (frequency band)
' Comments: (1) From 10Hz to 20kHz
'==============================================================================
Function AWeightCorrections(fStr As String)
Dim dBAAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fStr)

ArrayIndex = 999 'for error catching

'                     Corrections for 1/3 octave band centre frequencies
'                     10     12.5   16     20     25     31.5   40     50     63     80     100    125    160    200    250   315   400   500   630   800 1000 1250 1600 2000 2500 3150 4000 5000 6300 8000  10000 12500 16000 20000)
dBAAdjustment = Array(-70.4, -63.4, -56.7, -50.5, -44.7, -39.4, -34.6, -30.2, -26.2, -22.5, -19.1, -16.1, -13.4, -10.9, -8.6, -6.6, -4.8, -3.2, -1.9, -0.8, 0#, 0.6, 1#, 1.2, 1.3, 1.2, 1#, 0.5, -0.1, -1.1, -2.5, -4.3, -6.6, -9.3)
freqTitles = Array(10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000)
'freqTitlesAlt = Array("", 13, 32, "1k", "1.25k", "1.6k", "2k", "2.5k", "3.15k", "4k", "5k", "6k", "6.3k", "8k", "10k", "12.5k", "16k", "20k")
    
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
' Args:     fstr (frequency band)
' Comments: (1) From 10Hz to 20kHz
'==============================================================================
Function CWeightCorrections(fStr As String)
Dim dBCAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fStr)

ArrayIndex = 999 'for error catching

'                     Corrections for 1/3 octave band centre frequencies
'                     10     12.5   16     20     25     31.5   40     50     63     80     100    125    160    200    250   315   400   500   630   800 1000 1250 1600 2000 2500 3150 4000 5000 6300 8000  10000 12500 16000 20000)
dBCAdjustment = Array(-14.3, -11.2, -8.5, -6.2, -4.4, -3.1, -2#, -1.3, -0.8, -0.5, -0.3, -0.2, -0.1, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.1, -0.2, -0.3, -0.5, -0.8, -1.3, -2#, -3#, -4.4, -6.2, -8.5, -11.2)
freqTitles = Array(10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000, 12500, 16000, 20000)
'freqTitlesAlt = Array("", 13, 32, "1k", "1.25k", "1.6k", "2k", "2.5k", "3.15k", "4k", "5k", "6k", "6.3k", "8k", "10k", "12.5k", "16k", "20k")

'alternative methods!!!!
'http://www.beis.de/Elektronik/AudioMeasure/WeightingFilters.html
'=10*LOG(((35041384000000000*f^8)/((20.598997^2+f^2)^2*(107.65265^2+f^2)*(737.86223^2+f^2)*(12194.217^2+f^2)^2)))

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
' Name:     NRcurve
' Author:   PS
' Desc:     Returns the values of the NR curve, at CurveNo
' Args:     CurveNo (name/number of the curve) fstr (frequency band)
' Comments: (1) Based on AS1469 - Acoustics - Methods for the determination of
'           noise rating numbers
'==============================================================================
Function NRcurve(CurveNo As Integer, fStr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim Ifreq As Integer
Dim freq As Double
freq = freqStr2Num(fStr)

    If freq < 31 Or freq > 8000 Then
    NRcurve = "-"
    Exit Function
    End If

'coefficients from Table 1 of AS1469
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)
'''''''''''''''''''''''''''''''''
Ifreq = GetArrayIndex_OCT(fStr, 1)

NRcurve = A_f(Ifreq) + (B_f(Ifreq) * CurveNo)
End Function

'==============================================================================
' Name:     PNCcurve
' Author:   PS
' Desc:     Returns the value of the PNC curve at that frequency band
' Args:     fstr, the frequency band centre frequency as a string, CurveNo the
'           index number of the curve
' Comments: (1) Based on someone else's code I think?
'==============================================================================
Function PNCcurve(CurveNo As Integer, fStr As String)

Dim DataTable(0 To 10, 0 To 8) As Double
Dim IStart As Integer
Dim Ifreq As Integer
Dim freq As Double

freq = freqStr2Num(fStr) 'for checking only

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
Ifreq = GetArrayIndex_OCT(fStr, 1)
    
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
'           fstr, the starting frequency
' Comments: (1) Based on AS1469 - Acoustics - Methods for the determination of
'           noise rating numbers
'==============================================================================
Function NR_rate(DataTable As Variant, Optional fStr As String)
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
    If fStr = "" Then fStr = "31.5"
    
'get array index, from 31.5Hz band
IStart = GetArrayIndex_OCT(fStr, 1)
    
    'Debug.Print DataTable.Columns.Count
    For Col = 1 To DataTable.Columns.Count
        If IsNumeric(DataTable(1, Col)) Then
            NR_f = (DataTable(1, Col) - A_f(IStart + Col - 1)) / _
                B_f(IStart + Col - 1) 'get the NR for that octave band
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
' Args:     CurveNo, the index number of the NC curve, fstr, the frequency band
'           as a string
' Comments: (1) Based on ANSI S12.2 2008
'==============================================================================
Function NCcurve(CurveNo As Integer, fStr As String)
Dim Ifreq As Integer
Dim freq As Integer

freq = freqStr2Num(fStr)

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
Ifreq = GetArrayIndex_OCT(fStr, 2)
    
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
    NCcurve = InterpolateNCcurve(CurveNo, fStr)
    End If
    
End Function

'==============================================================================
' Name:     InterpolateNCcurve
' Author:   PS
' Desc:     Returns the value of the interpolated NC curve at input freuqency
' Args:     CurveNo, the index number of the NC curve,
'           fstr, the frequency band as a string
' Comments: (1) Based on ANSI S12.2 2008
'==============================================================================
Function InterpolateNCcurve(CurveNo As Integer, fStr As String)

Dim freq As Integer
Dim Remainder As Integer
Dim UpperCurveValue As Single
Dim LowerCurveValue As Single
Dim UpperCurve As Integer
Dim LowerCurve As Integer

freq = freqStr2Num(fStr)

Remainder = CurveNo Mod 5
'x values
UpperCurve = CurveNo + (5 - Remainder)
LowerCurve = CurveNo - Remainder
'y values
UpperCurveValue = NCcurve(UpperCurve, fStr)
LowerCurveValue = NCcurve(LowerCurve, fStr)

'interpolate linearly
m = (UpperCurveValue - LowerCurveValue) / (UpperCurve - LowerCurve)
InterpolateNCcurve = LowerCurveValue + (m * (CurveNo - LowerCurve))

End Function

'==============================================================================
' Name:     NCrate
' Author:   PS
' Desc:     Rates the input spectrum against the NC curve
' Args:     DataTable - the spectrum to be rated, assumed to start at 16Hz band
'           Optional fstr, the starting frequency band
' Comments: (1) Calls NCcurve for values
'==============================================================================
Function NCrate(DataTable As Variant, Optional fStr As String)

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

    If fStr = "" Then
    freq = 16 'if no frequency input, assume data starts at 16Hz octave band
    Else
    freq = freqStr2Num(fStr)
    End If

'get array index, from 16Hz band
IStart = GetArrayIndex_OCT(fStr, 2)

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
' Desc:     Returns the value of the RwCurve at frequency fstr
' Args:     CurveNo (index at 500Hz band), fstr(frequency band), Mode (can
'           optionally set to 'oct' mode)
' Comments: (1) Based on ISO717.1 - Acoustics Rating of Sound Insulation in
'           Buildings and of Building Elements  - Part 1: Airborne Sound
'           Insulation
'==============================================================================
Function RwCurve(CurveNo As Variant, fStr As String, Optional Mode As String)
Dim freq As Double
'''''''''''''''''''''''''''''''
'REFERENCE CURVES FROM ISO717.1
'band          125 250 500 1k  2k
Rw_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz, Rw52 curve
'band           100 125 160 200 250 315 400 500 630 800  1k 1.2k 1.6k 2k 2.5k 3.15k
Rw_ThOct = Array(33, 36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56) 'From 100 Hz to 3150 Hz, Rw52 curve
''''''''''''''''''''''''''''''''

freq = freqStr2Num(fStr)

    If freq < 100 Or freq > 3150 Then 'catch out of range errors
    RwCurve = "-"
    Exit Function
    End If
    
    IStart = 999 'for error checking
    
    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    'get array index, from 31.5Hz band
    IStart = GetArrayIndex_OCT(fStr, -1)
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
'==============================================================================
Function RwRate(DataTable As Variant, Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Double 'empty array for deficiences

'<-------------------------TODO - make this reference RwCurve
'Rw10 curves
'band          100 125 160 200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
Rw_ThOct = Array(-9, -6, -3, 0, 3, 6, 9, 10, 11, 12, 13, 14, 14, 14, 14, 14)
'band         125 250 500 1k  2k
Rw_Oct = Array(-6, 3, 10, 13, 14)

SumDeficiencies = 0

CurveIndex = Rw_ThOct(7) '500 Hz band

    If Mode = "oct" Then
        While SumDeficiencies <= 10
            For Y = LBound(Rw_Oct) To UBound(Rw_Oct)
            Rw_Oct(Y) = Rw_Oct(Y) + 1
            Next Y
            
            CurveIndex = CurveIndex + 1
        
        SumDeficiencies = 0 'reset at each evaluation
            
            For X = LBound(Rw_Oct) To UBound(Rw_Oct)
            CheckDef = Rw_Oct(X) - DataTable(X + 1) ' VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are deficient
                Deficiencies(X) = CheckDef
                Else
                Deficiencies(X) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(X)
            Next X
    '    Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
    '    Debug.Print "Rw = " & CurveIndex
        Wend
    Else 'one-third octave mode
        While SumDeficiencies <= 32#
        
            'index Rw curves
            For Y = LBound(Rw_ThOct) To UBound(Rw_ThOct)
            Rw_ThOct(Y) = Rw_ThOct(Y) + 1
            Next Y
            
            CurveIndex = CurveIndex + 1
        
        SumDeficiencies = 0 'reset at each evaluation
    
            For X = LBound(Rw_ThOct) To UBound(Rw_ThOct)
            CheckDef = Rw_ThOct(X) - DataTable(X + 1) ' VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are deficient
                Deficiencies(X) = CheckDef
                Else
                Deficiencies(X) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(X)
            Next X
'        Debug.Print "Rw = " & CurveIndex
'        Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
        Wend
    End If 'end of Mode switch

RwRate = CurveIndex - 1

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
Function CtrRate(DataTable As Variant, Rw As Integer, Optional Mode As String)
Dim i As Integer
Dim PartialSum As Double
Dim A As Double
'curves from ISO717.1
'band             100  125  160  200  250  315  400  500  630  800 1k  1.2k 1.6k 2k 2.5k 3.15k
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15) 'as per ISO717-1
'band           125  250  500 1k  2k
Ctr_Oct = Array(-14, -10, -7, -4, -6) 'as per ISO717-1
PartialSum = 0

    'Octave Band mode
    If Mode = "oct" Or Mode = "Oct" Or Mode = "OCT" Then
        For i = LBound(Ctr_Oct) To UBound(Ctr_Oct)
        PartialSum = PartialSum + (10 ^ ((Ctr_Oct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    Else 'One third octave band mode
        For i = 0 To 15
        PartialSum = PartialSum + (10 ^ ((Ctr_ThOct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
        Next i
    End If
    
A = Round(-10 * Application.WorksheetFunction.Log10(PartialSum), 0)
CtrRate = A - Rw
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
CurveIndex = STC_ThOct(6) '500 Hz band

    While SumDeficiencies <= 32 And MaxDeficiency <= 8

    'index STC curves
        For Y = LBound(STC_ThOct) To UBound(STC_ThOct)
        STC_ThOct(Y) = STC_ThOct(Y) + 1
        Next Y

    CurveIndex = CurveIndex + 1

    SumDeficiencies = 0 'reset at each evaluation
    MaxDeficiency = 0

        For X = LBound(STC_ThOct) To UBound(STC_ThOct)
        CheckDef = STC_ThOct(X) - Round(DataTable(X + 1), 0)
            If CheckDef > 0 Then 'only positive values are deficient
            Deficiencies(X) = CheckDef
            Else
            Deficiencies(X) = 0
            End If
        SumDeficiencies = SumDeficiencies + Deficiencies(X)
        Next X
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
' Desc:     Returns the value of the STC Curve at frequency fstr
' Args:     CurveNo (index at 500Hz band), fstr(frequency band)
' Comments: (1) Based on ASTM E413-16 Classification for Rating Sound Insulation
'==============================================================================
Function STCCurve(CurveNo As Variant, fStr As String) 'Optional Mode As String)
Dim freq As Double
Dim IStart As Integer

'REFERENCE CURVES from Table 1 of ASTM E413-16
'STC_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz <---There is no octave version?
'band           125  160  200  250  315 400 500 600 1k 1.2k 1.6k 2k 2.5k 3.15k 4k
STC_ThOct = Array(-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4) 'STC0

freq = freqStr2Num(fStr)

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
' Desc:     Returns the value of the Lnw Curve at frequency fstr
' Args:     CurveNo (index at 500Hz band)
'           fstr(frequency band)
'           Mode - Set to "oct" for octave band
' Comments: (1) Based on ISO717-2 Acoustics — Rating of sound insulation
'           in buildings and of building elements - Part 2:  Impact sound
'           insulation
'==============================================================================
Function LnwCurve(CurveNo As Variant, fStr As String, Optional Mode As String)
Dim freq As Double

'REFERENCE CURVE Lnw60
'bands          125 250 500 1k  2k
Lnw_Oct = Array(67, 67, 65, 62, 49)
'bands          100  125 160  200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
Lnw_ThOct = Array(62, 62, 62, 62, 62, 62, 61, 60, 59, 58, 57, 54, 51, 48, 45, 42)

freq = freqStr2Num(fStr)

    If freq < 100 Or freq > 3150 Then 'catch out of range errors
    LnwCurve = "-"
    Exit Function
    End If

IStart = 999 'for error checking

    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    'get array index, from 125Hz band
    IStart = GetArrayIndex_OCT(fStr, -1)
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
' Comments: (1) Based on ISO717-2 Acoustics — Rating of sound insulation
'           in buildings and of building elements - Part 2:  Impact sound
'           insulation
'==============================================================================
Function LnwRate(DataTable As Variant, Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Double
Dim X, Y As Integer

'REFERENCE CURVE Lnw88
'bands          125 250 500 1k  2k
Lnw_Oct = Array(90, 90, 88, 85, 72)
'bands          100  125 160  200 250 315 400 500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
Lnw_ThOct = Array(90, 90, 90, 90, 90, 90, 89, 88, 87, 86, 85, 82, 79, 76, 73, 70)

SumDeficiencies = 0
    
    If Mode = "oct" Then
        While SumDeficiencies <= 10
            'move curve down by one
            For Y = LBound(Lnw_Oct) To UBound(Lnw_Oct)
            Lnw_Oct(Y) = Lnw_Oct(Y) - 1
            Next Y
            
            CurveIndex = CurveIndex - 1
        
        SumDeficiencies = 0 'reset at each evaluation
            
            For X = LBound(Lnw_Oct) To UBound(Lnw_Oct)
            CheckDef = DataTable(X + 1) - Lnw_ThOct(X) ' VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are deficient
                Deficiencies(X) = CheckDef
                Else
                Deficiencies(X) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(X)
            Next X
    '    Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
    '    Debug.Print "Rw = " & CurveIndex
        Wend
        
    Else 'one-third octave mode
        
        While SumDeficiencies <= 32
        
            'index Lnw Curve
            For Y = LBound(Lnw_ThOct) To UBound(Lnw_ThOct)
            Lnw_ThOct(Y) = Lnw_ThOct(Y) - 1
            Next Y
            
        CurveIndex = Lnw_ThOct(7) '500 Hz band (zero index)
        'Debug.Print "Lnw: " & CurveIndex
        
        SumDeficiencies = 0 'reset at each evaluation
    
            For X = LBound(Lnw_ThOct) To UBound(Lnw_ThOct)
            CheckDef = DataTable(X + 1) - Lnw_ThOct(X) 'VBA and it's stupid 1 indexing
                If CheckDef > 0 Then 'only positive values are 'deficient' i.e. too loud
                'Debug.Print CheckDef
                Deficiencies(X) = CheckDef
                Else
                Deficiencies(X) = 0
                End If
            SumDeficiencies = SumDeficiencies + Deficiencies(X)
            Next X
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
' Comments: (1) Based on ISO717-2 Acoustics — Rating of sound insulation
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

'Reference curve, from ASTM E9890-6, from 100Hz
'bands         100 125 160 200250315400500 630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
IIC_ThOct = Array(2, 2, 2, 2, 2, 2, 1, 0, -1, -2, -3, -6, -9, -12, -15, -18)

SumExceedances = 99999
MaxExceedance = 99999

    While SumExceedances > 32 Or MaxExceedance > 8
    
        'index Lnw Curve, one up from previous
        For Y = LBound(IIC_ThOct) To UBound(IIC_ThOct)
        IIC_ThOct(Y) = IIC_ThOct(Y) + 1
        Next Y
        
    CurveIndex = IIC_ThOct(7) '500 Hz band (zero index)
    
    SumExceedances = 0 'reset at each evaluation
    'MaxExceedance = 0
        For X = LBound(IIC_ThOct) To UBound(IIC_ThOct)
        CheckDef = Round(DataTable(X + 1), 0) - IIC_ThOct(X) 'round values
            If CheckDef > 0 Then 'only positive values are 'deficient' / too loud
            'Debug.Print CheckDef
            Exceedances(X) = CheckDef
            Else
            Exceedances(X) = 0
            End If
        SumExceedances = SumExceedances + Exceedances(X)
        MaxExceedance = Application.WorksheetFunction.Max(Exceedances)
        Next X
    'Debug.Print "Exceedances: " & SumExceedances; "Max: " & MaxExceedance
    Wend
IICRate = 110 - CurveIndex
End Function

'==============================================================================
' Name:     IICCurve
' Author:   PS
' Desc:     Returns the value of the IIC curve at frequency band fstr
' Args:     CurveNo, fstr
' Comments: (1) rough as guts!
'==============================================================================
Function IICCurve(CurveNo As Variant, fStr As String) 'Optional Mode As String)
Dim freq As Double

'Reference curve
'bands         100 125 160 200250315400500630 800 1k 1.2k 1.6k 2k 2.5k 3.15k
IIC_ThOct = Array(2, 2, 2, 2, 2, 2, 1, 0, -1, -2, -3, -6, -9, -12, -15, -18) 'IIC0


freq = freqStr2Num(fStr)

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
' Desc:     Returns the value of the RNC curve at frequency band fstr
' Args:     CurveNo, fstr
' Comments: (1) rough as guts!
'==============================================================================
Function RNCcurve(CurveNo As Integer, fStr As String) '<------TODO check this function
Dim OctaveBandIndex As Integer
'Table 5 of ANSI S12.2
'coefficients for 16,31.5,63,125,250,500,1000,2000,4000,8000 bands
bands = Array(16, 31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000)
LevelRanges = Array(81, 76, 71, 66) 'first 4 bands
K1 = Array(64.3333, 51, 37.6667, 24.3333, 11, 6, 2, -2, -6, -10)
K1alt = Array(31, 26, 21, 16, 11, 6, 2, -2, -6, -10)
K2 = Array(3, 2, 1.5, 1.2, 1, 1, 1, 1, 1, 1)
K2alt = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1)

freq = freqStr2Num(fStr)

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
Sub PutNR()
Cells(Selection.Row, T_Description).Value = "NR Curve"

    If T_BandType <> "oct" Then ErrorOctOnly

'rate function
Range(T_ParamRng(0)) = "=NR_rate(" & Range(Cells(Selection.Row - 1, T_LossGainStart), _
    Cells(Selection.Row - 1, T_LossGainEnd)).Address(False, False) & "," & T_FreqStartRng & ")"
    
    'Curve function
    If T_SheetType <> "MECH" Then
    Cells(Selection.Row, T_LossGainStart).Value = "=NRcurve(" & T_ParamRng(0) & _
        "," & T_FreqStartRng & ")"
    ExtendFunction
    End If
    
'formatting
Range(T_ParamRng(0)).NumberFormat = """NR = ""0"
Call ParameterMerge(Selection.Row)
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutNC
' Author:   PS
' Desc:     Builds the NC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'==============================================================================
Sub PutNC()

Cells(Selection.Row, T_Description).Value = "NC Curve"

If T_BandType <> "oct" Then ErrorOctOnly
'rate function
Range(T_ParamRng(0)) = "=NCrate(" & Range(Cells(Selection.Row - 1, T_LossGainStart), _
    Cells(Selection.Row - 1, T_LossGainEnd)).Address(False, False) & "," & T_FreqStartRng & ")"
    'Curve function
    If T_SheetType <> "MECH" Then
    Cells(Selection.Row, T_LossGainStart).Value = "=NCcurve(" & T_ParamRng(0) & _
        "," & T_FreqStartRng & ")"
    ExtendFunction
    End If
'formatting
Range(T_ParamRng(0)).NumberFormat = """NC = ""0"
Call ParameterMerge(Selection.Row)
SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutPNC
' Author:   PS
' Desc:     Builds the PNC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for octave band sheets - how to do mech?
'==============================================================================
Sub PutPNC()

Cells(Selection.Row, T_LossGainStart).Value = "=PNCcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 40 '<default to 40
Cells(Selection.Row, T_ParamStart).NumberFormat = """PNC = ""0"

Cells(Selection.Row, T_Description).Value = "PNC Curve"

ExtendFunction

SetDataValidation T_ParamStart, "15,20,25,30,35,40,45,50,55,60,65,70"

'parameter column styles
SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutRw
' Author:   PS
' Desc:     Builds the Rw rating and curve formulas for the row above
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRw()
Dim StartBandCol As Integer
Dim EndBandCol As Integer

Cells(Selection.Row, T_Description).Value = "Rw Curve"
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If T_BandType = "oct" Then
    'Rw Curve
    Cells(Selection.Row, T_LossGainStart).Value = "=RwCurve(" & T_ParamRng(0) _
        & "," & T_FreqStartRng & ",""oct"")"
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
    'Rw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=RwRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & ",""oct"")" '125 hz to 2kHz
    'Ctr Rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & "," & T_ParamRng(0) & ",""oct"")"
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf T_BandType = "to" Then
    'Rw Curve
    Cells(Selection.Row, T_LossGainStart).Value = "=RwCurve(" & T_ParamRng(0) _
        & "," & T_FreqStartRng & ")"
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
    'Rw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=RwRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & ")"    '125 hz to 2kHz
    'Ctr rate
    Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & "," & _
    T_ParamRng(0) & ")"
    End If

'formatting
Cells(Selection.Row, T_ParamStart).NumberFormat = """Rw ""0"
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """Ctr"" 0;""Ctr -""0"

ExtendFunction

SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutSTC
' Author:   PS
' Desc:     Builds the STC rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for third-octave sheets
'==============================================================================
Sub PutSTC()
Dim StartBandCol As Integer
Dim EndBandCol As Integer

    If T_BandType = "oct" Then
    ErrorThirdOctOnly
    ElseIf T_BandType = "to" Then
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("4k")
    Cells(Selection.Row, T_Description).Value = "STC Curve"
    ParameterMerge (Selection.Row)
    'STC curve
    Cells(Selection.Row, T_LossGainStart).Value = "=STCCurve(" & T_ParamRng(0) _
        & "," & T_FreqStartRng & ")"
    ExtendFunction
    'STC rate
    Cells(Selection.Row, T_ParamStart).Value = "=STCRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & ")" '125 hz to 4kHz
    'Formatting
    Cells(Selection.Row, T_ParamStart).NumberFormat = """STC""0"
    SetTraceStyle "Input", True
    End If

End Sub


'==============================================================================
' Name:     PutLnw
' Author:   PS
' Desc:     Builds the Lnw rating and curve formulas for the row above
' Args:     None
' Comments: (1) only enabled for third-octave sheets
'==============================================================================
Sub PutLnw()
Dim StartBandCol As Integer
Dim EndBandCol As Integer

Cells(Selection.Row, T_Description).Value = "Lnw Curve"
ParameterMerge (Selection.Row)
    
    If T_BandType = "oct" Then 'octave band mode
    StartBandCol = FindFrequencyBand("125")
    EndBandCol = FindFrequencyBand("2k")
    'Lnw Curve
    Cells(Selection.Row, T_LossGainStart).Value = "=LnwCurve(" & T_ParamRng(0) _
        & "," & T_FreqStartRng & ",""oct"")"
    ExtendFunction
    'Lnw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=LnwRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & ",""oct"")"
        
    ElseIf T_BandType = "to" Then 'one third octave band mode
    StartBandCol = FindFrequencyBand("100")
    EndBandCol = FindFrequencyBand("3.15k")
    'Lnw Curve
    Cells(Selection.Row, T_LossGainStart).Value = "=LnwCurve(" & T_ParamRng(0) _
        & "," & T_FreqStartRng & ")"
    ExtendFunction
    'Lnw Rate
    Cells(Selection.Row, T_ParamStart).Value = "=LnwRate(" & Range( _
        Cells(Selection.Row - 1, StartBandCol), Cells(Selection.Row - 1, EndBandCol)) _
        .Address(False, False) & ")"
    End If
    
'Formatting
Cells(Selection.Row, T_ParamStart).NumberFormat = """Lnw""0"
SetTraceStyle "Input", True
End Sub


'==============================================================================
' Name:     PutAWeight
' Author:   PS
' Desc:     Inserts the A weighting curve
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
Cells(Selection.Row, T_LossGainStart).Value = "=AWeightCorrections(" & _
    T_FreqStartRng & ")"
ExtendFunction
Cells(Selection.Row, T_Description).Value = "A Weighting Curve"

    If ApplyAsStatic = vbYes Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial Paste:=xlValues
    End If

SetTraceStyle "Curve"
End Sub

'==============================================================================
' Name:     PutAWeight
' Author:   PS
' Desc:     Inserts the A weighting curve
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
Cells(Selection.Row, T_LossGainStart).Value = "=CWeightCorrections(" & _
    T_FreqStartRng & ")"
ExtendFunction
Cells(Selection.Row, T_Description).Value = "C Weighting Curve"

    If ApplyAsStatic = vbYes Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial Paste:=xlValues
    End If

SetTraceStyle "Curve"
End Sub

'''''''''''''''''
'RC curve
'Eqn 4.45 of Biess and Hansen
'L_B=RC+ (5/0.3) * log(1000/f)
