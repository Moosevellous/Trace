Attribute VB_Name = "Curves"
Function AWeightCorrections(fstr As String)
Dim dBAAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fstr)

ArrayIndex = 999 'for error catching
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

Function CWeightCorrections(fstr As String)
Dim dBCAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

freq = freqStr2Num(fstr)

ArrayIndex = 999 'for error catching
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

Function NRcurve(Curve_no As Integer, fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim Ifreq As Integer
Dim freq As Double
freq = freqStr2Num(fstr)
Ifreq = 0
'coefficients from Table 1 of AS1469
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)
'''''''''''''''''''''''''''''''''
    Select Case freq
    Case 31.5
        Ifreq = 0
    Case 63
        Ifreq = 1
    Case 125
        Ifreq = 2
    Case 250
        Ifreq = 3
    Case 500
        Ifreq = 4
    Case 1000
        Ifreq = 5
    Case 2000
        Ifreq = 6
    Case 4000
        Ifreq = 7
    Case 8000
        Ifreq = 8
    End Select
NRcurve = A_f(Ifreq) + (B_f(Ifreq) * Curve_no)
End Function

Function PNCcurve(Curve_no As Integer, fstr As String)

Dim DataTable(0 To 10, 0 To 8) As Double

Dim IStart As Integer
Dim Ifreq As Integer

freq = freqStr2Num(fstr)

'define curves
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
    
    'select column of Data
    Select Case freq
    Case 31.5
        Ifreq = 0
    Case 63
        Ifreq = 1
    Case 125
        Ifreq = 2
    Case 250
        Ifreq = 3
    Case 500
        Ifreq = 4
    Case 1000
        Ifreq = 5
    Case 2000
        Ifreq = 6
    Case 4000
        Ifreq = 7
    Case 8000
        Ifreq = 8
    End Select
    
    'select row of Data
    Select Case Curve_no
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

Function NR_rate(DataTable As Variant, Optional fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim NR_f, NR As Double
Dim NRTemp, freq As Double
Dim IStart, Col As Integer

    If DataTable.Rows.Count <> 1 Then
        NR_rate = "ERROR!"
        Exit Function
    End If
NRTemp = 0

'coefficients from Table 1 of AS1469
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)
    If fstr = "" Then
    freq = 31.5 'if no frequency input, assume data starts at 31.5Hz
    Else
    freq = freqStr2Num(fstr)
    End If
    
    Select Case freq
        Case 31.5
            IStart = 0
        Case 63
            IStart = 1
        Case 125
            IStart = 2
        Case 250
            IStart = 3
        Case 500
            IStart = 4
        Case 1000
            IStart = 5
        Case 2000
            IStart = 6
        Case 4000
            IStart = 7
        Case 8000
            IStart = 8
    End Select
    
    'Debug.Print DataTable.Columns.Count
    For Col = 1 To DataTable.Columns.Count
        If IsNumeric(DataTable(1, Col)) Then
            NR_f = (DataTable(1, Col) - A_f(IStart + Col - 1)) / B_f(IStart + Col - 1) 'get the NR for that octave band
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

Function NCcurve(Curve_no As Integer, fstr As String)
Dim Ifreq As Integer
Dim freq As Integer

freq = freqStr2Num(fstr)

    If freq < 16 Then
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

    Select Case freq
    Case 16
        Ifreq = 0
    Case 31.5
        Ifreq = 1
    Case 63
        Ifreq = 2
    Case 125
        Ifreq = 3
    Case 250
        Ifreq = 4
    Case 500
        Ifreq = 5
    Case 1000
        Ifreq = 6
    Case 2000
        Ifreq = 7
    Case 4000
        Ifreq = 8
    Case 8000
        Ifreq = 9
    End Select
    
    If Curve_no Mod 5 = 0 Then 'simply return the defined value
        Select Case Curve_no
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
    NCcurve = InterpolateNCcurve(Curve_no, fstr)
    End If
    
End Function

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

Function NCrate(DataTable As Variant, Optional StartFreqStr As String)

Dim NC_f, NC As Double
Dim NCTemp, freq As Double
Dim IStart, Col As Integer
Dim i As Integer
   
   If DataTable.Rows.Count <> 1 Then
        NCrate = "ERROR!"
        Exit Function
    End If

octaveBands = Array(16, 31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000)

    If StartFreqStr = "" Then
    freq = 16 'if no frequency input, assume data starts at 16Hz octave band
    Else
    freq = freqStr2Num(StartFreqStr)
    End If

    Select Case freq
        Case 16
            IStart = 0
        Case 31.5
            IStart = 1
        Case 63
            IStart = 2
        Case 125
            IStart = 3
        Case 250
            IStart = 4
        Case 500
            IStart = 5
        Case 1000
            IStart = 6
        Case 2000
            IStart = 7
        Case 4000
            IStart = 8
        Case 8000
            IStart = 9
    End Select

    i = 15
    found = False
    SumExceedances = 0
    While found = False
    'Debug.Print "Checking NC"; i
    test_freq = octaveBands(IStart)
        For Col = 1 To DataTable.Columns.Count 'all input value
        test_freq = octaveBands(IStart + Col - 1) 'DataTable is indexed from 1, not 0
            If IsNumeric(DataTable(1, Col)) Then
            'get value of curve at that band
            NC_curve_value = NCcurve(i, CStr(test_freq))
            'Debug.Print DataTable(1, Col + 1).Value; "    NCvalue: "; NC_curve_value
                If DataTable(1, Col).Value > NC_curve_value Then
                SumExceedances = SumExceedances + (DataTable(1, Col) - NC_curve_value)
                End If
            End If
        Next Col
    
        'catch error
        If i > 70 Then
        found = True
        errnc = True
        ElseIf SumExceedances = 0 Then
        found = True
        NCrate = i
        End If
    
    i = i + 1
    SumExceedances = 0
    Wend
    
If errnc = True Then
NCrate = "ERROR"
End If

End Function

Function RwCurve(CurveNo As Variant, fstr As String, Optional Mode As String)

'If Mode <> "Oct" Or Mode <> "ThirdOct" Then
'    RwCurve = "ERROR!"
'    Exit Function
'End If

'''''''''''''''''''''''''''''''
'REFERENCE CURVES FROM ISO717.1
Rw_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz, Rw52 curve
Rw_ThOct = Array(33, 36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56) 'From 100 Hz to 3150 Hz, Rw52 curve
Ctr_Oct = Array(-14, -10, -7, -4, -6)
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15)
''''''''''''''''''''''''''''''''

    If fstr = "" Then
    freq = 31.5
    Else
    freq = freqStr2Num(fstr)
    End If
    
    IStart = 999 'for error checking
    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
        Select Case freq
            Case 125
                IStart = 0
            Case 250
                IStart = 1
            Case 500
                IStart = 2
            Case 1000
                IStart = 3
            Case 2000
                IStart = 4
        End Select
    Else
        Select Case freq
            Case 100
                IStart = 0
            Case 125
                IStart = 1
            Case 160
                IStart = 2
            Case 200
                IStart = 3
            Case 250
                IStart = 4
            Case 315
                IStart = 5
            Case 400
                IStart = 6
            Case 500
                IStart = 7
            Case 630
                IStart = 8
            Case 800
                IStart = 9
            Case 1000
                IStart = 10
            Case 1250
                IStart = 11
            Case 1600
                IStart = 12
            Case 2000
                IStart = 13
            Case 2500
                IStart = 14
            Case 3150
                IStart = 15
        End Select
    End If
    
    If IStart = 999 Then ' no matching band
        RwCurve = "-"
        Exit Function
    End If
        
    If Mode = "oct" Or Mode = "OCT" Or Mode = "Oct" Then
    RwCurve = Rw_Oct(IStart) + CurveNo - 52
    Else
    RwCurve = Rw_ThOct(IStart) + CurveNo - 52
    End If

End Function

Function RwRate(DataTable As Variant, Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Long 'empty array for deficiences

'TODO - make this reference RwCurve
Rw_ThOct = Array(-9, -6, -3, 0, 3, 6, 9, 10, 11, 12, 13, 14, 14, 14, 14, 14) 'From 100 Hz to 3150 Hz, Rw10 curve
Rw_Oct = Array(-6, 3, 10, 13, 14) 'From 125 Hz to 2kHz octave bands, Rw10 curve

SumDeficiencies = 0

CurveIndex = Rw_ThOct(7) '500 Hz band

    If Mode = "oct" Then
        While SumDeficiencies < 10
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
    Else
        While SumDeficiencies < 32
        
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
    '    Debug.Print "SUM DEFICIENCIES= " & SumDeficiencies
    '    Debug.Print "Rw = " & CurveIndex
        Wend
    End If 'end of Mode switch

RwRate = CurveIndex - 1

End Function


Function CtrRate(DataTable As Variant, rw As Integer, Optional Mode As String)
' Rw + Ctr  for third octaves between 100 and 3150 Hz
Dim i As Integer
Dim PartialSum As Double
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15) 'From 100 Hz to 3150 Hz, as per ISO717-1
Ctr_Oct = Array(-14, -10, -7, -4, -6) 'From 100 Hz to 3150 Hz, as per ISO717-1
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
a = Round(-10 * Application.WorksheetFunction.Log10(PartialSum), 0)
CtrRate = a - rw
End Function


Function STCRate(DataTable As Variant, Optional Mode As String)

Dim MaxDeficiency As Long
Dim SumDeficiencies As Long
Dim Deficiencies(16) As Long

STC_ThOct = Array(-6, -3, 0, 3, 6, 9, 10, 11, 12, 13, 14, 14, 14, 14, 14, 14) 'STC10 from 125Hz to 4kHz
CurveIndex = STC_ThOct(6) '500 Hz band

    While SumDeficiencies < 32 And MaxDeficiency < 8

    'index STC curves
        For Y = LBound(STC_ThOct) To UBound(STC_ThOct)
        STC_ThOct(Y) = STC_ThOct(Y) + 1
        Next Y

    CurveIndex = CurveIndex + 1

        SumDeficiencies = 0 'reset at each evaluation
        MaxDeficiency = 0

            For X = LBound(STC_ThOct) To UBound(STC_ThOct)
            CheckDef = STC_ThOct(X) - DataTable(X + 1) ' VBA and it's stupid 1 indexing
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

Function STCCurve(CurveNo As Variant, fstr As String) 'Optional Mode As String)

'If Mode <> "Oct" Or Mode <> "ThirdOct" Then
'    RwCurve = "ERROR!"
'    Exit Function
'End If

'''''''''''''''''''''''''''''''
'REFERENCE CURVES
'STC_Oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz
STC_ThOct = Array(36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56, 56) 'From 125 Hz to 4000 Hz, STC52 curve
''''''''''''''''''''''''''''''''

    If fstr = "" Then
    freq = 31.5 'why?
    Else
    freq = freqStr2Num(fstr)
    End If

    IStart = 999 'for error checking

    Select Case freq
        Case 125
            IStart = 0
        Case 160
            IStart = 1
        Case 200
            IStart = 2
        Case 250
            IStart = 3
        Case 315
            IStart = 4
        Case 400
            IStart = 5
        Case 500
            IStart = 6
        Case 630
            IStart = 7
        Case 800
            IStart = 8
        Case 1000
            IStart = 9
        Case 1250
            IStart = 10
        Case 1600
            IStart = 11
        Case 2000
            IStart = 12
        Case 2500
            IStart = 13
        Case 3150
            IStart = 14
        Case 4000
            IStart = 15
    End Select

    If IStart = 999 Then ' no matching band
        STCCurve = "-"
        Exit Function
    End If

    STCCurve = STC_ThOct(IStart) + CurveNo - 52

End Function


Function LnwCurve(CurveNo As Variant, fstr As String) 'Optional Mode As String)

'If Mode <> "Oct" Or Mode <> "ThirdOct" Then
'    RwCurve = "ERROR!"
'    Exit Function
'End If

'''''''''''''''''''''''''''''''
'REFERENCE CURVES FROM ISO717.2
'Lnw_Oct = Array(67, 67, 65, 62, 49)
Lnw_ThOct = Array(62, 62, 62, 62, 62, 62, 61, 60, 59, 58, 57, 54, 51, 48, 45, 42) 'From 100 Hz to 3150 Hz, Lnw60 curve
'Ci_oct = Array(-14, -10, -7, -4, -6)
'Ci_ThOct = Array(-20, -20, -18, -16, -15,-14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15)
''''''''''''''''''''''''''''''''

    If fstr = "" Then
    freq = 31.5
    Else
    freq = freqStr2Num(fstr)
    End If
    
IStart = 999 'for error checking
    
    Select Case freq
        Case 100
            IStart = 0
        Case 125
            IStart = 1
        Case 160
            IStart = 2
        Case 200
            IStart = 3
        Case 250
            IStart = 4
        Case 315
            IStart = 5
        Case 400
            IStart = 6
        Case 500
            IStart = 7
        Case 630
            IStart = 8
        Case 800
            IStart = 9
        Case 1000
            IStart = 10
        Case 1250
            IStart = 11
        Case 1600
            IStart = 12
        Case 2000
            IStart = 13
        Case 2500
            IStart = 14
        Case 3150
            IStart = 15
    End Select
        
    If IStart = 999 Then ' no matching band
        LnwCurve = "-"
        Exit Function
    End If
        
LnwCurve = Lnw_ThOct(IStart) + CurveNo - 60

End Function


Function LnwRate(DataTable As Variant)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double
Dim Deficiencies(16) As Long

'Lnw for third octaves between 100 and 3150Hz
Lnw_ThOct = Array(90, 90, 90, 90, 90, 90, 89, 88, 87, 86, 85, 82, 79, 76, 73, 70) 'Lnw88 Reference curve, from ISO717-2
Lnw_Oct = Array(90, 90, 88, 85, 72)
SumDeficiencies = 0

    While SumDeficiencies < 32
    
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
LnwRate = CurveIndex + 1
End Function

Function CiRate(DataTable As Variant, Lnw As Integer)
Dim LnSum As Double
Dim PartialSum As Single
Dim i As Integer

    'Debug.Print "No of elements="; DataTable.Count
    If DataTable.Count = 15 And IsNumeric(DataTable(1)) = True Then 'check for 15 values, 100Hz to 2500Hz, as per ISO 717.2
    LnSum = SPLSUM(DataTable)
    'Debug.Print "LnSum:"; LnSum; "- 15 -"; Lnw
    CiRate = Round(LnSum, 0) - 15 - Lnw 'from A.2.1 of ISO 717.2
    Else 'too many columns
    CiRate = "Error"
    End If
    
End Function

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

f = freqStr2Num(fstr)

    For i = LBound(bands) To UBound(bands)
        If f = bands(i) Then
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

'''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''

Sub PutNR(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "NR Curve"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=NRcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 14) = "=NR_rate(" & Range(Cells(Selection.Row - 1, 5), Cells(Selection.Row - 1, 13)).Address(False, False) & ")"
    Cells(Selection.Row, 14).NumberFormat = """NR = ""0"
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=10*LOG($AA" & Selection.Row & "/(4*PI()*$Z" & Selection.Row & "^2))"
    End If
ExtendFunction (SheetType)
Call ParameterMerge(Selection.Row, SheetType)

fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub

Sub PutNC(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "NC Curve"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=NCcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 14) = "=NCrate(" & Range(Cells(Selection.Row - 1, 5), Cells(Selection.Row - 1, 13)).Address(False, False) & ",$E$6)"
    Cells(Selection.Row, 14).NumberFormat = """NC = ""0"
    ParamCol = 14
    ElseIf Left(SheetType, 2) = "TO" Then
    'none
    End If
ExtendFunction (SheetType)
Call ParameterMerge(Selection.Row, SheetType)

fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub

Sub PutPNC(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 3) = "OCT" Then
    ParamCol = 14
    Cells(Selection.Row, 5).Value = "=PNCcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, ParamCol) = 40 '<default to 40
    Cells(Selection.Row, ParamCol).NumberFormat = """PNC = ""0"
    Else
    ErrorOctOnly 'catch error
    End If
    
Cells(Selection.Row, 2).Value = "PNC Curve"

ExtendFunction (SheetType)

Call ParameterMerge(Selection.Row, SheetType)

    With Cells(Selection.Row, ParamCol).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="15,20,25,30,35,40,45,50,55,60,65,70"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub

Sub PutRw(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 3) = "OCT" Then
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 5).Value = "=RwCurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",""oct"")"
    Cells(Selection.Row, 14).Value = "=RwRate(" & Range(Cells(Selection.Row - 1, 7), Cells(Selection.Row - 1, 11)).Address(False, False) & ",""oct"")" '125 hz to 2kHz
    Cells(Selection.Row, 14).NumberFormat = """Rw ""0"
    'Cells(Selection.Row, 27).Value = "=CtrRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & "," & Cells(Selection.Row, 26).Address(False, False) & ")" '125 hz to 2kHz
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 2).Value = "Rw Curve"
    Call ParameterUnmerge(Selection.Row, SheetType)
    Cells(Selection.Row, 5).Value = "=RwCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 26).Value = "=RwRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & ")" '100 hz to 5kHz
    Cells(Selection.Row, 26).NumberFormat = """Rw ""0"
    Cells(Selection.Row, 27).Value = "=CtrRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & "," & Cells(Selection.Row, 26).Address(False, False) & ")" '100 hz to 5kHz
    Cells(Selection.Row, 27).NumberFormat = """Ctr"" 0;""Ctr -""0"
    End If
    ExtendFunction (SheetType)
    
fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub

Sub PutSTC(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 3) = "OCT" Then
    'Call ParameterMerge(Selection.Row, SheetType)
    'Cells(Selection.Row, 5).Value = "=RwCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    'Cells(Selection.Row, 14).Value = "=STCRate(" & Range(Cells(Selection.Row, 7), Cells(Selection.Row, 11)).Address(False, False) & ",""oct"")" '125 hz to 2kHz
    'Cells(Selection.Row, 14).NumberFormat = """STC""0"
    'Cells(Selection.Row, 27).Value = "=CtrRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & "," & Cells(Selection.Row, 26).Address(False, False) & ")" '125 hz to 2kHz
    'ExtendFunction (SheetType)
    ErrorThirdOctOnly
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 2).Value = "STC Curve"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 5).Value = "=STCCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 26).Value = "=STCRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & ")" '100 hz to 3.15kHz
    Cells(Selection.Row, 26).NumberFormat = """STC""0"
    ExtendFunction (SheetType)
    End If

fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub



Sub PutLnw(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 3) = "OCT" Then
    'Call ParameterMerge(Selection.Row, SheetType)
    'Cells(Selection.Row, 5).Value = "=RwCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    'Cells(Selection.Row, 14).Value = "=STCRate(" & Range(Cells(Selection.Row, 7), Cells(Selection.Row, 11)).Address(False, False) & ",""oct"")" '125 hz to 2kHz
    'Cells(Selection.Row, 14).NumberFormat = """STC""0"
    'Cells(Selection.Row, 27).Value = "=CtrRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & "," & Cells(Selection.Row, 26).Address(False, False) & ")" '125 hz to 2kHz
    'ExtendFunction (SheetType)
    ErrorThirdOctOnly
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 2).Value = "Lnw Curve"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 5).Value = "=LnwCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 26).Value = "=LnwRate(" & Range(Cells(Selection.Row - 1, 9), Cells(Selection.Row - 1, 24)).Address(False, False) & ")"
    Cells(Selection.Row, 26).NumberFormat = """Lnw""0"
    ExtendFunction (SheetType)
    End If

fmtUserInput SheetType, True '<-Format Parameter column as user input

End Sub

'''''''''''''''''
'RC curve
'Eqn 4.45 of Biess and Hansen
'L_B=RC+ (5/0.3) * log(1000/f)
