Attribute VB_Name = "Curves"
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

Function NCcurve(Curve_no As Integer, fstr As String)
Dim DataTable(0 To 11, 0 To 7) As Integer
Dim Ifreq As Integer
Dim freq As Integer
Dim DataRow As Integer

freq = freqStr2Num(fstr)

    If freq < 63 Then
    NCcurve = "-"
    Exit Function
    End If
'NC curves
NC15 = Array(47, 36, 29, 22, 17, 14, 12, 11)
NC20 = Array(51, 40, 33, 26, 22, 19, 17, 16)
NC25 = Array(55, 44, 37, 31, 27, 24, 22, 21)
NC30 = Array(57, 48, 41, 35, 31, 29, 28, 27)
NC35 = Array(60, 52, 45, 40, 36, 34, 33, 32)
NC40 = Array(64, 57, 50, 45, 41, 39, 38, 37)
NC45 = Array(67, 60, 54, 49, 46, 44, 43, 42)
NC50 = Array(71, 64, 58, 54, 51, 49, 48, 47)
NC55 = Array(74, 67, 62, 58, 56, 54, 53, 52)
NC60 = Array(77, 71, 67, 63, 61, 59, 58, 57)
NC65 = Array(80, 75, 71, 68, 66, 64, 63, 62)
NC70 = Array(83, 79, 75, 72, 71, 70, 69, 68)

    For i = 0 To 7
    DataTable(0, i) = NC15(i)
    DataTable(1, i) = NC20(i)
    DataTable(2, i) = NC25(i)
    DataTable(3, i) = NC30(i)
    DataTable(4, i) = NC35(i)
    DataTable(5, i) = NC40(i)
    DataTable(6, i) = NC45(i)
    DataTable(7, i) = NC50(i)
    DataTable(8, i) = NC55(i)
    DataTable(9, i) = NC60(i)
    DataTable(10, i) = NC65(i)
    DataTable(11, i) = NC65(i)
    Next i

    Select Case freq
    Case 63
        Ifreq = 0
    Case 125
        Ifreq = 1
    Case 250
        Ifreq = 2
    Case 500
        Ifreq = 3
    Case 1000
        Ifreq = 4
    Case 2000
        Ifreq = 5
    Case 4000
        Ifreq = 6
    Case 8000
        Ifreq = 7
    End Select

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
    Case Is = 70
    DataRow = 11
    End Select
    
NCcurve = DataTable(DataRow, Ifreq)

End Function

Function NR_rate(DataTable As Variant, Optional fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim NR_f, NR As Double
Dim NRTemp, temp_NR, freq As Double
Dim IStart, Col As Integer

    If DataTable.Rows.count <> 1 Then
        NRrate = "ERROR!"
        Exit Function
    End If
NRTemp = 0

'coefficients from Table 1 of AS1469
A_f = Array(55.4, 35.5, 22, 12, 4.8, 0, -3.5, -6.1, -8)
B_f = Array(0.681, 0.79, 0.87, 0.93, 0.974, 1, 1.015, 1.025, 1.03)
    If fstr = "" Then
    freq = 31.5
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
    For Col = 1 To DataTable.Columns.count
        If DataTable(1, Col) <> "-" Then
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
STC_ThOct = Array(36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56, 56) 'From 125 Hz to 4000 Hz, Rw52 curve
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
Dim PartialSum As Double
Dim i As Integer



'PartialSum = 0
'    For i = 0 To DataTable.Count
'    Debug.Print "No of elements="; DataTable.Count
'    PartialSum = PartialSum + (10 ^ ((DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
'    Next i
'LnSum = Round(10 * Application.WorksheetFunction.Log10(PartialSum), 0)
'Debug.Print "LnSum:"; LnSum; "- 15 -"; Lnw

'LnSum = SPLSUM(DataTable)
CiRate = Round(LnSum, 0) - 15 - Lnw
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
    Cells(Selection.Row, 14) = 40
    Cells(Selection.Row, 14).NumberFormat = """NC = ""0"
    paramcol = 14
    ElseIf Left(SheetType, 2) = "TO" Then
    'none
    End If
ExtendFunction (SheetType)
Call ParameterMerge(Selection.Row, SheetType)

    With Cells(Selection.Row, paramcol).Validation
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

Sub PutPNC(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 3) = "OCT" Then
    paramcol = 14
    Cells(Selection.Row, 5).Value = "=PNCcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, paramcol) = 40 '<default to 40
    Cells(Selection.Row, paramcol).NumberFormat = """PNC = ""0"
    Else
    ErrorOctOnly 'catch error
    End If
    
Cells(Selection.Row, 2).Value = "PNC Curve"

ExtendFunction (SheetType)

Call ParameterMerge(Selection.Row, SheetType)

    With Cells(Selection.Row, paramcol).Validation
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
