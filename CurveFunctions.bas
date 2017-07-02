Attribute VB_Name = "CurveFunctions"
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

Function NCcurve(Curve_no As Integer, fstr As String)
Dim Data(0 To 11, 0 To 7) As Integer
Dim Ifreq As Integer
Dim freq As Integer
freq = freqStr2Num(fstr)
    If freq < 63 Then
    NCcurve = "-"
    Exit Function
    End If
'NC curves
nc15 = Array(47, 36, 29, 22, 17, 14, 12, 11)
nc20 = Array(51, 40, 33, 26, 22, 19, 17, 16)
nc25 = Array(55, 44, 37, 31, 27, 24, 22, 21)
nc30 = Array(57, 48, 41, 35, 31, 29, 28, 27)
nc35 = Array(60, 52, 45, 40, 36, 34, 33, 32)
nc40 = Array(64, 57, 50, 45, 41, 39, 38, 37)
nc45 = Array(67, 60, 54, 49, 46, 44, 43, 42)
nc50 = Array(71, 64, 58, 54, 51, 49, 48, 47)
nc55 = Array(74, 67, 62, 58, 56, 54, 53, 52)
nc60 = Array(77, 71, 67, 63, 61, 59, 58, 57)
nc65 = Array(80, 75, 71, 68, 66, 64, 63, 62)
nc70 = Array(83, 79, 75, 72, 71, 70, 69, 68)

    For i = 0 To 7
    Data(0, i) = nc15(i)
    Data(1, i) = nc20(i)
    Data(2, i) = nc25(i)
    Data(3, i) = nc30(i)
    Data(4, i) = nc35(i)
    Data(5, i) = nc40(i)
    Data(6, i) = nc45(i)
    Data(7, i) = nc50(i)
    Data(8, i) = nc55(i)
    Data(9, i) = nc60(i)
    Data(10, i) = nc65(i)
    Data(11, i) = nc70(i)
    Next i
    
Ifreq = 0
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
  
    If (Curve_no > 70) Then
    NCcurve = "Max NC = 70"
    Exit Function
    End If
If (Curve_no < 72.5) Then NC_curve = 11
If (Curve_no < 67.5) Then NC_curve = 10
If (Curve_no < 62.5) Then NC_curve = 9
If (Curve_no < 57.5) Then NC_curve = 8
If (Curve_no < 52.5) Then NC_curve = 7
If (Curve_no < 47.5) Then NC_curve = 6
If (Curve_no < 42.5) Then NC_curve = 5
If (Curve_no < 37.5) Then NC_curve = 4
If (Curve_no < 32.5) Then NC_curve = 3
If (Curve_no < 27.5) Then NC_curve = 2
If (Curve_no < 22.5) Then NC_curve = 1
If (Curve_no < 17.5) Then NC_curve = 0
    If (Curve_no < 12.5) Then
    NCgen = "Min NC = 15"
    Exit Function
    End If
NCcurve = Data(NC_curve, Ifreq)
End Function

Function NR_rate(DataTable As Variant, Optional fstr As String)
Dim A_f As Variant
Dim B_f As Variant
Dim NR_f, NR As Double
Dim NRTemp, temp_NR, freq As Double
Dim IStart, Col As Integer

    If DataTable.Rows.Count <> 1 Then
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
    For Col = 1 To DataTable.Columns.Count
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

Function RwCurve(CurveNo As Variant, fstr As String) 'Optional Mode As String)

'If Mode <> "Oct" Or Mode <> "ThirdOct" Then
'    RwCurve = "ERROR!"
'    Exit Function
'End If

'''''''''''''''''''''''''''''''
'REFERENCE CURVES FROM ISO717.1
Rw_oct = Array(36, 45, 52, 55, 56) 'From 125 Hz to 2000 Hz
Rw_ThOct = Array(33, 36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56) 'From 100 Hz to 3150 Hz, Rw52 curve
Ctr_oct = Array(-14, -10, -7, -4, -6)
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15)
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
        RwCurve = "-"
        Exit Function
    End If
        
    
    RwCurve = Rw_ThOct(IStart) + CurveNo - 52


End Function

Function RwRate(DataTable As Variant)  'Optional Mode As String)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double

Rw_ThOct = Array(-9, -6, -3, 0, 3, 6, 9, 10, 11, 12, 13, 14, 14, 14, 14, 14) 'From 100 Hz to 3150 Hz, Rw10 curve
Deficiencies = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'empty array for deficiences
SumDeficiencies = 0
CurveIndex = Rw_ThOct(7) '500 Hz band

    While SumDeficiencies < 32
    
        'index Rw curve
        For y = 0 To 15
        Rw_ThOct(y) = Rw_ThOct(y) + 1
        Next y
        CurveIndex = CurveIndex + 1
    
    SumDeficiencies = 0 'reset at each evaluation
        For x = 0 To 15
        CheckDef = Rw_ThOct(x) - DataTable(x + 1) ' VBA and it's stupid 1 indexing
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

RwRate = CurveIndex - 1

End Function

Function CtrRate(DataTable As Variant, rw As Integer)
' Rw+ Ctr  for third octaves between 100 and 3150 Hz
Dim i As Integer
Dim PartialSum As Double
Ctr_ThOct = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15) 'From 100 Hz to 3150 Hz, as per ISO717-1
PartialSum = 0
    For i = 0 To 15
    PartialSum = PartialSum + (10 ^ ((Ctr_ThOct(i) - DataTable(i + 1)) / 10)) ' VBA and it's stupid 1 indexing
    Next i
A = Round(-10 * Application.WorksheetFunction.Log10(PartialSum), 0)
CtrRate = A - rw
End Function

Function LnwRate(DataTable As Variant)

Dim CurveIndex As Integer
Dim SumDeficiencies As Double

'Lnw for third octaves between 100 and 3150Hz
Lnw_ThOct = Array(90, 90, 90, 90, 90, 90, 89, 88, 87, 86, 85, 82, 79, 76, 73, 70) 'Lnw88 Reference curve, from ISO717-2
Deficiencies = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) 'empty array for deficiences
SumDeficiencies = 0

    While SumDeficiencies < CDbl(32)
    
        'index Lnw Curve
        For y = 0 To 15
        Lnw_ThOct(y) = Lnw_ThOct(y) - 1
        Next y
        
    CurveIndex = Lnw_ThOct(7) '500 Hz band (zero index)
    'Debug.Print "Lnw: " & CurveIndex
    
    SumDeficiencies = 0 'reset at each evaluation

        For x = 0 To 15
        CheckDef = DataTable(x + 1) - Lnw_ThOct(x) 'VBA and it's stupid 1 indexing
            If CheckDef > 0 Then 'only positive values are deficient
            'Debug.Print CheckDef
            Deficiencies(x) = CheckDef
            Else
            Deficiencies(x) = 0
            End If
        SumDeficiencies = SumDeficiencies + Deficiencies(x)
        Next x
    'Debug.Print SumDeficiencies
    Wend
LnwRate = CurveIndex + 1
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
End Sub

Sub PutNC(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "NC Curve"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=NCcurve($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
    Cells(Selection.Row, 14) = 40
    Cells(Selection.Row, 14).NumberFormat = """NC = ""0"
    ParamCol = 14
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=10*LOG($AA" & Selection.Row & "/(4*PI()*$Z" & Selection.Row & "^2))"
    End If
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

End Sub

Sub PutRw(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "Rw Curve"
Call ParameterUnmerge(Selection.Row, SheetType)
If Left(SheetType, 3) = "OCT" Then
'TODO Octave version
ElseIf Left(SheetType, 2) = "TO" Then
Cells(Selection.Row, 5).Value = "=RwCurve($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ")"
Cells(Selection.Row, 26).Value = "=RwRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & ")" '100 hz to 5kHz
Cells(Selection.Row, 26).NumberFormat = """Rw ""0"
Cells(Selection.Row, 27).Value = "=CtrRate(" & Range(Cells(Selection.Row - 1, 8), Cells(Selection.Row - 1, 23)).Address(False, False) & "," & Cells(Selection.Row, 26).Address(False, False) & ")" '100 hz to 5kHz
End If
ExtendFunction (SheetType)
End Sub

