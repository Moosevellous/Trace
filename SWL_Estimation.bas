Attribute VB_Name = "SWL_Estimation"
Public DescriptionString As String

Public FanType As String
Public FanV As Long
Public FanP As Long

Public PumpEqn As String
Public PumpPower As Long

Public CTEqn As String
Public CTPower As Long
Public CT_Type As String
Public CT_Correction(0 To 8) As Long
Public CT_Direction(0 To 9) As Variant
Public CT_Dir_checked As Boolean

Public MotorType As String
Public MotorEqn As String
Public MotorPower As Long
Public MotorSpeed As Long
Public Motor_Correction(0 To 8) As Long

Public TurbinePower As Long
Public TurbineEqn As String
Public TurbineCorrection(0 To 9) As Long
Public TurbineEnclosure(0 To 9) As Long
Public GasTurbineType As String
Public EnclosureDescription As String

Public CompressorSPL(0 To 8) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function LwFanSimple(freq As String, V As Double, P As Double, FanType As String)

LwOverall = 10 * Application.WorksheetFunction.Log10(V) + 20 * Application.WorksheetFunction.Log10(P) + 40 'v in m^3, p in Pa
 
    Select Case FanType
    Case Is = ""
    LwFanSimple = LwSimple
    Case Is = "Forward curved centrifugal"
    Correction = Array(-5, -10, -15, -20, -25, -28, -31) 'SRL
    Case Is = "Backward curved centrifugal"
    Correction = Array(-10, -11, -10, -15, -20, -25, -30) 'SRL
    Case Is = "Radial or paddle blade"
    Correction = Array(3, -3, -10, -11, -15, -19, -23) 'SRL
    Case Is = "Axial"
    Correction = Array(-8, -8, -6, -7, -8, -12, -16) 'MDA/Woods
    Case Is = "Bifurcated"
    Correction = Array(-3, -3, -4, -5, -7, -8, -11) 'SRL
    Case Is = "Propeller fan(approx)"
    Correction = Array(-3, -4, -1, -8, -12, -13, -20) 'SRL
    'Variable Inlet Vanes
    Case Is = "Variable inlet vanes - 100%"
    Correction = Array(0, 0, 0, 0, 0, 0, 0) 'RICHDS
    Case Is = "Variable inlet vanes - 80%"
    Correction = Array(8, 5, 4, 4, 4, 4, 4) 'RICHDS
    Case Is = "Variable inlet vanes - 60%"
    Correction = Array(8, 7, 6, 5, 5, 5, 5) 'RICHDS
    Case Is = "Variable inlet vanes - 40%"
    Correction = Array(3, 2, 1, 0, 0, 0, 1) 'RICHDS
    Case Is = ""
    LwFanSimple = LwSimple
    End Select
    
    Select Case freq
    Case Is = "63"
    LwFanSimple = LwOverall + Correction(0)
    Case Is = "125"
    LwFanSimple = LwOverall + Correction(1)
    Case Is = "250"
    LwFanSimple = LwOverall + Correction(2)
    Case Is = "500"
    LwFanSimple = LwOverall + Correction(3)
    Case Is = "1k"
    LwFanSimple = LwOverall + Correction(4)
    Case Is = "2k"
    LwFanSimple = LwOverall + Correction(5)
    Case Is = "4k"
    LwFanSimple = LwOverall + Correction(6)
    Case Else
    LwFanSimple = ""
    End Select
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Here be subs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub PutLwFanSimple(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
frmEstFanLw.Show

    If btnOkPressed Then 'ok
    
    Call ParameterUnmerge(Selection.Row, SheetType)
    
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = FanV
        Cells(Selection.Row, 15).Value = FanP
        Cells(Selection.Row, 5).Value = "=LwFanSimple(E$6,$N" & Selection.Row & ",$O" & Selection.Row & ",""" & FanType & """)"
        ExtendFunction (SheetType)
        Cells(Selection.Row, 14).NumberFormat = "0""m" & chr(179) & "/s"""
        Cells(Selection.Row, 15).NumberFormat = "0""Pa"""
        Cells(Selection.Row, 2).Value = "Fan Estimate - Simple"
        Else 'Third octave - not implemented
        End If
    
    fmtUserInput SheetType, True
    End If

End Sub



Sub PutLwPumpSimple(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
frmEstPumpLw.Show

    If btnOkPressed Then 'ok
    
    Call ParameterMerge(Selection.Row, SheetType)
    
        If Left(SheetType, 3) = "OCT" Then
        PumpEqn = Right(PumpEqn, Len(PumpEqn) - 3)
        PumpEqn = Replace(PumpEqn, "kW", "$N" & Selection.Row, 1, Len(PumpEqn), vbTextCompare)
        Cells(Selection.Row, 5).Value = "=" & PumpEqn & "-13"
        Cells(Selection.Row, 6).Value = "=" & PumpEqn & "-12"
        Cells(Selection.Row, 7).Value = "=" & PumpEqn & "-11"
        Cells(Selection.Row, 8).Value = "=" & PumpEqn & "-9"
        Cells(Selection.Row, 9).Value = "=" & PumpEqn & "-9"
        Cells(Selection.Row, 10).Value = "=" & PumpEqn & "-6"
        Cells(Selection.Row, 11).Value = "=" & PumpEqn & "-9"
        Cells(Selection.Row, 12).Value = "=" & PumpEqn & "-13"
        Cells(Selection.Row, 13).Value = "=" & PumpEqn & "-19"
        Cells(Selection.Row, 14).Value = PumpPower
        'ExtendFunction (SheetType)
        Cells(Selection.Row, 14).NumberFormat = "0"" kW"""
        Cells(Selection.Row, 2).Value = DescriptionString 'set by form code
        Else
        'Third octaves not provided
        End If
    fmtUserInput SheetType, True
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    
    'Assume spherical spreading
    Distance (SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = 1 'assume 1m
        FlipSign SheetType, True 'skip user input
        Else
        'Third octaves not provided
        End If
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    Cells(Selection.Row, 5).Value = "=" & Cells(Selection.Row - 2, 5).Address(False, False) & "+" & Cells(Selection.Row - 1, 5).Address(False, False)
    ExtendFunction (SheetType)
    Cells(Selection.Row, 2).Value = "SWL - Pump"
    
    End If 'close if statement for btnOK

End Sub

Sub PutLwCoolingTower(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmEstCoolingTower.Show
    If btnOkPressed Then 'ok
    Call ParameterMerge(Selection.Row, SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = CTPower
        CTEqn = Right(CTEqn, Len(CTEqn) - 2) 'chop off "Lw", start with "="
        CTEqn = Replace(CTEqn, "kW", "$N" & Selection.Row, 1, Len(CTEqn), vbTextCompare)
        CTEqn = Replace(CTEqn, "log(", "*LOG(", 1, Len(CTEqn), vbTextCompare)
        Cells(Selection.Row, 5).Value = CTEqn
        ExtendFunction (SheetType)
            'apply correction
            For C = LBound(CT_Correction) To UBound(CT_Correction)
                If CT_Correction(C) >= 0 Then
                Cells(Selection.Row, 5 + C).Formula = Cells(Selection.Row, 5 + C).Formula & "+" & CStr(CT_Correction(C))
                Else
                Cells(Selection.Row, 5 + C).Formula = Cells(Selection.Row, 5 + C).Formula & CStr(CT_Correction(C)) 'minus already in there
                End If
            Next C
        Cells(Selection.Row, 14).NumberFormat = "0"" kW"""
        Cells(Selection.Row, 2).Value = "Cooling Tower Estimate - " & CT_Type & " Type"
        Else
        End If
        
    fmtUserInput SheetType, True
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    
     'Assume spherical spreading
    Distance (SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = 6 'assume minimum distance 6m
        Range(Cells(Selection.Row, 14), Cells(Selection.Row, 14)).ClearComments
        Range(Cells(Selection.Row, 14), Cells(Selection.Row, 14)).AddComment ("Minimum distance: 6m")
        'TODO
        End If
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
        
        If CT_Dir_checked = True Then
        Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)) = CT_Direction
        Cells(Selection.Row, 2).Value = CT_Direction(9)
        'move down one row
        Cells(Selection.Row + 1, Selection.Column).Select
        End If
        
    
    'add it up!
    AutoSum (SheetType)
    
    End If 'close if statement for btnOK
End Sub

Sub PutCompressorSmall(SheetType As String)

Dim i As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmEstCompressorSmall.Show

    If btnOkPressed Then 'ok
        If Left(SheetType, 3) = "OCT" Then
            For i = 0 To 8
            Cells(Selection.Row, 5 + i).Formula = Cells(Selection.Row, 5 + i).Formula & "+" & CompressorSPL(i)
            Next i
        Else
        ErrorOctOnly
        End If
    End If
Cells(Selection.Row, 2).Value = "Compressor (small) SPL Estimate"

'move down one row
Cells(Selection.Row + 1, Selection.Column).Select

'Assume spherical spreading
Distance (SheetType)
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 14).Value = 1 'assume minimum distance 6m
    FlipSign SheetType, True 'skip user input
    Else
    ErrorOctOnly
    End If

'move down one row
Cells(Selection.Row + 1, Selection.Column).Select
'Add divergence formula
Cells(Selection.Row, 5).Value = "=" & Cells(Selection.Row - 2, 5).Address(False, False) & "+" & Cells(Selection.Row - 1, 5).Address(False, False)
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "SWL - Compressor"

End Sub

Sub PutElectricMotorSmall(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmEstElectricMotorSmall.Show

    If btnOkPressed Then 'ok
    Call ParameterMerge(Selection.Row, SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = MotorPower
        Cells(Selection.Row, 14).NumberFormat = "0""kW"""
        
        MotorEqn = Right(MotorEqn, Len(MotorEqn) - 3)
        MotorEqn = Replace(MotorEqn, "kW", "$N" & Selection.Row, 1, Len(MotorEqn), vbTextCompare)
        MotorEqn = Replace(MotorEqn, "RPM", MotorSpeed, 1, Len(MotorEqn), vbTextCompare)
        Cells(Selection.Row, 5).Value = "=" & MotorEqn
        ExtendFunction (SheetType)
            For corNum = 0 To 8
                If Motor_Correction(corNum) >= 0 Then
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & "+" & Motor_Correction(corNum)
                Else
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & Motor_Correction(corNum)  'number includes minus sign
                End If
            Next corNum
        Else
        ErrorOctOnly
        End If
    End If
    
    fmtUserInput SheetType, True
    
    Cells(Selection.Row, 2).Value = "Electric Motor SPL Estimate - " & MotorType & " Type"

 'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    
     'Assume spherical spreading
    Distance (SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = 1 'assume minimum distance 6m
        Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)).Select
        FlipSign SheetType, True 'skip user input
        End If
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    'Add divergence formula
    Cells(Selection.Row, 5).Value = "=" & Cells(Selection.Row - 2, 5).Address(False, False) & "+" & Cells(Selection.Row - 1, 5).Address(False, False)
    ExtendFunction (SheetType)
    Cells(Selection.Row, 2).Value = "SWL - Motor"

End Sub


Sub PutGasTurbine(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmEstGasTurbine.Show

    If btnOkPressed Then 'ok
    Call ParameterMerge(Selection.Row, SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = TurbinePower
        Cells(Selection.Row, 14).NumberFormat = "0""MW"""
        
        TurbineEqn = Right(TurbineEqn, Len(TurbineEqn) - 3)
        TurbineEqn = Replace(TurbineEqn, "MW", "$N" & Selection.Row, 1, Len(TurbineEqn), vbTextCompare)
        Cells(Selection.Row, 5).Value = "=" & TurbineEqn
        ExtendFunction (SheetType)
            For corNum = 0 To 8
                If TurbineCorrection(corNum) >= 0 Then
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & "+" & TurbineCorrection(corNum)
                Else
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & TurbineCorrection(corNum)  'number includes minus sign
                End If
            Cells(Selection.Row + 1, 5 + corNum).Value = TurbineEnclosure(corNum)
            Next corNum
        Else
        ErrorOctOnly
        End If
        
    fmtUserInput SheetType, True
    
    Cells(Selection.Row, 2).Value = "Gas Turbine SWL Estimate - " & GasTurbineType
    Cells(Selection.Row + 1, 2).Value = "Turbine Enclosure - " & EnclosureDescription
    'move down
    Cells(Selection.Row + 2, 2).Select
    AutoSum (SheetType)

    End If
End Sub


Sub PutSteamTurbine(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmEstSteamTurbine.Show

    If btnOkPressed Then 'ok
    Call ParameterMerge(Selection.Row, SheetType)
        If Left(SheetType, 3) = "OCT" Then
        Cells(Selection.Row, 14).Value = TurbinePower
        Cells(Selection.Row, 14).NumberFormat = "0""kW"""
        
        TurbineEqn = Right(TurbineEqn, Len(TurbineEqn) - 3)
        TurbineEqn = Replace(TurbineEqn, "kW", "$N" & Selection.Row, 1, Len(TurbineEqn), vbTextCompare)
        Cells(Selection.Row, 5).Value = "=" & TurbineEqn
        ExtendFunction (SheetType)
            For corNum = 0 To 8
                If TurbineCorrection(corNum) >= 0 Then
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & "+" & TurbineCorrection(corNum)
                Else
                Cells(Selection.Row, 5 + corNum).Formula = Cells(Selection.Row, 5 + corNum).Formula & TurbineCorrection(corNum)  'number includes minus sign
                End If
            Cells(Selection.Row + 1, 5 + corNum).Value = TurbineEnclosure(corNum)
            Next corNum
        Else
        ErrorOctOnly
        End If
        
    fmtUserInput SheetType, True
    
    Cells(Selection.Row, 2).Value = "Steam Turbine SWL Estimate"
     Cells(Selection.Row + 1, 2).Value = "Turbine Enclosure - " & EnclosureDescription
    'move down
    Cells(Selection.Row + 2, 2).Select
    AutoSum (SheetType)
    End If
End Sub
