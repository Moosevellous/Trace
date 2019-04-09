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

Public TurbineType As String
Public TurbinePower As Long
Public TurbineEqn As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function LwFanSimple(freq As String, v As Double, P As Double, FanType As String)

LwOverall = 10 * Application.WorksheetFunction.Log10(v) + 20 * Application.WorksheetFunction.Log10(P) + 40 'v in m^3, p in Pa
 
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


Public TurbineType As String

Public TurbineW As Long
Public TurbineEqn As String
Public TurbinePower As String
Public DescriptionString As String
Public Eqn As String
Public Power As Long
Public CT_Correction(0 To 8) As Long




Function LwTurbine(freq As String, W As Double, TurbineType As String, EnclosureType As String)

'Casing output sound power level is the same for steam and gas turbines
LwCasing = 120 + 5 * Application.WorksheetFunction.Log10(W)
            
Correction = Array(-10, -7, -5, -4, -4, -4, -4, -4) 'gas casing = steam overall correction

'add casing corrections based on B&H, octave bands from 63 to 8k Hz
    If TurbineType = "Gas" Then
        Select Case EnclosureType
        Case Is = ""
        Casingreduction = Array(0, 0, 0, 0, 0, 0, 0, 0) 'default to no casing reduction
        Case Is = "1"
        Casingreduction = Array(-2, -2, -3, -3, -3, -4, -5, -6)
        Case Is = "2"
        Casingreduction = Array(-5, -5, -6, -6, -7, -8, -9, -10)
        Case Is = "3"
        Casingreduction = Array(-1, -1, -2, -2, -2, -2, -3, -3)
        Case Is = "4"
        Casingreduction = Array(-4, -4, -5, -6, -7, -8, -8, -8)
        Case Is = "5"
        Casingreduction = Array(-7, -8, -9, -10, -11, -12, -13, -14)
        End Select
    End If

    Select Case freq
    Case Is = "63"
    LwTurbineCasing = LwCasing + CorrectionCasing(0) + Casingreduction(0)
    Case Is = "125"
    LwTurbineCasing = LwCasing + CorrectionCasing(1) + Casingreduction(1)
    Case Is = "250"
    LwTurbineCasing = LwCasing + CorrectionCasing(2) + Casingreduction(2)
    Case Is = "500"
    LwTurbineCasing = LwCasing + CorrectionCasing(3) + Casingreduction(3)
    Case Is = "1k"
    LwTurbineCasing = LwCasing + CorrectionCasing(4) + Casingreduction(4)
    Case Is = "2k"
    LwTurbineCasing = LwCasing + CorrectionCasing(5) + Casingreduction(5)
    Case Is = "4k"
    LwTurbineCasing = LwCasing + CorrectionCasing(6) + Casingreduction(6)
    Case Else
    LwTurbineCasing = ""
    End Select


'Inlet and outlet are gas turbine only, to be excluded for steam turbine output
LwInlet = 127 + 15 * Application.WorksheetFunction.Log10(W)

'Inlet SWL correction values from overall Lw
CorrectionInlet = Array(-19, -18, -17, -17, -14, -8, -3, -3)

    Select Case freq
    Case Is = "63"
    LwTurbineInlet = LwInlet + CorrectionInlet(0)
    Case Is = "125"
    LwTurbineInlet = LwInlet + CorrectionInlet(1)
    Case Is = "250"
    LwTurbineInlet = LwInlet + CorrectionInlet(2)
    Case Is = "500"
    LwTurbineInlet = LwInlet + CorrectionInlet(3)
    Case Is = "1k"
    LwTurbineInlet = LwInlet + CorrectionInlet(4)
    Case Is = "2k"
    LwTurbineInlet = LwInlet + CorrectionInlet(5)
    Case Is = "4k"
    LwTurbineInlet = LwInlet + CorrectionInlet(6)
    Case Else
    LwTurbineInlet = ""
    End Select

LwOutlet = 133 + 10 * Application.WorksheetFunction.Log10(W)

'Outlet SWL correction values from overall Lw
CorrectionOutlet = Array(-12, -8, -6, -6, -7, -9, -11, -15)

    Select Case freq
    Case Is = "63"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(0)
    Case Is = "125"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(1)
    Case Is = "250"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(2)
    Case Is = "500"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(3)
    Case Is = "1k"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(4)
    Case Is = "2k"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(5)
    Case Is = "4k"
    LwTurbineOutlet = LwOutlet + CorrectionOutlet(6)
    Case Else
    LwTurbineOutlet = ""
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
        Cells(Selection.Row, 14).NumberFormat = "0""m" & Chr(179) & "/s"""
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
        Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)).Select
        FlipSign (SheetType)
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
        FlipSign (SheetType)
        End If
    
    'move down one row
    Cells(Selection.Row + 1, Selection.Column).Select
    'Add divergence formula
    Cells(Selection.Row, 5).Value = "=" & Cells(Selection.Row - 2, 5).Address(False, False) & "+" & Cells(Selection.Row - 1, 5).Address(False, False)
    ExtendFunction (SheetType)
    Cells(Selection.Row, 2).Value = "SWL - Motor"

End Sub


