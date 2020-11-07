Attribute VB_Name = "SWL_Estimation"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================

Public DescriptionString As String

'Fans
Public FanType As String
Public FanV As Long
Public FanP As Long

'Pumps
Public PumpEqn As String
Public PumpPower As Long

'Cooling Towers
Public CTEqn As String
Public CTPower As Long
Public CT_Type As String
Public CT_Correction(0 To 8) As Long
Public CT_Directivity(0 To 9) As Variant
Public CT_Dir_checked As Boolean

'Electric Motors
Public MotorType As String
Public MotorEqn As String
Public MotorPower As Long
Public MotorSpeed As Long
Public Motor_Correction(0 To 8) As Long

'Tuurbines (steam and gas)
Public TurbinePower As Long
Public TurbineEqn As String
Public TurbineCorrection(0 To 9) As Long
Public TurbineEnclosure(0 To 9) As Long
Public GasTurbineType As String
Public EnclosureDescription As String

'Compressors
Public CompressorSPL(0 To 8) As Long

'Boilers
Public BoilerPower As Long
Public BoilerEqn As String
Public BoilerCorrection(0 To 9) As Long
Public BoilerType As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     FanSimple
' Author:   PS
' Desc:     Sound power from fans
' Args:     freq - octave band centre frequency string
'           V - Volumetric air flow in m^3/s
'           P - Pressure in Pascals
'           FanType - identifier string for different fan types
' Comments: (1)
'==============================================================================
Function LwFanSimple(freq As String, V As Double, P As Double, FanType As String)

LwOverall = 10 * Application.WorksheetFunction.Log10(V) + _
    20 * Application.WorksheetFunction.Log10(P) + 40
 
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     FanSimple
' Author:   PS
' Desc:     Adds simple fan equation estimation function from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub FanSimple()

frmEstFanLw.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterUnmerge (Selection.Row)
'place in values
Cells(Selection.Row, T_ParamStart).Value = FanV
Cells(Selection.Row, T_ParamStart + 1).Value = FanP
'build formula
Cells(Selection.Row, T_LossGainStart).Value = "=LwFanSimple(" & _
    T_FreqStartRng & "," & T_ParamRng(0) & "," & T_ParamRng(1) & _
    ",""" & FanType & """)"
ExtendFunction
'format parameter cells
SetUnits "mps", T_ParamStart
SetUnits "Pa", T_ParamStart + 1, 0
Cells(Selection.Row, T_Description).Value = "SWL Estimate - Fan Simple"
SetTraceStyle "Input", True

End Sub


'==============================================================================
' Name:     PumpSimple
' Author:   PS
' Desc:     Adds sound power estimation for pumps from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub PumpSimple()

frmEstPumpLw.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)
'build formulas
PumpEqn = Right(PumpEqn, Len(PumpEqn) - 3)
PumpEqn = Replace(PumpEqn, "kW", T_ParamRng(0), 1, Len(PumpEqn), vbTextCompare)
Cells(Selection.Row, T_LossGainStart).Value = "=" & PumpEqn & "-13"
Cells(Selection.Row, 6).Value = "=" & PumpEqn & "-12"
Cells(Selection.Row, 7).Value = "=" & PumpEqn & "-11"
Cells(Selection.Row, 8).Value = "=" & PumpEqn & "-9"
Cells(Selection.Row, 9).Value = "=" & PumpEqn & "-9"
Cells(Selection.Row, 10).Value = "=" & PumpEqn & "-6"
Cells(Selection.Row, 11).Value = "=" & PumpEqn & "-9"
Cells(Selection.Row, 12).Value = "=" & PumpEqn & "-13"
Cells(Selection.Row, 13).Value = "=" & PumpEqn & "-19"
Cells(Selection.Row, T_ParamStart).Value = PumpPower

'format parameter cells
SetUnits "kW", T_ParamStart
Cells(Selection.Row, T_Description).Value = DescriptionString 'set by form code
SetTraceStyle "Input", True

'move down one row
SelectNextRow

'Assume spherical spreading
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 1 'assume 1m
FlipSign

'move down one row and sum
SelectNextRow
AutoSum "Subtotal", "SWL Estimate - Pump"

End Sub


'==============================================================================
' Name:     CoolingTower
' Author:   PS
' Desc:     Sound power estimation for Cooling Towers from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub CoolingTower()

Dim i As Integer

frmEstCoolingTower.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)

'build formulas
Cells(Selection.Row, T_ParamStart).Value = CTPower
CTEqn = Right(CTEqn, Len(CTEqn) - 2) 'chop off "Lw", start with "="
CTEqn = Replace(CTEqn, "kW", T_ParamRng(0), 1, Len(CTEqn), vbTextCompare)
CTEqn = Replace(CTEqn, "log(", "*LOG(", 1, Len(CTEqn), vbTextCompare)
Cells(Selection.Row, T_LossGainStart).Value = CTEqn
ExtendFunction

    'apply correction
    For i = LBound(CT_Correction) To UBound(CT_Correction)
        If CT_Correction(i) >= 0 Then 'add a plus to the formula
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            "+" & CStr(CT_Correction(i))
        Else 'minus already in there
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            CStr(CT_Correction(i))
        End If
    Next i

SetUnits "kW", T_ParamStart
Cells(Selection.Row, T_Description).Value = "Cooling Tower SWL Estimate - " & _
    CT_Type & " Type"
    
SetTraceStyle "Input", True

'move down one row
SelectNextRow

 'Assume spherical spreading
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 6 'assume minimum distance 6m
InsertComment "Minimum distance: 6m", T_Description, False

'move down one row
SelectNextRow
    
    'add directional effects
    If CT_Dir_checked = True Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)) = CT_Directivity
    Cells(Selection.Row, T_Description).Value = CT_Directivity(9)
    'move down one row
    SelectNextRow
    End If

'add it up!
AutoSum "Subtotal", "Cooling Tower SPL"

End Sub


'==============================================================================
' Name:     CompressorSmall
' Author:   PS
' Desc:     Sound power estimation for Small Compressors from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub CompressorSmall()

Dim i As Integer

frmEstCompressorSmall.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

    For i = 0 To 8
    Cells(Selection.Row, T_LossGainStart + i).Formula = CompressorSPL(i)
    Next i

Cells(Selection.Row, T_Description).Value = "Compressor (small) - SPL Estimate"

'move down one row
SelectNextRow

'Assume spherical spreading
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 1
FlipSign

'move down one row
SelectNextRow
AutoSum "Subtotal", "SWL Estimate - Compressor"

End Sub



'==============================================================================
' Name:     ElectricMotorSmall
' Author:   PS
' Desc:     Sound power estimation for Small Motors from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub ElectricMotorSmall()

Dim i As Integer

frmEstElectricMotorSmall.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)

'motor power
Cells(Selection.Row, T_ParamStart).Value = MotorPower
SetUnits "kW", T_ParamStart
Range(T_ParamRng(0)).ClearComments
Range(T_ParamRng(0)).AddComment ("Maximum motor power: 300kW")

'build formula
MotorEqn = Right(MotorEqn, Len(MotorEqn) - 3) 'trim 'Lw='
MotorEqn = Replace(MotorEqn, "kW", T_ParamRng(0), 1, Len(MotorEqn), vbTextCompare)
MotorEqn = Replace(MotorEqn, "RPM", MotorSpeed, 1, Len(MotorEqn), vbTextCompare)
Cells(Selection.Row, T_LossGainStart).Value = "=" & MotorEqn
ExtendFunction
    For i = 0 To 8
        If Motor_Correction(i) >= 0 Then 'add a plus to the formula
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            "+" & Motor_Correction(i)
        Else 'minus already in there
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            Motor_Correction(i)
        End If
    Next i
    
SetTraceStyle "Input", True
    
Cells(Selection.Row, T_Description).Value = "Electric Motor SPL Estimate - " & _
    MotorType & " Type"

'move down one row
SelectNextRow
    
'Assume spherical spreading
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 1
FlipSign

'move down and sum
SelectNextRow
AutoSum "Subtotal", "SWL Estimate - Motor"

End Sub

'==============================================================================
' Name:     GasTurbine
' Author:   PS
' Desc:     Sound power estimation for Gas Turbines from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub GasTurbine()

Dim i As Integer

frmEstGasTurbine.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)

Cells(Selection.Row, T_ParamStart).Value = TurbinePower
SetUnits "MW", T_ParamStart
'build formula
TurbineEqn = Right(TurbineEqn, Len(TurbineEqn) - 3) 'trim 'Lw='
TurbineEqn = Replace(TurbineEqn, "MW", T_ParamRng(0), 1, Len(TurbineEqn), _
    vbTextCompare)
Cells(Selection.Row, T_LossGainStart).Value = "=" & TurbineEqn
ExtendFunction
    For i = 0 To 8
        If TurbineCorrection(i) >= 0 Then 'add a plus to the formula
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & "+" & _
            TurbineCorrection(i)
        Else 'minus already in there
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
        Cells(Selection.Row, T_LossGainStart + i).Formula & _
        TurbineCorrection(i)
        End If
    Cells(Selection.Row + 1, T_LossGainStart + i).Value = TurbineEnclosure(i)
    Next i
    
SetTraceStyle "Input", True

Cells(Selection.Row, T_Description).Value = "SWL Estimate - Gas Turbine - " _
    & GasTurbineType
Cells(Selection.Row + 1, T_Description).Value = "Turbine Enclosure - " _
    & EnclosureDescription
'move down and sum
Cells(Selection.Row + 2, T_Description).Select
AutoSum
Cells(Selection.Row, T_Description).Value = "SWL Estimate - Gas Turbine"
End Sub

'==============================================================================
' Name:     SteamTurbine
' Author:   PS
' Desc:     Sound power estimation for Steam Turbines from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub SteamTurbine()

Dim i As Integer

frmEstSteamTurbine.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)

Cells(Selection.Row, T_ParamStart).Value = TurbinePower
SetUnits "kW", T_ParamStart
'build formula
TurbineEqn = Right(TurbineEqn, Len(TurbineEqn) - 3) 'trim 'Lw='
TurbineEqn = Replace(TurbineEqn, "kW", T_ParamRng(o), 1, Len(TurbineEqn), _
    vbTextCompare)
Cells(Selection.Row, T_LossGainStart).Value = "=" & TurbineEqn
ExtendFunction
    For i = 0 To 8
        If TurbineCorrection(i) >= 0 Then 'add a plus to the formula
        Cells(Selection.Row, T_ParamStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & "+" _
            & TurbineCorrection(i)
        Else 'minus already in there
        Cells(Selection.Row, T_ParamStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula _
            & TurbineCorrection(i)
        End If
    Cells(Selection.Row + 1, T_ParamStart + i).Value = TurbineEnclosure(i)
    Next i

    
SetTraceStyle "Input", True

Cells(Selection.Row, T_Description).Value = "SWL Estimate - Steam Turbine"
Cells(Selection.Row + 1, T_Description).Value = "Turbine Enclosure - " _
    & EnclosureDescription
'move down and sum
Cells(Selection.Row + 2, 2).Select
AutoSum
Cells(Selection.Row, T_Description).Value = "SWL Estimate - Steam Turbine"
End Sub


'==============================================================================
' Name:     Boiler
' Author:   PS
' Desc:     Sound power estimation for Boilers from B&H
' Args:     None
' Comments: (1)
'==============================================================================
Sub Boiler()

Dim i As Integer

frmEstBoiler.Show

If btnOkPressed = False Then End
If T_BandType <> "oct" Then ErrorOctOnly

ParameterMerge (Selection.Row)

Cells(Selection.Row, T_ParamStart).Value = BoilerPower
BoilerEqn = Right(BoilerEqn, Len(BoilerEqn) - 3) 'trim 'Lw='
    'build formula based on boiler type
    If BoilerType = "General Purpose" Then
    SetUnits "kW", T_ParamStart
    BoilerEqn = Replace(BoilerEqn, "kW", T_ParamRng(0), 1, Len(BoilerEqn), _
    vbTextCompare) 'for General, input is kW
    ElseIf BoilerType = "Large Power Plant" Then
    SetUnits "MW", T_ParamStart
    BoilerEqn = Replace(BoilerEqn, "MW", T_ParamRng(0), 1, Len(BoilerEqn), _
    vbTextCompare) 'for Large power plants, input is MW
    Else
    msg = MsgBox("Error - nothing selected??????", vbOKOnly, "How????")
    End If

Cells(Selection.Row, T_LossGainStart).Value = "=" & BoilerEqn
ExtendFunction

    For i = 0 To 8
        If BoilerCorrection(i) >= 0 Then 'add a plus to the formula
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            "+" & BoilerCorrection(i)
        Else 'minus already in there
        Cells(Selection.Row, T_LossGainStart + i).Formula = _
            Cells(Selection.Row, T_LossGainStart + i).Formula & _
            BoilerCorrection(i)
        End If
    Next i
    
SetTraceStyle "Input", True

Cells(Selection.Row, T_Description).Value = "SWL Estimate - Boiler - " _
    & BoilerType

End Sub
