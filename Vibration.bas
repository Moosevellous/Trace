Attribute VB_Name = "Vibration"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public VibRef As String
Public ConversionFactorStr As String
Public VibConversionDescription As String
Public BuildingType As String
Public AmplificationType As String

Public AS2670_Axis As String
Public AS2670_Multiplier As Single
Public AS2670_Order
Public AS2670_dbUnit As Boolean
Public AS2670_RateCurve As Boolean
Public VibRateAddr As String 'address of range to be rated

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     VCcurve
' Author:   PS
' Desc:     Returns value of ASHRAE VC curves at one-third octave frequency
' Args:     CurveName - VC-A VC-B etc
'           freq - one third octave band centre frequency
'           Mode - dB mode, input string "dB"
' Comments: (1)
'==============================================================================
Function VCcurve(CurveName As String, freq As String, Optional Mode As String)

Dim VC_OR() As Variant
Dim VC_A() As Variant
Dim VC_B() As Variant
Dim VC_C() As Variant
Dim VC_D() As Variant
Dim VC_E() As Variant
Dim i As Integer
Dim f As Double
Dim ChosenCurve() As Variant

'bands         2       2.5     3.15    4       5       6.3     8       10 _
'     12.5     16      20      25      31.5    40      50      63      80
VC_E = Array(0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, _
    0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032)
VC_D = Array(0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, _
    0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064)
VC_C = Array(0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, _
    0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013)
VC_B = Array("-", "-", "-", 0.05, 0.0397, 0.0315, 0.025, 0.025, _
    0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025)
VC_A = Array("-", "-", "-", 0.102, 0.081, 0.0643, 0.051, 0.051, _
    0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051)
VC_OR = Array(0.306, 0.2548, 0.2122, 0.1767, 0.1471, 0.1225, 0.102, 0.102, _
    0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102)
    
'Debug.Print CurveName

'ChosenCurve = ""
    Select Case CurveName
    Case "VC-OR"
    ChosenCurve = VC_OR
    Case "VC-A"
    ChosenCurve = VC_A
    Case "VC-B"
    ChosenCurve = VC_B
    Case "VC-C"
    ChosenCurve = VC_C
    Case "VC-D"
    ChosenCurve = VC_D
    Case "VC-E"
    ChosenCurve = VC_E
    Case Is = ""
    ChosenCurve = Array(0, 0, 0, 0, 0, 0, 0, 0, _
                        0, 0, 0, 0, 0, 0, 0, 0, 0)
    End Select

f = freqStr2Num(freq)
i = GetArrayIndex_TO(f, 14) '14 bands offset from 50Hz to 2Hz

VCcurve = "-" 'catch for errors
    If i > 0 And i < UBound(ChosenCurve) + 1 Then
    VCcurve = ChosenCurve(i)
    End If

    'dB mode
    If VCcurve <> "-" And Mode = "dB" Then
    VCcurve = 20 * Application.WorksheetFunction.Log10(VCcurve / 0.000001)
    End If

End Function


'==============================================================================
' Name:     VcRate
' Author:   PS
' Desc:     Rates vibration spectrum against ASHRAE VC curves
' Args:     CurveName - VC-OR / VC-A / VC-B etc
'           freqTable - one third octave band centre frequencies
'           Mode - set to "dB" to rate against dB
' Comments: (1) TODO: implement dB mode
'==============================================================================
Function VcRate(DataTable As Variant, freqTable As Variant, _
    Optional Mode As String)

Dim MaxCurve As Integer
Dim CurrentCurve As Integer
Dim i As Integer

MaxCurve = 0

MapValue = Array("VC-E", "VC-D", "VC-C", "VC-B", "VC-A", "VC-OR")
    
    For i = 0 To 26 '26 columns is all you'll need
    
        Select Case DataTable(i)
        
        Case Is > VCcurve("VC-OR", CStr(freqTable(i)))
        CurrentCurve = 6
        Case Is > VCcurve("VC-A", CStr(freqTable(i)))
        CurrentCurve = 5
        Case Is > VCcurve("VC-B", CStr(freqTable(i)))
        CurrentCurve = 4
        Case Is > VCcurve("VC-C", CStr(freqTable(i)))
        CurrentCurve = 3
        Case Is > VCcurve("VC-D", CStr(freqTable(i)))
        CurrentCurve = 2
        Case Is > VCcurve("VC-E", CStr(freqTable(i)))
        CurrentCurve = 1
        End Select
        
        If CurrentCurve > MaxCurve Then
        MaxCurve = CurrentCurve
        End If
        
    Next i

VcRate = MapValue(MaxCurve)

End Function

'==============================================================================
' Name:     AS2670_Curve
' Author:   AA
' Desc:     Returns 1/3 octave value according to AS2670 vibration curve
' Args:     Axis--{z, xy, combined}. Choose which axis to return
'           Multiplier--any number. {1, 1.4, 2, 4, 8} multipliers correspond to
'               {Baseline, Residential Night, Residential Day, Office,
'               Workshop}
'           freq--single value freq band, one of 1/3 oct AccelVel headers
'           AccelVel--Specifies which curve to return. Must be either of
'               {"Accel", "Vel"}.
'           Mode = Optional, specified dB result, can be left out to return
'               linear value i.e. inputs are either of {"", "dB"}.
' Comments:
'==============================================================================
' GENERAL SETUP

Function AS2670_Curve(Axis As String, Multiplier As Variant, freq As String, _
    AccelVel As String, Optional Mode As String)
    
Dim i As Integer
Dim f As Double

'------------------------------------------------------------------------------
' AS2670 REFERENCE DATA

'Reference vibration curves - ACCELERATION for multipliers {1, 1.4, 2, 4, 8}
'1hz, 1.25hz, 1.6hz, 2hz, 2.5hz, 3.15hz, 4hz, 5hz, 6.3hz, 8hz, 10hz, 12.5hz,
'   16hz, 20hz, 25hz, 31.5hz, 40hz, 50hz, 63hz, 80hz
Curve_Accel_z = Array(0.01, 0.0089, 0.008, 0.007, 0.0063, 0.0057, 0.005, _
    0.005, 0.005, 0.005, 0.0063, 0.00781, 0.01, 0.0125, 0.0156, 0.0197, _
    0.025, 0.0313, 0.0394, 0.05)
Curve_Accel_xy = Array(0.0036, 0.0036, 0.0036, 0.0036, 0.00451, 0.00568, _
    0.00721, 0.00902, 0.0114, 0.0144, 0.018, 0.0225, 0.0289, 0.0361, 0.0451, _
    0.0568, 0.0721, 0.0902, 0.114, 0.144)
Curve_Accel_combined = Array(0.0036, 0.0036, 0.0036, 0.0036, 0.00372, _
    0.00387, 0.00407, 0.0043, 0.0046, 0.005, 0.0063, 0.0078, 0.01, 0.0125, _
    0.0156, 0.0197, 0.025, 0.0313, 0.0394, 0.05)

'Reference vibration curves - VELOCITY (RMS) for multipliers {1, 1.4, 2, 4, 8}
'1hz, 1.25hz, 1.6hz, 2hz, 2.5hz, 3.15hz, 4hz, 5hz, 6.3hz, 8hz, 10hz, 12.5hz,
'   16hz, 20hz, 25hz, 31.5hz, 40hz, 50hz, 63hz, 80hz
Curve_Vel_z = Array(0.00159, 0.00113, 0.000796, 0.000557, 0.000401, 0.000288, _
    0.000199, 0.000159, 0.000126, 0.0000995, 0.0000995, 0.0000995, 0.0000995, _
    0.0000995, 0.0000995, 0.0000995, 0.0000995, 0.0000995, 0.0000995, _
    0.0000995)
Curve_Vel_xy = Array(0.000573, 0.000458, 0.000358, 0.000287, 0.000287, _
    0.000287, 0.000287, 0.000287, 0.000287, 0.000287, 0.000287, 0.000287, _
    0.000287, 0.000287, 0.000287, 0.000287, 0.000287, 0.000287, 0.000287, _
    0.000287)
Curve_Vel_combined = Array(0.000573, 0.000458, 0.000358, 0.000287, 0.000237, _
    0.000195, 0.000162, 0.000136, 0.000116, 0.0000995, 0.0000995, 0.0000995, _
    0.0000995, 0.0000995, 0.0000995, 0.0000995, 0.0000995, 0.0000995, _
    0.0000995, 0.0000995)

    If IsMissing(Mode) Then Mode = ""

'------------------------------------------------------------------------------
' MAIN

    'Selection of reference curve to display.
    If Axis = "z" And AccelVel = "Accel" Then
        ChosenCurve = Curve_Accel_z
    ElseIf Axis = "xy" And AccelVel = "Accel" Then
        ChosenCurve = Curve_Accel_xy
    ElseIf Axis = "comb." And AccelVel = "Accel" Then
        ChosenCurve = Curve_Accel_combined
    ElseIf Axis = "z" And AccelVel = "Vel" Then
        ChosenCurve = Curve_Vel_z
    ElseIf Axis = "xy" And AccelVel = "Vel" Then
        ChosenCurve = Curve_Vel_xy
    ElseIf Axis = "comb." And AccelVel = "Vel" Then
        ChosenCurve = Curve_Vel_combined
    End If

' Catch for errors/non-values
AS2670_Curve = "-"

    ' Multiply baseline curve by multiplier
    If Multiplier = "NONE" Then
        Exit Function
    Else
        For i = 0 To UBound(ChosenCurve)
            ChosenCurve(i) = ChosenCurve(i) * Multiplier
            'NOTE:  If {Multiplier} is neither equal to "NONE" or a number, this
            '       line will cause a #VALUE error in the cell.
        Next
    End If

' Find vibration corresponding to freq to display
f = freqStr2Num(freq)
i = GetArrayIndex_TO(f, 17) '17 bands offset from 50Hz to 1Hzo
    If i >= 0 And i < UBound(ChosenCurve) + 1 Then
    AS2670_Curve = ChosenCurve(i)
    End If

    'Converts to dB units if Mode is "dB"
    If AS2670_Curve <> "-" And Mode = "dB" And AccelVel = "Accel" Then
        AS2670_Curve = _
            20 * Application.WorksheetFunction.Log10(AS2670_Curve / 0.001)
    ElseIf AS2670_Curve <> "-" And Mode = "dB" And AccelVel = "Vel" Then
        AS2670_Curve = _
            20 * Application.WorksheetFunction.Log10(AS2670_Curve / 0.000001)
    End If

End Function


'==============================================================================
' Name:     AS2670_Rate
' Author:   AA
' Desc:     Returns AS2670 low frequency vibration curve with which the
'           selected data complies
' Args:     DataTable--selected data cells
'           freqTable--selected freq cells corresponding to data cells
'           Axis--Choose assessment axis. Must be one of the following
'               {z, xy, combined}
'           AccelVel--order of vibration, must be one of the following
'               {"Accel", "Vel"}
' Comments: (1)
'==============================================================================
' GENERAL SETUP

Function AS2670_Rate(DataTable As Variant, freqTable As Variant, _
    Axis As String, AccelVel As String, Optional Mode As String)
    
Dim MaxCurve As Integer
Dim CurrentCurve As Integer
Dim MapValue As Variant
Dim i As Integer

' Error catching default value for max vibration curve met
MaxCurve = 0

' Set array of values to return. Corresponds to the multipliers in AS2670.
MapValue = Array(1, 1.4, 2, 4, 8, "NONE")

    If IsMissing(Mode) Then Mode = ""

    ' For 1hz to 400hz columns
    For i = 1 To 27
        'selection of which AS2670 curve is exceeded, {CurrentCurve} value is
        'used later to return {MapValue} array item
        Select Case DataTable(i)
        Case Is > AS2670_Curve(Axis, 8, CStr(freqTable(i)), AccelVel, Mode)
            CurrentCurve = 5
        Case Is > AS2670_Curve(Axis, 4, CStr(freqTable(i)), AccelVel, Mode)
            CurrentCurve = 4
        Case Is > AS2670_Curve(Axis, 2, CStr(freqTable(i)), AccelVel, Mode)
            CurrentCurve = 3
        Case Is > AS2670_Curve(Axis, 1.4, CStr(freqTable(i)), AccelVel, Mode)
            CurrentCurve = 2
        Case Is > AS2670_Curve(Axis, 1, CStr(freqTable(i)), AccelVel, Mode)
            CurrentCurve = 1
        End Select
        
        'sets maximum curve which is met
        If CurrentCurve > MaxCurve Then
            MaxCurve = CurrentCurve
        End If
    Next i

' Return curve that is met
AS2670_Rate = MapValue(MaxCurve)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'==============================================================================
' Name:     VibLin2DB
' Author:   PS
' Desc:     Converts mm/s to dBV
' Args:     None
' Comments: (1) Variable VibRef is set in frmVibUnits
'==============================================================================
Sub VibLin2DB()
'set form title bar
frmVibUnits.Caption = "Vibration - Convert Units (Linear to dB)"
frmVibUnits.Show
    If btnOkPressed = False Then End
'build formula
SetDescription "Convert to dB"
BuildFormula "20*LOG(" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "/" & _
    VibRef & ")"

End Sub

'==============================================================================
' Name:     VibLin2DB
' Author:   PS
' Desc:     Converts dBV to mm/s
' Args:     None
' Comments: (1) Variable VibRef is set in frmVibUnits
'==============================================================================
Sub VibDB2Lin()
'set form title bar
frmVibUnits.Caption = "Vibration - Convert Units (dB to Linear)"
frmVibUnits.Show
    If btnOkPressed = False Then End
'build formula
SetDescription "Convert to Linear"
BuildFormula "" & VibRef & "*10^(" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "/20)"

End Sub

'==============================================================================
' Name:     CouplingLoos
' Author:   PS
' Desc:     Inserts Coupling Loss values, in one-third octave bands
' Args:     None
' Comments: (1) Coupling loss values have been obtained from Nelson and have
'           been extrapolated to include frequency bands below 16Hz
'           (2) Nelson = Transportation Noise Reference Book, Nelson, P (1987)
'==============================================================================
Sub PutCouplingLoss()

Dim StartFreq As Integer

'SET VARIABLES
'bands     5 6.3 8 10 12.5 16 20 25 31.5 40 50 63 80 100 125 160 200 250 315
CRL = Array(2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2, 2, 2)
'bands                      5 6.3 8 10 12.5 16 20 25 31 40 50  63  80 100 125
'   160 200 250 315
LargeMasonryOnPiles = Array(6, 6, 6, 6, 7, 7, 7, 8, 9, 10, 11, 12, 13, 13, 14, _
    14, 15, 15, 15)
'bands                               5  6.3  8   10 12.5 16  20  25 31.5 40
'  50  63  80 100 125 160 200  250 315
LargeMasonryOnSpreadFootings = Array(11, 11, 11, 11, 12, 13, 14, 14, 15, 15, _
   15, 15, 14, 14, 14, 14, 13, 12, 11)
'bands                                         5 6.3 8 10 12.5 16 20 25 31.5 40
'   50 63 80 100 125 160 200 250 315
TwoToFourStoreyMasonryOnSpreadFootings = Array(5, 6, 6, 7, 9, 11, 11, 12, 13, 13, _
    13, 13, 13, 12, 12, 11, 10, 9, 8)
'bands                          5 6.3 8 10 12.5 16 20 25 31.5 40 50 63 80 100
'   125 160 200 250 315
OneToTwoStoreyCommercial = Array(4, 5, 5, 6, 7, 8, 8, 9, 9, 9, 9, 9, 9, 8, _
    8, 8, 7, 6, 5)
'bands           5 6.3 8 10 12.5 16 20 25 31.5 40 50 63 80 100 125 160 200 250 315
SingleResidential = Array(3, 3, 4, 4, 5, 5, 6, 6, 6, 6, 6, 6, 6, 5, 5, 5, 4, 4, 4)

'Low frequency third-octave sheet check
If T_SheetType <> "LF_TO" Then ErrorLFTOOnly

frmCouplingLoss.Show

If btnOkPressed = False Then End

    Select Case BuildingType 'public variable
    Case Is = "CRL"
    SelectedLoss = CRL
    Case Is = "Large Masonry On Piles"
    SelectedLoss = LargeMasonryOnPiles
    Case Is = "Large Masonry on Spread Footings"
    SelectedLoss = LargeMasonryOnSpreadFootings
    Case Is = "2-4 Storey Masonry on Spread Footings"
    SelectedLoss = TwoToFourStoreyMasonryOnSpreadFootings
    Case Is = "1-2 Storey Commercial"
    SelectedLoss = OneToTwoStoreyCommercial
    Case Is = "Single Residential"
    SelectedLoss = SingleResidential
    End Select
    
    If IsEmpty(SelectedLoss) Then End
    
    'insert values all start from 5Hz band
    StartFreq = FindFrequencyBand("5")
    For i = LBound(SelectedLoss) To UBound(SelectedLoss)
    Cells(Selection.Row, StartFreq + i).Value = -1 * SelectedLoss(i) 'negative values!
    Next i

SetDescription "Coupling Loss: " & BuildingType
    
End Sub


'==============================================================================
' Name:     Building Amplification
' Author:   PS
' Desc:     Puts in amplification values into buildings for vibration and GBN
' Args:     None
' Comments: (1) Note: the frequency range used for vibration assessment is 5Hz
'           to 80Hz and the frequency range for ground-borne noise assessment
'           is 20Hz to 315Hz.
'           (2) ANC Guidelines - Measurement and Assessment of Ground-borne
'           Noise & Vibration, Association of Noise Consultants (2001)
'==============================================================================
Sub BuildingAmplification()
Dim StartFreq As Integer
'bands                 5  6.3  8   10 12.5 16  20  25 31.5 40  50 63 80 100 125
'   160 200 250 315
FloorVibration = Array(10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 10, 9, 9, 0, 0, _
    0, 0, 0, 0)
'bands     5 6.3 8 10 12.5 16 20 25 31.5 40 50 63 80 100 125 160 200 250 315
GBN = Array(0, 0, 0, 0, 0, 0, 6, 7, 7, 7, 6, 6, 5, 5, 4, 3, 2, 1, 1)

'Low frequency third-octave sheet check
If T_SheetType <> "LF_TO" Then ErrorLFTOOnly

frmBuildingAmplification.Show

If btnOkPressed = False Then End

    Select Case AmplificationType 'public variable
    Case Is = "Ground-borne Noise"
    SelectedLoss = GBN
    Case Is = "Floor Vibration"
    SelectedLoss = FloorVibration
    End Select
    
    If IsEmpty(SelectedLoss) Then End
    
    'insert values all start from 5Hz band
    StartFreq = FindFrequencyBand("5")
    For i = LBound(SelectedLoss) To UBound(SelectedLoss)
        If SelectedLoss(i) <> 0 Then
        Cells(Selection.Row, 12 + i).Value = SelectedLoss(i) 'negative values!
        End If
    Next i

SetDescription "Building Amplification: " & AmplificationType

End Sub

'==============================================================================
' Name:     PutVCcurve
' Author:   PS
' Desc:     Inserts rating formula and presents VC curve
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutVCcurve()

msg = MsgBox("Linear values (mm/s)? " & chr(10) & _
    "[Note that 'No' will choose dB mode.]", vbYesNoCancel, "Lin/Log mode")
If msg = vbCancel Then End

ParameterMerge (Selection.Row)

'Low frequency third-octave sheet check
    If T_SheetType <> "LF_TO" Then ErrorLFTOOnly

SetDescription "VC Curve"

'build formula
Cells(Selection.Row, T_ParamStart) = "VC-A"
    If msg = vbYes Then
    BuildFormula "VCcurve(" & T_ParamRng(0) & _
        "," & T_FreqStartRng & ")"
    ElseIf msg = vbNo Then 'dB mode
    BuildFormula "VCcurve(" & T_ParamRng(0) & _
        "," & T_FreqStartRng & ",""dB"")"
    End If

    
'format parameter columns
SetTraceStyle "Input", True

SetDataValidation T_ParamStart, "VC-OR,VC-A,VC-B,VC-C,VC-D,VC-E"

End Sub

'==============================================================================
' Name:     VibConvert
' Author:   PS
' Desc:     Inserts conversion between displacement, velocity, acceleration
' Args:     None
' Comments: (1)
'==============================================================================
Sub VibConvert()

Dim FormulaStr As String

frmVibConvert.Show
    
    If btnOkPressed = False Then End
    'Low frequency third-octave sheet check
    If T_SheetType <> "LF_TO" Then ErrorLFTOOnly

'build formula
FormulaStr = Replace(ConversionFactorStr, "pi", "PI()")
FormulaStr = Replace(FormulaStr, "f", T_FreqStartRng)
FormulaStr = Replace(FormulaStr, chr(178), "^2")
BuildFormula "" & FormulaStr

Range(Cells(Selection.Row, T_LossGainStart), _
    Cells(Selection.Row, T_LossGainEnd)).NumberFormat = "0E+0"
SetDescription "Vibration Conversion"
InsertComment VibConversionDescription, T_Description, False

SelectNextRow
'TODO: multiply or add to row above

'    'apply style
'    If BasicsApplyStyle <> "" Then
'    ApplyTraceStyle "Trace " & BasicsApplyStyle, Selection.Row
'    End If

End Sub


'==============================================================================
' Name:     PutAS2670
' Author:   AA
' Desc:     Interface sub for AS2670 functions. Either inserts a reference
'           vibration curve at the selected row or reads the previous row and
'           uses function AS2670_Rate to insert the corresponding vibration
'           curve, depending on user choice.
' Args:     TypeCode--Sheet type must be LF_TO
' Comments: (1) Requires frmAS2670 form to function
'           (2) Requires the following public variables to function:
'                   Public AS2670_Axis As String
'                   Public AS2670_Multiplier As Single
'                   Public AS2670_Order
'                   Public AS2670_dbUnit As Boolean
'                   Public AS2670_RateCurve As Boolean
'==============================================================================
' GENERAL SETUP

Sub PutAS2670curve()
Dim DataTable As Variant
Dim Mode As String, RowTitle As String, RowTitleUnit As String
Dim RateRow As Integer

    'Low frequency third-octave sheet check
    If T_SheetType <> "LF_TO" Then ErrorLFTOOnly

'------------------------------------------------------------------------------
' INTERFACE WITH POP-UP USERFORM FOR USER INPUT
frmAS2670.RefVibRange.Value = Cells(Selection.Row - 1, T_LossGainStart).Address

frmAS2670.Show

    ' catch error
    If btnOkPressed = False Then End
    
    ' Assign dB unit variable to local function for later use
    If AS2670_dbUnit = True Then Mode = "dB"

'------------------------------------------------------------------------------

' Parameter columns unmerge and formatting
ParameterUnmerge (Selection.Row)
SetTraceStyle "Input", True

' Parameter column 1 (Column AF) cell contents and format
Cells(Selection.Row, T_ParamStart) = AS2670_Axis
SetDataValidation T_ParamStart, "z, xy, comb."

' Parameter column 2 (Column AG) format
SetUnits "General", T_ParamStart + 1

    ' Parameter column 2 (Column AG) contents. If "Rate existing curve" button
    ' in the AS2670 form is selected, rate the row above the current row,
    ' otherwise use input multiplier provided by user.
    ' And Description title
    If AS2670_RateCurve = True Then
    RateRow = ExtractAddressElement(VibRateAddr, 2)
    Cells(Selection.Row, T_ParamStart + 1).Value = "=AS2670_Rate(" _
        & Range(Cells(RateRow, T_LossGainStart), _
        Cells(RateRow, T_ParamStart - 1)).Address(False, True) _
        & "," & Range(Cells(T_FreqRow, T_LossGainStart), _
        Cells(T_FreqRow, T_LossGainEnd)).Address(True, True) & "," _
        & T_ParamRng(0) & "," _
        & """" & AS2670_Order & """" & "," & """" & Mode & """" & ")"
    
        ' Formatting to normal
        With Cells(Selection.Row, T_ParamStart + 1)
        .Style = "Trace Normal"
        .Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
    RowTitle = "AS2670 Curve (" & AS2670_Order & "): "
    
    Else 'no rating, just the curve as presented
    Cells(Selection.Row, T_ParamStart + 1) = AS2670_Multiplier
    RowTitle = "Ref. AS2670 Curve (" & AS2670_Order & "): "
    End If


' Main body contents
BuildFormula "AS2670_Curve(" _
    & T_ParamRng(0) & "," & T_ParamRng(1) & "," _
    & Cells(T_FreqRow, 5).Address(True, False) & "," _
    & """" & AS2670_Order & """" & "," & """" & Mode & """" & ")"


' Main body format and Title component if dB units selected
    If AS2670_dbUnit = True Then
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_ParamStart - 1)).NumberFormat = "0.0"
    RowTitleUnit = "dB"
    Else
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_ParamStart - 1)).NumberFormat = "0.000"
    RowTitleUnit = "Linear"
    End If

' Assign title cell contents
SetDescription RowTitle & RowTitleUnit

End Sub


