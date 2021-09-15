Attribute VB_Name = "Noise"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================

'plane waves
Public PlaneH As Double
Public PlaneL As Double
Public PlaneDist As Double

'rooms
Public roomType As String
Public roomL As Double
Public roomW As Double
Public roomH As Double
Public roomLossType As String
Public OffsetDistance As Double


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     AirAbsorb
' Author:   PS
' Desc:     Sound energy absorbed per km of air, interpolated to metres
' Args:     freq (frequency band), Distance (in metres)
' Comments: (1) Legacy code, no longer in ribbon
'==============================================================================
Private Function AirAbsorb(freq As String, Distance As Integer)
    Select Case freq
    Case Is = "63"
    AirAbsorb = -0.1 * (Distance / 1000)
    Case Is = "125"
    AirAbsorb = -0.3 * (Distance / 1000)
    Case Is = "250"
    AirAbsorb = -1.1 * (Distance / 1000)
    Case Is = "500"
    AirAbsorb = -2.8 * (Distance / 1000)
    Case Is = "1k"
    AirAbsorb = -5# * (Distance / 1000)
    Case Is = "2k"
    AirAbsorb = -9# * (Distance / 1000)
    Case Is = "4k"
    AirAbsorb = -22.9 * (Distance / 1000)
    Case Is = "8k"
    AirAbsorb = -76.6 * (Distance / 1000)
    End Select
End Function

'==============================================================================
' Name:     RoomAlphaDefault
' Author:   PS
' Desc:     Returns absorption values for different room types
' Args:     roomType - String of different room types
' Comments: (1)
'==============================================================================
Function RoomAlphaDefault(roomType As String)
    Select Case roomType
    Case Is = "Live"
    'bands                   31.5  63    125   250   500  1k   2k   4k   8k
    RoomAlphaDefault = Array(0.2, 0.18, 0.14, 0.11, 0.1, 0.1, 0.1, 0.1, 0.1)
    Case Is = "Av. Live"
    'bands                   31.5   63    125   250   500   1k    2k    4k    8k
    RoomAlphaDefault = Array(0.19, 0.18, 0.17, 0.14, 0.15, 0.15, 0.14, 0.13, 0.12)
    Case Is = "Average"
    'bands                   31.5   63    125   250   500   1k   2k    4k   8k
    RoomAlphaDefault = Array(0.2, 0.18, 0.19, 0.19, 0.2, 0.23, 0.22, 0.21, 0.2)
    Case Is = "Av. Dead"
    'bands                   31.5   63    125   250   500   1k   2k    4k    8k
    RoomAlphaDefault = Array(0.21, 0.2, 0.23, 0.24, 0.25, 0.28, 0.27, 0.26, 0.25)
    Case Is = "Dead"
    'bands                   31.5   63    125   250   500  1k  2k    4k   8k
    RoomAlphaDefault = Array(0.22, 0.2, 0.28, 0.3, 0.4, 0.47, 0.45, 0.44, 0.45)
    Case Is = ""
    RoomAlphaDefault = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    End Select
End Function

'==============================================================================
' Name:     RoomAlphaRTcurves
' Author:   PS
' Desc:     Returns absorption values for different RT ranges (midfrequency)
' Args:     RT_Type - string of different ranges of RT
' Comments: (1)
'==============================================================================
Function RoomAlphaRTcurves(RT_Type As String)
    'Alpha values are based on getting the desired midfrequency reverberation time
    Select Case RT_Type
    Case Is = "<0.2 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.21, 0.277, 0.331, 0.385, 0.435, 0.446, 0)
    Case Is = "0.2 to 0.5 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.125, 0.138, 0.183, 0.233, 0.288, 0.296, 0)
    Case Is = "0.5 to 1 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.109, 0.112, 0.137, 0.18, 0.214, 0.225, 0)
    Case Is = "1 to 1.5 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.057, 0.056, 0.058, 0.069, 0.08, 0.082, 0)
    Case Is = "1.5 to 2 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.053, 0.053, 0.06, 0.08, 0.095, 0.1, 0)
    Case Is = ">2 sec"
    'bands                  31.5 63 125   250    500     1k     2k     4k   8k
    RoomAlphaRTcurves = Array(0, 0, 0.063, 0.052, 0.036, 0.041, 0.035, 0.04, 0)
    Case Is = ""
    RoomAlphaRTcurves = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    End Select
End Function

'==============================================================================
' Name:     RoomLossTypical
' Author:   PS
' Desc:     Returns the SWL to SPL conversion in octave bands, given input
'           dimensions and roomType descriptor
' Args:     fstr - octave band centre frequency
'           L/W/H - room dimensions in metres
'           rooomType - description string of room reverberance
' Comments: (1) Generalised for alpha values from Potorff AIM, these seem to be
'           ok for office type spaces, but it would be great to verify where
'           they apply and where they don't. <--TODO: this
'==============================================================================
Function RoomLossTypical(fStr As String, L As Double, W As Double, H As Double, _
roomType As String)

Dim alpha() As Variant
Dim alpha_av As Double
Dim Rc As Double
Dim bandIndex As Integer

    alpha = RoomAlphaDefault(roomType)
    bandIndex = GetArrayIndex_OCT(fStr, 1)

    If bandIndex = 999 Or bandIndex = -1 Then
    RoomLossTypical = "-" 'no band, no result!
    Else
    
    S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)
    alpha_av = ((L * W * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) _
        + (W * H * alpha(bandIndex) * 2)) / S_total
    Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)
    
    'Debug.Print "Room Contant " Rc
        If Rc <> 0 Then
        RoomLossTypical = 10 * Application.WorksheetFunction.Log10(4 / Rc)
        Else
        RoomLossTypical = 0
        End If
        
    End If
    
End Function

'==============================================================================
' Name:     RoomLossTypicalRT
' Author:   PS
' Desc:     Returns room loss based
' Args:     fStr - octave bandd centre frequency, Hz
'           L/W/H - room dimensions in metres
'           RT_Type - reverberance length as text descriptor (set in form)
' Comments: (1)
'==============================================================================
Function RoomLossTypicalRT(fStr As String, L As Double, W As Double, H As Double, _
RT_Type As String)

Dim alpha() As Variant
Dim alpha_av As Double
Dim Rc As Double
Dim bandIndex As Integer

alpha = RoomAlphaRTcurves(RT_Type)

bandIndex = GetArrayIndex_OCT(fStr, 1)

S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)

alpha_av = ((L * W * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) + _
    (W * H * alpha(bandIndex) * 2)) / S_total
    
Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)

    If Rc <> 0 Then
    RoomLossTypicalRT = 10 * Application.WorksheetFunction.Log10(4 / Rc)
    Else
    RoomLossTypicalRT = 0
    End If

End Function

'==============================================================================
' Name:     ParallelipipedSurfaceArea
' Author:   PS
' Desc:     Parallel box method, integrated into area correction
' Args:
' Comments: (1)
'==============================================================================
Function ParallelipipedSurfaceArea(L As Double, W As Double, H As Double, _
    Offset As Double)
    
ParallelipipedSurfaceArea = ((L + (Offset * 2)) * (W + (Offset * 2)) + _
    (W + (Offset * 2)) * (H + (Offset * 2)) + _
    (L + (Offset * 2)) * (H + (Offset * 2))) * 2
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     DistancePoint
' Author:   PS
' Desc:     Inserts distance attenuation formula (spherical spreading)
' Args:     None
' Comments: (1)
'==============================================================================
Sub DistancePoint()

SetDescription "Distance Attenuation - point"

BuildFormula "10*LOG(" & T_ParamRng(1) & _
    "/(4*PI()*" & T_ParamRng(0) & "^2))"

SetTraceStyle "Input", True

ParameterUnmerge (Selection.Row)

'formatting
Cells(Selection.Row, T_ParamStart) = 10 'default to 10 metres
Cells(Selection.Row, T_ParamStart + 1) = 2 'default to half spherical
SetUnits "m", T_ParamStart, 1
SetUnits "Q", T_ParamStart + 1
SetDataValidation T_ParamStart + 1, "1,2,4,8"
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     DistanceLine
' Author:   PS
' Desc:     Inserts distance attenuation formula (cylindrical spreading)
' Args:     None
' Comments: (1)
'==============================================================================
Sub DistanceLine()

SetDescription "Distance Attenuation - line"

BuildFormula "10*LOG(" & T_ParamRng(1) & _
    "/(2*PI()*" & T_ParamRng(0) & "))"

SetTraceStyle "Input", True

ParameterUnmerge (Selection.Row)

'formatting
Cells(Selection.Row, T_ParamStart) = 10 'default to 10 metres
Cells(Selection.Row, T_ParamStart + 1) = 2 'default to half cylindrical
SetUnits "m", T_ParamStart, 1
SetUnits "Q", T_ParamStart + 1
SetDataValidation T_ParamStart + 1, "1,2,4,8"
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     DistancePlane
' Author:   PS
' Desc:     Inserts distance loss formula for a plane wave into free space
' Args:     None
' Comments: (1) From Biess and Hansen:
'               In the near field (approximately `r < a//pi`), the sound level
'               can be approximated as: `L_p=L_W-10log_10 S+DI`
'               In the line-source intermediate region
'               (approximately `a//pi < r < b//pi`), the sound level can be
'               approximated as: `L_p=L_W-10log_10 S-10log_10(d/(a//pi))+DI`
'               In the point-source far region (approximately `r > b//pi`),
'               the sound level can be approximated as:
'               `L_p=L_W-10log_10 S-10log_10(a/b)-20log(d/(b//pi))+DI`
'               where `r` is the distance from the source, `H` and `L` are the
'               minor and major source dimensions (m), `S=H*L` is the area of
'               the source (m?) and `DI` is the directivity index of the
'               source (dB).
'==============================================================================
Sub DistancePlane()

frmPlaneSource.Show

If btnOkPressed = False Then End 'catch cancel

SetDescription "Distance Attenuation - plane"

BuildFormula "-10*LOG(" & PlaneH & "*" & _
    PlaneL & ")+10*LOG(ATAN((" & PlaneH & "*" & PlaneL & ")/(2*" & T_ParamRng(0) & _
    "*SQRT((" & PlaneH & "^2)+(" & PlaneL & "^2)+(4*" & T_ParamRng(0) & "^2)))))-2"

SetTraceStyle "Input", True

ParameterMerge (Selection.Row)

Cells(Selection.Row, T_ParamStart) = PlaneDist
SetUnits "m", T_ParamStart, 1

End Sub

'==============================================================================
' Name:     Distance Ratio Point
' Author:   PS
' Desc:     Inserts ratio of the distances (point sources)
' Args:     None
' Comments: (1)
'==============================================================================
Sub DistanceRatioPoint()
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

SetDescription "Distance Attenuation - ratio (point)"

BuildFormula "20*LOG(" & T_ParamRng(0) & "/" & _
    T_ParamRng(1) & ")"

SetTraceStyle "Input", True

Cells(Selection.Row, T_ParamStart).Value = 1
Cells(Selection.Row, T_ParamStart + 1).Value = 2
SetUnits "m", T_ParamStart, 0, T_ParamStart + 1
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     Distance Ratio Line
' Author:   PS
' Desc:     Inserts ratio of the distances (line sources)
' Args:     None
' Comments: (1)
'==============================================================================
Sub DistanceRatioLine()
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

SetDescription "Distance Attenuation - ratio (line)"

BuildFormula "10*LOG(" & T_ParamRng(0) & "/" & _
    T_ParamRng(1) & ")"

'set defaults and apply formats
SetTraceStyle "Input", True
Cells(Selection.Row, T_ParamStart).Value = 1
Cells(Selection.Row, T_ParamStart + 1).Value = 2
SetUnits "m", T_ParamStart, 0, T_ParamStart + 1
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     AreaCorrection
' Author:   PS
' Desc:     Creates a 10log(area) formula
' Args:     None
' Comments: (1) Simple, but gooooood
'==============================================================================
Sub AreaCorrection()
'description
SetDescription "Area Correction: 10log(A)"
'set parameter
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 2
SetUnits "m2", T_ParamStart
'build formula
BuildFormula "10*LOG(" & T_ParamRng(0) & ")"
'formatting
SetTraceStyle "Input", True
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     ParallelipipedCorrection
' Author:   PS
' Desc:     Calculated area of perpendicular box and correction from SPL to SWL
' Args:     None
' Comments: (1) What's in the box?????????
'==============================================================================
Sub ParallelipipedCorrection()
frmSoundPowerCalculator.Show
    If btnOkPressed = False Then End
    
SetDescription "Parellelipiped Correction"
BuildFormula "10*LOG(" & T_ParamRng(0) & ")"
    
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = "=ParallelipipedSurfaceArea(" & _
    roomL & "," & roomW & "," & roomH & "," & OffsetDistance & ")"
SetUnits "m2", T_ParamStart
End Sub

'==============================================================================
' Name:     TenLogN
' Author:   PS
' Desc:     Creates a 10log(n) formula
' Args:     None
' Comments: (1) Kinda the same as area, but with different formatting
'==============================================================================
Sub TenLogN()
SetDescription "Multiple sources: 10log(n)"

'set inputs
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 2
Cells(Selection.Row, T_ParamStart).NumberFormat = """n = ""0"
SetTraceStyle "Input", True

BuildFormula "10*LOG(" & _
    T_ParamRng(0) & ")"
End Sub

'==============================================================================
' Name:     TenLogOneOnT
' Author:   PS
' Desc:     Creates a 10log(1/T) formula
' Args:     None
' Comments: (1) Kinda the same as 10log(n), but as a ratio
'==============================================================================
Sub TenLogOneOnT()

SetDescription "Time Correction: 10log(t/t0)"

BuildFormula "10*LOG(" & _
    T_ParamRng(0) & "/" & _
    T_ParamRng(1) & ")"
ParameterUnmerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 1
Cells(Selection.Row, T_ParamStart).NumberFormat = """t = ""0"
Cells(Selection.Row, T_ParamStart + 1) = 2
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """t0 = ""0"

SetTraceStyle "Input", True
End Sub

'==============================================================================
' Name:     PutRoomLossTypical
' Author:   PS
' Desc:     Applies correction to account for difference between sound power and
'           sound pressure in a room
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRoomLossTypical()
Dim SplitStr() As String

    'if the user has already put it in, populate form with values previously entered!
    '<---TODO: make this a universal call for all 3 room loss forms
    If InStr(1, Cells(Selection.Row, T_LossGainStart).Formula, "RoomLossTypical", _
        vbTextCompare) > 1 Then
    SplitStr = Split(Cells(Selection.Row, T_LossGainStart).Formula, ",", _
        Len(Cells(Selection.Row, T_LossGainStart).Formula), vbTextCompare)
    roomL = CLng(SplitStr(1))
    roomW = CLng(SplitStr(2))
    roomH = CLng(SplitStr(3))
    roomType = Cells(Selection.Row, T_ParamStart).Value
    Call frmRoomLossClassic.PrePopulateForm
    End If

frmRoomLossClassic.Show

    If btnOkPressed = False Then End
    
SetDescription "Room Loss"

BuildFormula "RoomLossTypical(" & T_FreqStartRng & _
    "," & roomL & "," & roomW & "," & roomH & "," & T_ParamRng(0) & ")"

ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = roomType
SetTraceStyle "Input", True
SetDataValidation T_ParamStart, "Dead, Av. Dead, Average, Av. Live, Live"

End Sub

'==============================================================================
' Name:     PutRoomLossRC
' Author:   PS
' Desc:     Inserts formula for room loss based on a room constant
' Args:     None
' Comments: (1) Calculate RC in the Reverberation Time Calc Sheet
'           (2) Requires two rows of space
'==============================================================================
Sub PutRoomLossRC()
Dim DefaultArray() As Variant

'Nominal Rc, based on a 0.5sec RT, or something
DefaultArray = Array(17, 19, 22, 24, 31, 39, 43)

SetDescription "Room Constant"
'set default array
Range(Cells(Selection.Row, T_LossGainStart + 1), _
      Cells(Selection.Row, T_LossGainStart + 1 + UBound(DefaultArray))).Value _
      = DefaultArray

SetTraceStyle "Input"

'move one row down
Cells(Selection.Row + 1, Selection.Column).Select

'build RC formula
SetDescription "Room Loss - 10LOG(4/Rc)"
BuildFormula "10*LOG(4/" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & ")"


'delete first and last octave bands
Cells(Selection.Row, FindFrequencyBand("31.5")).ClearContents
Cells(Selection.Row, FindFrequencyBand("8k")).ClearContents

End Sub


'==============================================================================
' Name:     PutRoomLossTypicalRT
' Author:   PS
' Desc:     Returns a room loss, based on a reverb time and some BROAD
'           assumptions. Not super robust but ok.
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRoomLossTypicalRT()
Dim SplitStr() As String
Dim ParamCol As Integer

    'if the user has already put it in, populate form with values previously entered!
    '<---TODO: make this a universal call for all 3 room loss forms
    If InStr(1, Cells(Selection.Row, T_LossGainStart).Formula, _
        "RoomLossTypicalRT", vbTextCompare) > 1 Then
    'populate frmroomlossRT
    SplitStr = Split(Cells(Selection.Row, T_LossGainStart).Formula, ",", _
        Len(Cells(Selection.Row, T_LossGainStart).Formula), vbTextCompare)
    roomL = CLng(SplitStr(1))
    roomW = CLng(SplitStr(2))
    roomH = CLng(SplitStr(3))
    roomType = Cells(Selection.Row, T_ParamStart).Value
    Call frmRoomLossRT.PrePopulateForm
    End If

frmRoomLossRT.Show

    If btnOkPressed = False Then End

SetDescription "Room Loss - RT Typical"

BuildFormula "RoomLossTypicalRT(" & T_FreqStartRng & _
    "," & roomL & "," & roomW & "," & roomH & "," & T_ParamRng(0) & ")"


ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = roomType
SetTraceStyle "Input", True
SetDataValidation T_ParamStart, _
    "<0.2 sec,0.2 to 0.5 sec,0.5 to 1 sec,1.5 to 2 sec,>2 sec"

End Sub


'==============================================================================
' Name:     PutRoomLossRT
' Author:   PS
' Desc:     Returns a room loss, based on a reverb time from the row above
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutRoomLossRT()
Dim DefaultArray() As Variant 'some default values
Dim LineAbove As String 'variable for address of row above

'Nominal RT, based on measurement of some room
DefaultArray = Array(0.7, 0.8, 0.6, 0.6, 0.6, 0.5, 0.5, 0.4)

SetDescription "Reverberation Time"

'set default array
Range(Cells(Selection.Row, T_LossGainStart + 1), _
      Cells(Selection.Row, T_LossGainStart + 1 + UBound(DefaultArray))).Value _
      = DefaultArray
SetTraceStyle "Input"
Cells(Selection.Row, T_ParamStart).NumberFormat = "0""m" & chr(179) & """"
Range(Cells(Selection.Row, T_LossGainStart), _
    Cells(Selection.Row, T_LossGainEnd)).NumberFormat = "0.0"

'volume input
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = "=3*4*2.7"
SetTraceStyle "Input", True
Cells(Selection.Row, T_ParamStart).NumberFormat = "0""m" & chr(179) & """"

'move one row down
Cells(Selection.Row + 1, Selection.Column).Select
GetSettings

SetDescription "Room Loss - RT"
LineAbove = Cells(Selection.Row - 1, T_LossGainStart).Address(False, False)
BuildFormula "10*log(" & LineAbove & "/(0.163*" & T_ParamRng(0) & "))"

End Sub

'==============================================================================
' Name:     DirectReverberantSum
' Author:   PS
' Desc:     AutoSums rows, inserts spherical spreading, calculates the direct
'           path, adds Room Loss, calculates the reverberant path, log-sums the
'           two paths, applies formatting
' Args:     None
' Comments: (1) This macro is cool.
'==============================================================================
Sub DirectReverberantSum()
Dim SpareRow As Integer
Dim SpareCol As Integer
Dim isSpace As Boolean
Dim StartRw As Integer
Dim EndRw As Integer
Dim ScanCol As Integer
Dim TopOfSheet As Boolean

'code requires 6 free rows
isSpace = True
    For SpareRow = Selection.Row To Selection.Row + 5
        For SpareCol = T_LossGainStart To T_LossGainEnd
            If Cells(SpareRow, SpareCol).Value <> "" Then
            isSpace = False
            End If
        Next SpareCol
    Next SpareRow
    
    If isSpace = False Then
    msg = MsgBox("Not enough space. Do you wish to overwrite?", _
        vbYesNo, "SQUISH!")
        
        If msg = vbYes Then
        Range(Cells(Selection.Row, Selection.Column), _
            Cells(Selection.Row + 5, Selection.Column)).Select
        ClearRow (True) 'skips user input
        Else
        End
        End If
        
    End If

'find sum range
StartRw = Selection.Row - 1 'one above currently seelcted
ScanCol = Selection.Column
    While Cells(StartRw, ScanCol).Value <> "" 'looks for blank cell
    StartRw = StartRw - 1
        If StartRw < 7 Then
        TopOfSheet = True
        'msg = MsgBox("AutoSum Error", vbOKOnly, "ERROR")
        'End
        End If
    Wend
StartRw = StartRw + 1
'check if selection is in the forbidden zone
If TopOfSheet = True Then StartRw = 7

EndRw = Selection.Row - 1 'for reveberant sum

'<---------------------------------------TODO: Show form and let the user see the range to be summed

'distance correction
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 1  '1m by default
'move down
SelectNextRow
SetSheetTypeControls
'Sum direct
BuildFormula "SUM(" & _
    Cells(StartRw, T_LossGainStart).Address(False, False) & ":" & _
    Cells(EndRw + 1, T_LossGainStart).Address(False, False) & ")"
    

SetDescription "Direct component"
SetTraceStyle "Subtotal"

'move cursor
SelectNextRow

'Room loss
PutRoomLossTypical

'move down
SelectNextRow

'number of reverberant sources
TenLogN
Cells(Selection.Row, T_ParamStart).Value = 1 'default to 1 source ie 10log(n)=0

'move down
SelectNextRow

'Sum reverb
BuildFormula "SUM(" & _
    Cells(StartRw, T_LossGainStart).Address(False, False) & ":" & _
    Cells(EndRw, T_LossGainStart).Address(False, False) & "," & _
    Cells(Selection.Row - 2, T_LossGainStart).Address(False, False) & ":" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & ")"

SetDescription "Reverberant component"
SetTraceStyle "Subtotal"

'move down
SelectNextRow

'Sum Total
BuildFormula "SPLSUM(" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "," & _
    Cells(Selection.Row - 4, T_LossGainStart).Address(False, False) & ")"

SetDescription "Total"
SetTraceStyle "Total"

End Sub
