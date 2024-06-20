Attribute VB_Name = "Noise"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================

'plane waves
Public PlaneH As Double
Public PlaneW As Double
Public PlaneDist As Double

'rooms
Public RoomType As String
Public roomL As Double
Public roomW As Double
Public roomH As Double
Public roomLossType As String
Public OffsetDistance As Double

'barriers
Public Barrier_Method As String
Public Barrier_SourceToBarrier As Double
Public Barrier_SourceHeight As Double
Public Barrier_GroundUnderSrc As Double
Public Barrier_RecToBarrier As Double
Public Barrier_ReceiverHeight As Double
Public Barrier_GroundUnderRec As Double
Public Barrier_BarrierHeight As Double
Public Barrier_SpreadingType As String
Public Barrier_SrcToBarrierEdge As Double
Public Barrier_RecToBarrierEdge As Double
Public Barrier_BarrierHeightReceiverSide As Double
Public Barrier_DoubleDiffraction As Double
Public Barrier_BarrierThickness As Double
Public Barrier_MultiSource As Double
Public Barrier_SrcRecDistance As Double
Public Barrier_GtoRecheight As Double
Public Barrier_GtoSrcHeight As Double

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
Function RoomAlphaDefault(RoomType As String)
    Select Case RoomType
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
' Args:     fStr - octave band centre frequency
'           L/W/H - room dimensions in metres
'           rooomType - description string of room reverberance
' Comments: (1) Generalised for alpha values from Potorff AIM, these seem to be
'           ok for office type spaces, but it would be great to verify where
'           they apply and where they don't. <--TODO: this
'==============================================================================
Function RoomLossTypical(fstr As String, L As Double, W As Double, H As Double, _
RoomType As String)

Dim alpha() As Variant
Dim alpha_av As Double
Dim Rc As Double
Dim BandIndex As Integer

    alpha = RoomAlphaDefault(RoomType)
    BandIndex = GetArrayIndex_OCT(fstr, 1)

    If BandIndex = 999 Or BandIndex = -1 Then
    RoomLossTypical = "-" 'no band, no result!
    Else
    
    S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)
    alpha_av = ((L * W * alpha(BandIndex) * 2) + (L * H * alpha(BandIndex) * 2) _
        + (W * H * alpha(BandIndex) * 2)) / S_total
    Rc = (S_total * alpha(BandIndex)) / (1 - alpha_av)
    
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
Function RoomLossTypicalRT(fstr As String, L As Double, W As Double, H As Double, _
RT_Type As String)

Dim alpha() As Variant
Dim alpha_av As Double
Dim Rc As Double
Dim BandIndex As Integer

alpha = RoomAlphaRTcurves(RT_Type)

BandIndex = GetArrayIndex_OCT(fstr, 1)

S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)

alpha_av = ((L * W * alpha(BandIndex) * 2) + (L * H * alpha(BandIndex) * 2) + _
    (W * H * alpha(BandIndex) * 2)) / S_total
    
Rc = (S_total * alpha(BandIndex)) / (1 - alpha_av)

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
' Args:     L - length in metres
'           W - width in metres
'           H - height in metres
'           Offset - distance to the surface of the object, in metres
' Comments: (1)
'==============================================================================
Function ParallelipipedSurfaceArea(L As Double, W As Double, H As Double, _
    Offset As Double)
Dim A As Double
Dim B As Double
Dim C As Double

A = (0.5 * L) + Offset
B = (0.5 * W) + Offset
C = H + Offset

'Area is:
    '2 x sides front/back
    '2 x sides left/right
    '1 x top (ieno bottom)
    
ParallelipipedSurfaceArea = 4 * ((A * B) + (B * C) + (C * A))
    
End Function

'==============================================================================
' Name:     ConformalSurfaceArea
' Author:   PS
' Desc:     Conformal area method, integrated into area correction
' Args:     L - length in metres
'           W - width in metres
'           H - height in metres
'           Offset - distance to the surface of the object, in metres
' Comments: (1)
'==============================================================================
Function ConformalSurfaceArea(L As Double, W As Double, H As Double, _
    Offset As Double)
    
ConformalSurfaceArea = (L + W) * ((2 * H) + (Application.WorksheetFunction.Pi * Offset)) _
    + ((2 * Application.WorksheetFunction.Pi * Offset) * (H + Offset)) + (L * W)
    
End Function

'==============================================================================
' Name:     DistancePlaneSource
' Author:   PS
' Desc:     Parallel box method, integrated into area correction
' Args:     W: width of the plane source in metres
'           H: height of the plane source in metres
'           distance: Distance to the plane source, in metres
' Comments: (1)
'==============================================================================
Function DistancePlaneSource(H As Double, W As Double, Distance As Double)
    
DistancePlaneSource = -10 * Application.WorksheetFunction.Log(W * H) + _
    10 * Application.WorksheetFunction.Log( _
    Atn((W * H) / (2 * Distance * ((W ^ 2) + (H ^ 2) + (4 * Distance ^ 2)) ^ (1 / 2)))) - 2
    
End Function


'==============================================================================
' Name:     BarrierAtten_KurzeAnderson
' Author:   PS & CT
' Desc:
' Args:     fStr - frequency band, as string
'           SrcDistToBarrier - distance from source to barrier
'           SrcHeight - height of source above floor
'           GroundUnderSrc - height of floor above ground
'           RecDistToBarrier- distance from receiver to barrier
'           RecHeight - height of receiver above floor
'           GroundUnderRec - height of floor above ground
'           BarrierHeight - height of barrier above ground
'           IncludeMultiPAth - default to false, don't calc them!
' Comments: (1)
'==============================================================================
Function BarrierAtten_KurzeAnderson(fstr As String, SrcDistToBarrier As Double, SrcHeight As Double, _
    GroundUnderSrc As Double, RecDistToBarrier As Double, RecHeight As Double, _
    GroundUnderRec As Double, BarrierHeight As Double, Optional IncludeMultiPath As Boolean)

'paths
Dim p1 As Double
Dim p2 As Double
Dim p3 As Double
Dim p4 As Double
'distances
Dim d0 As Double
Dim d1 As Double
Dim d2 As Double
Dim d3 As Double
Dim d4 As Double
Dim d5 As Double
Dim d6 As Double
Dim d7 As Double
Dim d8 As Double
Dim d9 As Double
'Level from each path
Dim L1 As Double
Dim L2 As Double
Dim L3 As Double
Dim L4 As Double

Dim DirectSPL As Double
Dim BarrierSPL As Double

    'check for line of sight
    If BarrierCutsLineofSight(SrcDistToBarrier, SrcHeight, GroundUnderSrc, _
        RecDistToBarrier, RecHeight, GroundUnderRec, BarrierHeight) = False Then
    BarrierAtten_KurzeAnderson = "-"
    Exit Function
    End If

'calculate path lengths
d0 = ((SrcDistToBarrier + RecDistToBarrier) ^ 2 + _
    ((SrcHeight + GroundUnderSrc) - (RecHeight + GroundUnderRec)) ^ 2) ^ 0.5
d1 = (SrcDistToBarrier ^ 2 + _
    (BarrierHeight - (SrcHeight + GroundUnderSrc)) ^ 2) ^ 0.5
d2 = (RecDistToBarrier ^ 2 + (BarrierHeight - (RecHeight + GroundUnderRec)) ^ 2) ^ 0.5
'd3 = ((SrcDistToBarrier / 2) ^ 2 + (SrcHeight ^ 2)) ^ 0.5
d4 = ((SrcDistToBarrier / 2) ^ 2 + (BarrierHeight - GroundUnderSrc) ^ 2) ^ 0.5
d5 = ((RecDistToBarrier / 2) ^ 2 + (BarrierHeight - GroundUnderRec) ^ 2) ^ 0.5
'd6 = ((RecDistToBarrier / 2) ^ 2 + RecHeight ^ 2) ^ 0.5
d7 = (((SrcDistToBarrier / 2) + (RecDistToBarrier)) ^ 2 + _
    ((RecHeight + GroundUnderRec) - (GroundUnderSrc)) ^ 2) ^ 0.5
d8 = (((SrcDistToBarrier) + (RecDistToBarrier / 2)) ^ 2 + _
    ((SrcHeight + GroundUnderSrc) - (GroundUnderRec)) ^ 2) ^ 0.5
d9 = ((SrcDistToBarrier / 2 + RecDistToBarrier / 2) ^ 2 + _
    (Abs(GroundUnderSrc - GroundUnderRec)) ^ 2) ^ 0.5


'path 1
p1 = BarrierAtten_KA_path(d2, d1, d0, fstr) 'long + long - short

    'option for multipath
    If IncludeMultiPath = True Then
    'Path 2
    p2 = BarrierAtten_KA_path(d4, d2, d7, fstr)
    'path 3
    p3 = BarrierAtten_KA_path(d1, d5, d8, fstr)
    'path 4
    p4 = BarrierAtten_KA_path(d4, d5, d9, fstr)
    
    DirectSPL = 100 - 10 * Application.WorksheetFunction.Log(4 * _
        Application.WorksheetFunction.Pi() * (d0 ^ 2)) 'nominal, start from 100
    
    L1 = DirectSPL - p1
    L2 = DirectSPL - p2
    L3 = DirectSPL - p3
    L4 = DirectSPL - p4
    
    'Debug.Print L1; L2; L3; L4
    
    BarrierSPL = SPLSUM(L1, L2, L3, L4)
    
    BarrierAtten_KurzeAnderson = BarrierSPL - DirectSPL
    
    Else 'most of the time, default to the regular KA method
    BarrierAtten_KurzeAnderson = -1 * p1 '- DirectSPL
    End If

    'check for practical maximum value 30dB
    If BarrierAtten_KurzeAnderson < -30 Then
    BarrierAtten_KurzeAnderson = -30
    End If
    
End Function

'==============================================================================
' Name:     BarrierAtten_KA_path
' Author:   PS & CT
' Desc:     Calculates noise level over barrier from path length inputs
' Args:     d0 - short path length in metres
'           d1 - long path length pt1 (source to barrier top) in metres
'           d2 - long path length pt2 (receiver to barrier top) in metres
'           fStr - frequency as string
' Comments: (1)
'==============================================================================
Function BarrierAtten_KA_path(d1 As Double, d2 As Double, d0 As Double, fstr As String)

Dim SOS As Double
Dim f As Double
Dim Wavelength As Double
Dim FresnelNo As Double
Dim TwoPi As Double
f = freqStr2Num(fstr)

SOS = SpeedOfSound(20) 'm/s 'todo: add optional input for temperature??
Wavelength = SOS / f
TwoPi = 2 * Application.WorksheetFunction.Pi() 'save on characters later
'Debug.Print "Path difference; "; (d1 + d2 - d0)

FresnelNo = (2 / Wavelength) * (d1 + d2 - d0)

BarrierAtten_KA_path = 5 + 20 * Application.WorksheetFunction.Log( _
    ((TwoPi * FresnelNo) ^ 0.5) / _
    Application.WorksheetFunction.Tanh((TwoPi * FresnelNo) ^ 0.5))
End Function


'==============================================================================
' Name:     BarrierAtten_Menounou
' Author:   PS
' Desc:     Menounou's method for barrier insertion loss
' Args:     fStr - frequency band, as string
'           SrcDistToBarrier - distance from source to barrier
'           SrcHeight - height of source above floor
'           GroundUnderSrc - height of floor above ground
'           RecDistToBarrier- distance from receiver to barrier
'           RecHeight - height of receiver above floor
'           GroundUnderRec - height of floor above ground
'           BarrierHeight - height of barrier above ground
'           SpreadingType - string with either sphere cylinder or plane
' Comments: (1) All distances are in metres
'==============================================================================
Function BarrierAtten_Menounou(fstr As String, SrcDistToBarrier As Double, _
    SrcHeight As Double, GroundUnderSrc As Double, RecDistToBarrier As Double, _
    RecHeight As Double, GroundUnderRec As Double, BarrierHeight As Double, _
    SpreadingType As String)

'variables
Dim d0 As Double 'straight-line distance from src to rec
Dim d1 As Double 'distance from source to top of barrier
Dim d2 As Double 'distance from barrier to receiver
Dim dx As Double 'distance of the mirror source to the receiver????
Dim f As Double 'frequency in Hz
Dim Theta As Double
Dim Wavelength As Double 'in metres
Dim FresnelNo As Double 'of path length difference
Dim TwoPi As Double
Dim N2 As Double 'Fresnel number of mirror source

'Variables from Menounou's method
Dim IL_s As Double
Dim IL_b As Double
Dim IL_sb As Double
Dim IL_sp As Double

    'check for line of sight
    If BarrierCutsLineofSight(SrcDistToBarrier, SrcHeight, GroundUnderSrc, _
        RecDistToBarrier, RecHeight, GroundUnderRec, BarrierHeight) = False Then
    BarrierAtten_Menounou = "-"
    Exit Function
    
    'check spreading type is a defined option
    ElseIf SpreadingType <> "Plane" And SpreadingType <> "Cylindrical" And _
        SpreadingType <> "Spherical" Then
    BarrierAtten_Menounou = "-"
    Exit Function
    End If

'convert to value
f = freqStr2Num(fstr)


'calc distances with pythagoras
d0 = ((SrcDistToBarrier + RecDistToBarrier) ^ 2 + _
    ((SrcHeight + GroundUnderSrc) - (RecHeight + GroundUnderRec)) ^ 2) ^ 0.5
d1 = (SrcDistToBarrier ^ 2 + (BarrierHeight - (SrcHeight + GroundUnderSrc)) ^ 2) ^ 0.5
d2 = (RecDistToBarrier ^ 2 + (BarrierHeight - (RecHeight + GroundUnderRec)) ^ 2) ^ 0.5


SOS = SpeedOfSound(20) 'm/s 'todo: add optional input for temperature??
Wavelength = SOS / f
TwoPi = 2 * Application.WorksheetFunction.Pi() 'save on characters later
FresnelNo = (2 / Wavelength) * (d1 + d2 - d0)
Theta = 2 * Application.WorksheetFunction.Acos(SrcDistToBarrier / d1)

dx = ((d1 ^ 2) + (d2 ^ 2) + (2 * d1 * d2 * Cos(Theta))) ^ 0.5
N2 = (2 / Wavelength) * dx 'the other Fresnel's number, for dx

'let's calculate each component of the method!

IL_s = 20 * Application.WorksheetFunction.Log( _
    ((TwoPi * FresnelNo) ^ 0.5) / _
    Application.WorksheetFunction.Tanh((TwoPi * FresnelNo) ^ 0.5)) - 1
    
IL_b = 20 * Application.WorksheetFunction.Log(1 + _
    Application.WorksheetFunction.Tanh( _
    0.6 * Application.WorksheetFunction.Log(N2 / FresnelNo)))
    
IL_sb = (Application.WorksheetFunction.Tanh(N2 ^ 0.5 - 2 - IL_b)) * _
    (1 - Application.WorksheetFunction.Tanh((10 * FresnelNo) ^ 0.5))

    Select Case SpreadingType
    Case Is = "Spherical"
    IL_sp = 10 * Application.WorksheetFunction.Log(((d1 + d2) ^ 2 / d0 ^ 2) + _
        ((d1 + d2) / d0))
    Case Is = "Cylindrical"
    IL_sp = 10 * Application.WorksheetFunction.Log(1 + ((d1 + d2) / d0))
    Case Is = "Plane"
    IL_sp = 3
    End Select

'Debug.Print IL_s; IL_b; IL_sb; IL_sp

'add them up and make it negative!
BarrierAtten_Menounou = -1 * (IL_s + IL_b + IL_sb + IL_sp)
'what are the last two terms? D_Theta_R & D_Theta_B

End Function

'==============================================================================
' Name:     BarrierCutsLineofSight
' Author:   CT
' Desc:     Returns TRUE if barrier cuts line of sight (source to receiver line)
' Args:     None
' Comments: (1)
'==============================================================================
Function BarrierCutsLineofSight(SourceToBarrier As Double, SrcHeight As Double, _
    SrcGroundHeight As Double, RecDistToBarrier As Double, RecHeight As Double, _
    GroundUnderRec As Double, BarrierHeight As Double) As Boolean

Dim SlopeSrcRec As Double 'Slope of source and receiver
Dim SlopeSrcBar As Double 'Slope of source to the top of the barrier

SlopeSrcRec = (RecHeight + GroundUnderRec - SrcHeight - SrcGroundHeight) / _
    (SourceToBarrier + RecDistToBarrier)
    
    If SourceToBarrier <= 0 Then Exit Function

SlopeSrcBar = (BarrierHeight - SrcHeight - SrcGroundHeight) / (SourceToBarrier)
    
    If SlopeSrcBar > SlopeSrcRec Then
    BarrierCutsLineofSight = True
    Else
    BarrierCutsLineofSight = False
    End If

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

ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = PlaneDist

BuildFormula "DistancePlaneSource(" & PlaneH & "," & PlaneW & "," & _
    T_ParamRng(0) & ")"
'old version of build
'BuildFormula "-10*LOG(" & PlaneH & "*" & _
'    PlaneW & ")+10*LOG(ATAN((" & PlaneH & "*" & PlaneW & ")/(2*" & T_ParamRng(0) & _
'    "*SQRT((" & PlaneH & "^2)+(" & PlaneW & "^2)+(4*" & T_ParamRng(0) & "^2)))))-2"
InsertComment "Plane source: " & PlaneH & "m x " & PlaneW & "m", T_ParamStart, False
SetTraceStyle "Input", True

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
frmSoundPowerCalculator.optParallel.Value = True
frmSoundPowerCalculator.Show
    If btnOkPressed = False Then End

'user may have changed it so check!
If frmSoundPowerCalculator.optParallel.Value = True Then
    SetDescription "Parellelipiped Correction"
Else
    SetDescription "Conformal Surface Area Correction"
End If

BuildFormula "10*LOG(" & T_ParamRng(0) & ")"
    
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = "=ParallelipipedSurfaceArea(" & _
    roomL & "," & roomW & "," & roomH & "," & OffsetDistance & ")"
SetUnits "m2", T_ParamStart
End Sub

'==============================================================================
' Name:     ParallelipipedCorrection
' Author:   PS
' Desc:     Calculated area of perpendicular box and correction from SPL to SWL
' Args:     None
' Comments: (1) What's in the box?????????
'==============================================================================
Sub ConformalAreaCorrection()
frmSoundPowerCalculator.optConformal.Value = True
frmSoundPowerCalculator.Show
    If btnOkPressed = False Then End
    
'user may have changed it so check!
If frmSoundPowerCalculator.optParallel.Value = True Then
    SetDescription "Parellelipiped Correction"
Else
    SetDescription "Conformal Surface Area Correction"
End If

BuildFormula "10*LOG(" & T_ParamRng(0) & ")"
    
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = "=ConformalSurfaceArea(" & _
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
' Name:     FanSpeedCorrection
' Author:   PS
' Desc:     Creates a 50log(RPM1/RPM2) formula
' Args:     None
' Comments: (1) Used for fans operating below their maximum duty point
'==============================================================================
Sub FanSpeedCorrection()

SetDescription "Fan speed correction: 50log(RPM1/RPM2)"

BuildFormula "50*LOG(" & _
    T_ParamRng(0) & "/" & _
    T_ParamRng(1) & ")"
ParameterUnmerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = 1500
Cells(Selection.Row, T_ParamStart).NumberFormat = "0""RPM"""
Cells(Selection.Row, T_ParamStart + 1) = 1500
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = "0""RPM"""

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
    RoomType = Cells(Selection.Row, T_ParamStart).Value
    Call frmRoomLossClassic.PrePopulateForm
    End If

frmRoomLossClassic.Show

    If btnOkPressed = False Then End
    
SetDescription "Room Loss"

BuildFormula "RoomLossTypical(" & T_FreqStartRng & _
    "," & roomL & "," & roomW & "," & roomH & "," & T_ParamRng(0) & ")"

ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = RoomType
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
    RoomType = Cells(Selection.Row, T_ParamStart).Value
    Call frmRoomLossRT.PrePopulateForm
    End If

frmRoomLossRT.Show

    If btnOkPressed = False Then End

SetDescription "Room Loss - RT Typical"

BuildFormula "RoomLossTypicalRT(" & T_FreqStartRng & _
    "," & roomL & "," & roomW & "," & roomH & "," & T_ParamRng(0) & ")"


ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = RoomType
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
Dim startRw, endRw  As Integer
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

startRw = FindTopOfBlock(Selection.Column) 'temporary
endRw = Selection.Row - 1 'for reveberant sum

'distance correction
DistancePoint
Cells(Selection.Row, T_ParamStart).Value = 1  '1m by default
'move down
SelectNextRow
SetSheetTypeControls

'show the user the range to be summed
AutoSum_UserInput
'Sum direct
BuildFormula "SUM(" & _
    Cells(T_FirstSelectedRow, T_LossGainStart).Address(False, False) & ":" & _
    Cells(T_LastSelectedRow, T_LossGainStart).Address(False, False) & ")"
SetDescription "Direct component", endRw + 1, True
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
    Cells(T_FirstSelectedRow, T_LossGainStart).Address(False, False) & ":" & _
    Cells(T_LastSelectedRow, T_LossGainStart).Address(False, False) & "," & _
    Cells(Selection.Row - 2, T_LossGainStart).Address(False, False) & ":" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & ")"

SetDescription "Reverberant component"
SetTraceStyle "Subtotal"
ApplyTraceMarker ("Sum")
'move down
SelectNextRow

'Sum Total
BuildFormula "SPLSUM(" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "," & _
    Cells(Selection.Row - 4, T_LossGainStart).Address(False, False) & ")"

'description based on the first row above the start of the summed range
If Cells(T_FirstSelectedRow - 1, T_Description).Value <> "" Then
    SetDescription "=concat(""Total - ""," & Cells(T_FirstSelectedRow - 1, T_Description).Address(False, False) & ")"
Else
    SetDescription "Total"
End If

SetTraceStyle "Total"
ApplyTraceMarker ("Result")

End Sub

'==============================================================================
' Name:     BarrierAtten
' Author:   PS
' Desc:     Calls form and inserts barrier attenuation
' Args:     None
' Comments: (1) Fixed on 20220920 as 'hot patch' after rollout
'           (2) Changed to print method in the description, not just a comment
'==============================================================================
Sub BarrierAtten()

frmBarrierAtten.Show

    If btnOkPressed = False Then End

'description
SetDescription "Barrier Attenuation - " & Barrier_Method
'InsertComment Barrier_Method, T_Description, False

'parameter
ParameterMerge (Selection.Row)

SetTraceStyle "Input", True
SetUnits "m", T_ParamStart, 1
InsertComment "Barrier height", T_ParamStart, False

    'formula
    If Barrier_Method = "ISO9613_Abar" Then
    Cells(Selection.Row, T_ParamStart) = iso9613_BarrierHeight
    '@ISO9613_Abar(fStr,SourceHeight,ReceiverHeight,SourceReceiverDistance,SourceBarrierDistance,SrcDistanceEdge,RecDistanceEdge,HeightBarrierSource,DoubleDiffraction,BarrierThickness,HeightBarrierReceiver,multisource,GroundEffect)
    BuildFormula "ISO9613_Abar(" & T_FreqStartRng & "," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & iso9613_d & "," & iso9613_SourceToBarrier & "," & _
        iso9613_SrcToBarrierEdge & "," & iso9613_RecToBarrierEdge & "," & T_ParamRng(0) & "," & iso9613_DoubleDiffraction & "," & iso9613_BarrierThickness & "," & _
        iso9613_BarrierHeightReceiverSide & "," & iso9613_MultiSource & ",3)"
    ElseIf Barrier_Method = "KurzeAnderson" Then
    Cells(Selection.Row, T_ParamStart) = Barrier_BarrierHeight
    BuildFormula "BarrierAtten_" & Barrier_Method & "(" & T_FreqStartRng & "," & Barrier_SourceToBarrier & "," & Barrier_SourceHeight & "," & _
        Barrier_GroundUnderSrc & "," & Barrier_RecToBarrier & "," & Barrier_ReceiverHeight & "," & _
        Barrier_GroundUnderRec & "," & T_ParamRng(0) & ")" 'todo: option for multi-path
    ElseIf Barrier_Method = "Menounou" Then
    Cells(Selection.Row, T_ParamStart) = Barrier_BarrierHeight
    BuildFormula "BarrierAtten_" & Barrier_Method & "(" & T_FreqStartRng & "," & Barrier_SourceToBarrier & "," & Barrier_SourceHeight & "," & _
        Barrier_GroundUnderSrc & "," & Barrier_RecToBarrier & "," & Barrier_ReceiverHeight & "," & _
        Barrier_GroundUnderRec & "," & T_ParamRng(0) & ",""" & Barrier_SpreadingType & """)"
    Else
    MsgBox "Method not found!", vbOKOnly, "Error - Calc method"
    End If


End Sub

