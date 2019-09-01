Attribute VB_Name = "Noise"
Public ductL As Single
Public ductW As Single
Public ductH As Single
Public ductShape As String
Public ductMethod As String
Public ductLiningThickness As Single
Public roomType As String
Public roomL As Double
Public roomW As Double
Public roomH As Double
Public roomLossType As String
Public ductA1 As Double
Public ductA2 As Double
Public ductSplitType As String
Public btnOkPressed As Boolean
Public regenNoiseElement As String
Public elbowLining As String
Public elbowShape As String
Public elbowVanes As String
Public ERL_Area As Single
Public ERL_Mode As String
Public ERL_Termination As String

Public SilencerModel As String
Public SilencerIL() As Double
Public SilLength As Double
Public SilFA As Double
Public SilSeries As String

Public LouvreModel As String
Public LouvreIL() As Double
Public LouvreLength As Double
Public LouvreFA As String
Public LouvreSeries As String

Public SolverRow As Integer
Public PlaneH As Double
Public PlaneL As Double
Public PlaneDist As Double
Public MaterialDensity As Single
Public DuctWallThickness As Single

Public PlenumL As Long
Public PlenumW As Long
Public PlenumH As Long
Public DuctInL As Single
Public DuctInW As Single
Public DuctOutL As Single
Public DuctOutW As Single
Public PlenumQ As Integer
Public r_h As Long
Public r_v As Long
Public PlenumLiningType As String
Public UnlinedType As String
Public PlenumWallEffect As String
Public PlenumElbowEffect As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Function AirAbsorb(freq As String, Distance As Integer, Optional temp As Integer)
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

Function AWeightCorrections(freq)
Dim dBAAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

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

Function CWeightCorrections(freq)
Dim dBCAdjustment As Variant
Dim freqTitles As Variant
Dim freqTitlesAlt As Variant
Dim ArrayIndex As Integer

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

'legacy reasons, keep the old function and forward on to the new function
Function GetASHRAE(freq As String, L As Long, w As Long, DuctType As String, Distance As Double)
GetASHRAE = GetASHRAEDuct(freq, L, w, DuctType, Distance)
End Function

Function GetASHRAEDuct(freq As String, H As Long, w As Long, DuctType As String, Length As Double)
'On Error GoTo closefile
Dim ReadStr() As String
Dim i As Integer
Dim SplitStr() As String
Dim splitVal() As Double
Dim CurrentType As String
Dim InputArea As Double
'Get Array from text
Close #1

Call GetSettings

Open ASHRAE_DUCT For Input As #1  'global

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then
        
            'convert to values
            For Col = 0 To UBound(SplitStr)
                If SplitStr(Col) <> "" Then
                ReDim Preserve splitVal(Col)
                splitVal(Col) = CDbl(SplitStr(Col))
                End If
            Next Col
            
            ReDim Preserve splitVal(Col + 1)
            
                If Right(DuctType, 1) = "R" Then 'RECTANGULAR DUCT
                ReadArea = splitVal(0) * splitVal(1)
                InputArea = H * w
                ElseIf Right(DuctType, 1) = "C" Then 'CIRCULAR DUCT
                ReadArea = WorksheetFunction.Pi * ((splitVal(0) / 2) ^ 2)
                InputArea = WorksheetFunction.Pi * ((H / 2) ^ 2)
                Else
                'msg = MsgBox("UNKNOWN TYPE", vbOKOnly, "You done f**ked up now.")
                End If
            
            If InputArea <= ReadArea And CurrentType = DuctType Then
            'Debug.Print "AREA found - line " & i
                'select correct frequency band
'                    For x = 0 To 9
'                    Debug.Print splitVal(x)
'                    Next x
                
                Select Case freq 'catch for both kinds of header
                Case Is = "63"
                    If Right(CurrentType, 1) = "R" Then 'RECTANGULAR DUCT
                    GetASHRAEDuct = splitVal(2) * -Length / 2 ' assumes half the attenuation from 125Hz band at 63Hz band
                    ElseIf Right(CurrentType, 1) = "C" Then 'CIRCULAR DUCT
                    GetASHRAEDuct = splitVal(1) * -Length
                    End If
                Case Is = "125"
                GetASHRAEDuct = splitVal(2) * -Length
                Case Is = "250"
                GetASHRAEDuct = splitVal(3) * -Length
                Case Is = "500"
                GetASHRAEDuct = splitVal(4) * -Length
                Case Is = "1k"
                GetASHRAEDuct = splitVal(5) * -Length
                Case Is = "2k"
                GetASHRAEDuct = splitVal(6) * -Length
                Case Is = "4k"
                GetASHRAEDuct = splitVal(7) * -Length
                Case Is = 1000
                GetASHRAEDuct = splitVal(5) * -Length
                Case Is = 2000
                GetASHRAEDuct = splitVal(6) * -Length
                Case Is = 4000
                GetASHRAEDuct = splitVal(7) * -Length
                Case Else
                GetASHRAEDuct = ""
                End Select
                
                'Floor the value, duct attenuation shouldn't be above 40dB
                If GetASHRAEDuct < -40 Then
                GetASHRAEDuct = -40
                End If
                
            found = True '<-this will end the loop
            End If
            
            
        Else '* is the type identifier
        'ReDim Preserve SplitVal(1)
        CurrentType = Right(SplitStr(0), Len(SplitStr(0)) - 1)
        'Debug.Print "TYPE: " & currentType
        End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function

Function GetReynoldsDuct(freq As String, H As Double, w As Double, thickness As Double, L As Double)

Dim PonA As Double
Dim Attn As Double
Dim IL As Double

'Static Values from NEBB book
b = Array(0.0133, 0.0574, 0.271, 1.0147, 1.77, 1.392, 1.518, 1.581)
C = Array(1.959, 1.41, 0.824, 0.5, 0.695, 0.802, 0.451, 0.219)
d = Array(0.917, 0.941, 1.079, 1.087, 0, 0, 0, 0)

'convert to millimetres to metres
H = H / 1000
w = w / 1000

P = (H * 2) + (w * 2) 'perimeter
a = H * w 'area
PonA = P / a 'used later

f = freqStr2Num(freq)

i = GetOctaveColumnIndex(freq)


    'catch error: frequency bands not defined
    If i = 999 Then
    GetReynoldsDuct = ""
    Else
    
    'equation 5.16
    If thickness = 0 Then 'don't apply this correction
    IL = 0
    Else
    IL = (3.281 * b(i)) * ((0.305 * PonA) ^ C(i)) * ((0.039 * thickness) ^ d(i)) * L
    End If
    
        'applies from 500Hz octave band and up
        If f <= 250 Then
            If PonA >= 10 Then
            'equation 5.13
            Attn = 55.8 * ((0.305 * PonA) ^ -0.25) * (f ^ -0.85) * L
            Else
            'equation 5.14
            Attn = 5.38 * ((0.305 * PonA) ^ 0.73) * (f ^ -0.58) * L
            End If
        Else
        'equation 5.15
        Attn = 0.066 * ((0.305 * PonA) ^ 0.8) * L
        End If
        
        'Top out at 40dB attenuation
        'Remebmber, all losses are **negative**
        If IL + Attn > 40 Then
        GetReynoldsDuct = -40
        Else
        GetReynoldsDuct = Round((IL + Attn) * -1, 1)
        End If
    
    End If

End Function



Function GetDuctBreakout(freq As String, H As Single, w As Single, L As Single, MaterialDensity As Single, DuctWallThickness As Single) 'H and W in mm, L in m, SurfaceMass in kg/m2

Dim TLoutMin As Single
Dim TLout As Single
Dim SurfaceMass As Single
Dim fL As Long

f = freqStr2Num(freq)
fL = 613000# / ((w * H) ^ 0.5) 'W and H in mm

SurfaceMass = MaterialDensity * DuctWallThickness / 1000 'duct wall thickness is in mm

TLoutMin = 10 * Application.WorksheetFunction.Log10(2 * L * 1000 * ((1 / w) + (1 / H))) 'length in metres, needs to X1000


    If SurfaceMass <> 0 And w <> 0 And H <> 0 Then
        If f < fL Then
        TLout = 10 * Application.WorksheetFunction.Log10((f * (SurfaceMass ^ 2)) / (w + H)) + 17 'equation 6.11
        Else
        TLout = 20 * Application.WorksheetFunction.Log10(f * SurfaceMass) - 45 'equation 6.12
        End If
        
        'TLout can't be greater than 45dB (but why?)
        If TLout > 45 Then TLout = 45
        
        
        If TLout > TLoutMin Then
        GetDuctBreakout = TLoutMin - TLout 'comes out as negative
        Else
        GetDuctBreakout = 0
        End If
    End If
    
End Function


Function GetFlexDuct(freq As String, dia As Integer, L As Double)
On Error GoTo closefile
Dim ReadStr() As String
Dim i As Integer
Dim SplitStr() As String
Dim splitVal() As Double
Dim Col As Integer

Call GetSettings

Open ASHRAE_FLEX For Input As #1  'global

i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
        If Left(SplitStr(0), 1) <> "*" Then 'titles
        
            'convert to values
            For Col = 0 To UBound(SplitStr)
                If SplitStr(Col) <> "" Then
                ReDim Preserve splitVal(Col)
                splitVal(Col) = CDbl(SplitStr(Col))
                End If
            Next Col
            
            ReDim Preserve splitVal(Col + 1)
            
                If splitVal(0) = dia And splitVal(1) = L Then
                    Select Case freq
                    Case Is = "63"
                    GetFlexDuct = -splitVal(2)
                    Case Is = "125"
                    GetFlexDuct = -splitVal(3)
                    Case Is = "250"
                    GetFlexDuct = -splitVal(4)
                    Case Is = "500"
                    GetFlexDuct = -splitVal(5)
                    Case Is = "1k"
                    GetFlexDuct = -splitVal(6)
                    Case Is = "2k"
                    GetFlexDuct = -splitVal(7)
                    Case Is = "4k"
                    GetFlexDuct = -splitVal(8)
                    Case Is = 1000
                    GetFlexDuct = -splitVal(6)
                    Case Is = 2000
                    GetFlexDuct = -splitVal(7)
                    Case Is = 4000
                    GetFlexDuct = -splitVal(8)
                    Case Else
                    GetFlexDuct = ""
                    End Select
                End If
        End If
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function

Function GetERL(TerminationType As String, freq As String, DuctArea As Double) 'legacy reasons, keep the old function and forward on to the new function
GetERL = GetERL_ASHRAE(TerminationType, freq, DuctArea)
End Function

Function GetERL_ASHRAE(TerminationType As String, freq As String, DuctArea As Double)

Dim dia As Double
Dim A1 As Double
Dim A2 As Double
Dim f As Double
    If DuctArea <> 0 Then
    dia = (4 * DuctArea / Application.WorksheetFunction.Pi) ^ 0.5 'eqn 11
    
    f = freqStr2Num(freq)
    C0 = 343
        'table 28 of ASHRAE
        If TerminationType = "Flush" Then
        A1 = 0.7
        A2 = 2
        ElseIf TerminationType = "Free" Then
        A1 = 1
        A2 = 2
        End If
    GetERL_ASHRAE = -10 * Application.WorksheetFunction.Log10(1 + ((A1 * C0) / (f * dia * Application.WorksheetFunction.Pi)) ^ A2)
    Else
    GetERL_ASHRAE = 0
    End If
End Function

Function GetERL_NEBB(TerminationType As String, freq As String, DuctArea As Double)
Dim dia As Double
Dim A1 As Double
Dim A2 As Double
Dim f As Double
    If DuctArea <> 0 Then
    dia = (4 * DuctArea / Application.WorksheetFunction.Pi) ^ 0.5
    
    f = freqStr2Num(freq)
    C0 = 343
        If TerminationType = "Flush" Then
        A1 = 0.8
        A2 = 1.88
        ElseIf TerminationType = "Free" Then
        A1 = 1
        A2 = 1.88
        End If
    
    GetERL_NEBB = -10 * Application.WorksheetFunction.Log10(1 + ((A1 * C0) / (f * dia * Application.WorksheetFunction.Pi)) ^ A2)
    Else
    GetERL_NEBB = 0
    End If
End Function

Function GetRegenNoise(freq As String, Condition As String, Velocity As Double, Element As String)
On Error GoTo closefile
Dim ReadStr() As String
Dim splitVal() As Double
Dim Col As Integer

f = freqStr2Num(freq)

Call GetSettings

Open ASHRAE_REGEN For Input As #1  'global

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then
            
            If CurrentType = Element Then 'elbow, damper, or transition
                If SplitStr(0) = Condition And CDbl(SplitStr(1)) = Velocity Then 'vanes/no vanes
                
                'convert to values
                For Col = 1 To UBound(SplitStr)
                If SplitStr(Col) <> "" Then
                ReDim Preserve splitVal(Col)
                splitVal(Col) = CDbl(SplitStr(Col))
                End If
                Next Col
                
                    Select Case freq
                    Case Is = "63"
                    GetRegenNoise = splitVal(2)
                    Case Is = "125"
                    GetRegenNoise = splitVal(3)
                    Case Is = "250"
                    GetRegenNoise = splitVal(4)
                    Case Is = "500"
                    GetRegenNoise = splitVal(5)
                    Case Is = "1k"
                    GetRegenNoise = splitVal(6)
                    Case Is = "2k"
                    GetRegenNoise = splitVal(7)
                    Case Is = "4k"
                    GetRegenNoise = splitVal(8)
                    Case Is = 1000
                    GetRegenNoise = splitVal(6)
                    Case Is = 2000
                    GetRegenNoise = splitVal(7)
                    Case Is = 4000
                    GetRegenNoise = splitVal(8)
                    Case Else
                    GetRegenNoise = ""
                    End Select
                    
                End If
            End If
            
            'ReDim Preserve splitVal(Col + 1)
            
        Else '* is the type identifier
        CurrentType = Right(SplitStr(0), Len(SplitStr(0)) - 1)
        End If
        
            'catch for 0
            If GetRegenNoise = 0 Then
            GetRegenNoise = "-"
            End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1

End Function

Function GetRoomLoss(fstr As String, L As Double, w As Double, H As Double, roomType As String)
Dim alpha() As Variant
Dim alpha_av As Double
Dim Rc As Double
'freq = freqStr2Num(fstr)

    Select Case roomType
    Case Is = "Live"
    alpha = Array(0.2, 0.18, 0.14, 0.11, 0.1, 0.1, 0.1, 0.1, 0.1)
    Case Is = "Av. Live"
    alpha = Array(0.19, 0.18, 0.17, 0.14, 0.15, 0.15, 0.14, 0.13, 0.12)
    Case Is = "Average"
    alpha = Array(0.2, 0.18, 0.19, 0.19, 0.2, 0.23, 0.22, 0.21, 0.2)
    Case Is = "Av. Dead"
    alpha = Array(0.21, 0.2, 0.23, 0.24, 0.25, 0.28, 0.27, 0.26, 0.25)
    Case Is = "Dead"
    alpha = Array(0.22, 0.2, 0.28, 0.3, 0.4, 0.47, 0.45, 0.44, 0.45)
    End Select
    
    
    Select Case fstr
    Case Is = "31.5"
    bandIndex = 0
    Case Is = "63"
    bandIndex = 1
    Case Is = "125"
    bandIndex = 2
    Case Is = "250"
    bandIndex = 3
    Case Is = "500"
    bandIndex = 4
    Case Is = "1k"
    bandIndex = 5
    Case Is = "2k"
    bandIndex = 6
    Case Is = "4k"
    bandIndex = 7
    Case Is = "8k"
    bandIndex = 8
    Case Is = 1000
    bandIndex = 5
    Case Is = 2000
    bandIndex = 6
    Case Is = 4000
    bandIndex = 7
    Case Is = 8000
    bandIndex = 8
    End Select
        
    S_total = (L * w * 2) + (L * H * 2) + (w * H * 2)
    alpha_av = ((L * w * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) + (w * H * alpha(bandIndex) * 2)) / S_total
    Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)
    'Debug.Print "Room Contant " Rc
        If Rc <> 0 Then
        GetRoomLoss = 10 * Application.WorksheetFunction.Log10(4 / Rc)
        Else
        GetRoomLoss = 0
        End If
End Function


Function GetRoomLossRT(fstr As String, L As Double, w As Double, H As Double, RT_Type As String)
Dim RT() As Variant
Dim alpha_av As Double
Dim Rc As Double
'freq = freqStr2Num(fstr)

'Alpha values are based on getting the desired midfrequency reverberation time
    Select Case RT_Type
    Case Is = "<0.2 sec"
    alpha = Array(0, 0, 0.21, 0.277, 0.331, 0.385, 0.435, 0.446, 0)
    Case Is = "0.2 to 0.5 sec"
    alpha = Array(0, 0, 0.125, 0.138, 0.183, 0.233, 0.288, 0.296, 0)
    Case Is = "0.5 to 1 sec"
    alpha = Array(0, 0, 0.109, 0.112, 0.137, 0.18, 0.214, 0.225, 0)
    Case Is = "1 to 1.5 sec"
    alpha = Array(0, 0, 0.057, 0.056, 0.058, 0.069, 0.08, 0.082, 0)
    Case Is = "1.5 to 2 sec"
    alpha = Array(0, 0, 0.053, 0.053, 0.06, 0.08, 0.095, 0.1, 0)
    Case Is = ">2 sec"
    alpha = Array(0, 0, 0.063, 0.052, 0.036, 0.041, 0.035, 0.04, 0)
    End Select
    
    
    Select Case fstr
    Case Is = "31.5"
    bandIndex = 0
    Case Is = "63"
    bandIndex = 1
    Case Is = "125"
    bandIndex = 2
    Case Is = "250"
    bandIndex = 3
    Case Is = "500"
    bandIndex = 4
    Case Is = "1k"
    bandIndex = 5
    Case Is = "2k"
    bandIndex = 6
    Case Is = "4k"
    bandIndex = 7
    Case Is = "8k"
    bandIndex = 8
    Case Is = 1000
    bandIndex = 5
    Case Is = 2000
    bandIndex = 6
    Case Is = 4000
    bandIndex = 7
    Case Is = 8000
    bandIndex = 8
    End Select
    
    S_total = (L * w * 2) + (L * H * 2) + (w * H * 2)
    alpha_av = ((L * w * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) + (w * H * alpha(bandIndex) * 2)) / S_total
    Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)

        If Rc <> 0 Then
        GetRoomLossRT = 10 * Application.WorksheetFunction.Log10(4 / Rc)
        Else
        GetRoomLossRT = 0
        End If

End Function

Function GetElbowLoss(fstr As String, w As Double, elbowShape As String, DuctLining As String, VaneType As String)
Dim Unlined() As Variant
Dim Lined() As Variant
Dim RadiusBend() As Variant
Dim freq As Double
Dim FW As Double
Dim ArrayIndex As Integer
Dim linedDuct As Boolean
Dim Vanes As Boolean

    If DuctLining = "Lined" Then
    linedDuct = True
    ElseIf DuctLining = "Unlined" Then
    linedDuct = False
    End If
    
    If VaneType = "Vanes" Then
    Vanes = True
    ElseIf VaneType = "No Vanes" Then
    Vanes = False
    End If
    

Unlined = Array(0, -1, -5, -8, -4, -6) 'table 22 of ASHRAE
Lined = Array(0, -1, -6, -11, -10, -10)

UnlinedV = Array(0, -1, -4, -6, -4) 'table 24 of ASHRAE
LinedV = Array(0, -1, -4, -7, -7)

RadiusBend = Array(0, -1, -2, -3) 'table 23 of ASHRAE

freq = freqStr2Num(fstr)
FW = (freq / 1000) * w

    Select Case elbowShape
    Case Is = "Square"
        If Vanes = False Then
            Select Case FW
            Case Is < 48
            ArrayIndex = 0
            Case Is < 96
            ArrayIndex = 1
            Case Is < 190
            ArrayIndex = 2
            Case Is < 380
            ArrayIndex = 3
            Case Is < 760
            ArrayIndex = 4
            Case Is >= 760
            ArrayIndex = 5
            End Select
            
                If linedDuct = True Then
                GetElbowLoss = Lined(ArrayIndex)
                Else 'LinedDuct = False
                GetElbowLoss = Unlined(ArrayIndex)
                End If
                
        Else 'vanes=true
            Select Case FW
            Case Is < 48
            ArrayIndex = 0
            Case Is < 96
            ArrayIndex = 1
            Case Is < 190
            ArrayIndex = 2
            Case Is < 380
            ArrayIndex = 3
            Case Is >= 380
            ArrayIndex = 4
            End Select
            
                If linedDuct = True Then
                GetElbowLoss = LinedV(ArrayIndex)
                Else 'LinedDuct = False
                GetElbowLoss = UnlinedV(ArrayIndex)
                End If
            
        End If
        
    Case Is = "Radius"
        Select Case FW
        Case Is < 48
        ArrayIndex = 0
        Case Is < 96
        ArrayIndex = 1
        Case Is < 190
        ArrayIndex = 2
        Case Is >= 190
        ArrayIndex = 3
        End Select
        
    GetElbowLoss = RadiusBend(ArrayIndex)
            
    End Select

End Function

Function GetASHRAEPlenumLoss(fstr As String, L As Long, w As Long, H As Long, DuctInL As Single, DuctInW As Single, DuctOutL As Single, DuctOutW As Single, _
q As Integer, r_h As Long, r_v As Long, PlenumLiningType As String, UnlinedType As String, wallEffect As String, applyElbowEffect As Boolean, Optional OneThirdsMode As Boolean)

Dim f_OneUp As Integer
Dim f_OneDown As Integer
Dim f As Integer
Dim Loss1 As Double
Dim Loss2 As Double
Dim Loss3 As Double

f = freqStr2Num(fstr)

If IsMissing(OneThirdsMode) Then OneThirdsMode = False

    If OneThirdsMode = True Then
    GetASHRAEPlenumLoss = GetASHRAEPlenumLoss_OneThirdOctave(f, L, w, H, DuctInL, DuctInW, DuctOutL, DuctOutW, q, r_h, r_v, PlenumLiningType, UnlinedType, wallEffect, applyElbowEffect)
    Else
    f_OneUp = GetAdjacentFrequency(f, "Up")
    f_OneDown = GetAdjacentFrequency(f, "Down")
    'get for each one third octave and then Tl average them
    Loss1 = GetASHRAEPlenumLoss_OneThirdOctave(f_OneDown, L, w, H, DuctInL, DuctInW, DuctOutL, DuctOutW, q, r_h, r_v, PlenumLiningType, UnlinedType, wallEffect, applyElbowEffect)
    Loss2 = GetASHRAEPlenumLoss_OneThirdOctave(f, L, w, H, DuctInL, DuctInW, DuctOutL, DuctOutW, q, r_h, r_v, PlenumLiningType, UnlinedType, wallEffect, applyElbowEffect)
    Loss3 = GetASHRAEPlenumLoss_OneThirdOctave(f_OneUp, L, w, H, DuctInL, DuctInW, DuctOutL, DuctOutW, q, r_h, r_v, PlenumLiningType, UnlinedType, wallEffect, applyElbowEffect)
    'Note: losses are negative already so no need for negatives sign in formula
    GetASHRAEPlenumLoss = 10 * Application.WorksheetFunction.Log10((1 / 3) * ((10 ^ (Loss1 / 10)) + (10 ^ (Loss2 / 10)) + (10 ^ (Loss3 / 10))))
    End If

End Function


Function GetASHRAEPlenumLoss_OneThirdOctave(f As Integer, L As Long, w As Long, H As Long, DuctInL As Single, DuctInW As Single, DuctOutL As Single, DuctOutW As Single, _
q As Integer, r_h As Long, r_v As Long, PlenumLiningType As String, UnlinedType As String, wallEffect As String, applyElbowEffect As Boolean)

Dim Stotal As Single
Dim InletArea As Single
Dim OutletArea As Single
Dim R As Single 'inlet to outlet offset distance
Dim alphaTotal(7) As Single
Dim AbsorptionArea(7) As Single
Dim PlenumVolume As Single
Dim offsetangle As Single
Dim b As Single
Dim N As Single
'Dim f As Integer
Dim f_co As Single 'cutoff frequency
Dim OAE As Single
Dim DuctInL_OT As Single
Dim DuctInW_OT As Single
Dim WallEffectIndex As Integer
Dim W_e As Single
Dim AngleEffect As Single

'''''''''''''''''''''''''''''''''''''''''''''''''
'CONSTANTS
'''''''''''''''''''''''''''''''''''''''''''''''''

'Values from ASHRAE equation 5
b = 3.505
N = -0.359
'If IsMissing(OneThirdsMode) Then OneThirdsMode = False

'Lining Materials, from ASHRAE table 12
Concrete = Array(0.01, 0.01, 0.01, 0.02, 0.02, 0.02, 0.03)
Bare_Sheet_Metal = Array(0.04, 0.04, 0.04, 0.05, 0.05, 0.05, 0.07)
FG25 = Array(0.05, 0.11, 0.28, 0.68, 0.9, 0.93, 0.96)
FG50 = Array(0.1, 0.17, 0.86, 1, 1, 1, 1)
FG75 = Array(0.3, 0.53, 1, 1, 1, 1, 1)
FG100 = Array(0.5, 0.84, 1, 1, 1, 1, 0.97)

'''''''''''''''''''''''''''''''''''''''''''''''''
'CONVERSION
'''''''''''''''''''''''''''''''''''''''''''''''''

'Convert units and types
'f = freqStr2Num(fstr) <---- not required as input to function is already integer
DuctInL_OT = DuctInL / 1000 'convert to metres
DuctInW_OT = DuctInW / 1000 'convert to metres
DuctOutL_OT = DuctOutL / 1000 'convert to metres
DuctOutW_OT = DuctOutW / 1000 'convert to metres

'''''''''''''''''''''''''''''''''''''''''''''''''
'CALC
'''''''''''''''''''''''''''''''''''''''''''''''''

'cutoff frequency
f_co = GetCutoffFrequency(DuctInL_OT, DuctInW_OT)
'Debug.Print "Cutoff Frequency = " & Round(f_co, 1) & "Hz"

'Areas and Volumes
InletArea = DuctInL_OT * DuctInW_OT
OutletArea = DuctOutL_OT * DuctOutW_OT
PlenumVolume = (L / 1000) * (w / 1000) * (H / 1000) 'input in mm

'Surface area
Stotal = GetPlenumSurfaceArea(L, w, H, InletArea, OutletArea) ' Stotal doesn't include inlet and outlet area
'Debug.Print "S_total = " & Stotal

    'Linings - selected from static tables as array variables
    Select Case PlenumLiningType
    Case Is = "Concrete"
    PlenumLining = Concrete
    Case Is = "Bare Sheet Metal"
    PlenumLining = Bare_Sheet_Metal
    Case Is = "25mm fibreglass"
    PlenumLining = FG25
    Case Is = "50mm fibreglass"
    PlenumLining = FG50
    Case Is = "75mm fibreglass"
    PlenumLining = FG75
    Case Is = "100mm fibreglass"
    PlenumLining = FG100
    Case Else 'Error with lining type. MsgBox allows for debug to be called.
    msg = MsgBox("Error - no lining type selected.", vbOKOnly, "Check types")
    End
    End Select
    
    Select Case UnlinedType
    Case Is = "Concrete"
    Unlined = Concrete
    Case Is = "Bare Sheet Metal"
    Unlined = Bare_Sheet_Metal
    End Select
    
'    'catch blank lining error
'    If UnlinedType = "" And PlenumLiningType = "" Then
'    End
'    End If
    
    For i = 0 To UBound(alphaTotal) - 1 '7th column not used
    alphaTotal(i) = (((InletArea + OutletArea) * Unlined(i)) + (Stotal * PlenumLining(i))) / (InletArea + OutletArea + Stotal)
    AbsorptionArea(i) = (OutletArea * (1 - alphaTotal(i))) / (Stotal * alphaTotal(i))
    Next i
        
'Distance from inlet to outlet
R = GetPlenumDistanceR(r_h, r_v, L)
'Debug.Print "Offset Distance, R = " & Round(r, 1)

        
    If applyElbowEffect = True Then '90 degree bend, use the elbow effect method in ASHRAE
    AngleEffect = GetPlenumElbowEffect(f, f_co)
    Else
    'Offset Angle
    offsetangle = GetPlenumAngleTheta(L, R)
    'Debug.Print "Offset Angle = " & Round(OffsetAngle, 2)
    'Offset Angle Effect
    AngleEffect = GetPlenumOAE(f, f_co, offsetangle)
    End If
    

    If f < f_co Then
    
        If wallEffect = "False" Or wallEffect = "" Then
        WallEffectIndex = 0
        Else
        WallEffectIndex = CInt(Left(wallEffect, 1))
        End If
    
    W_e = GetPlenumWallEffect(f, WallEffectIndex)
    
    'Debug.Print "Wall effect = " & W_e
    A_f = GetPlenumAreaCoefficient(f, PlenumVolume)
    'Debug.Print "Area Coefficient, A_f = " & A_f
    GetASHRAEPlenumLoss_OneThirdOctave = -1 * Application.WorksheetFunction.Min((A_f * Stotal) + W_e + AngleEffect, 20) 'limit to 20dB, output is negative
    
    Else 'f>=f_co
    A_index = GetOctaveBandIndex(f)
    GetASHRAEPlenumLoss_OneThirdOctave = -1 * (b * (((OutletArea * q / (4 * Application.WorksheetFunction.Pi() * (R ^ 2))) + AbsorptionArea(A_index)) ^ N) + AngleEffect)
    End If
        
End Function


Function GetPlenumDistanceR(r_h As Long, r_v As Long, L As Long) As Single 'all units are input in millimetres
GetPlenumDistanceR = (((r_v / 1000) ^ 2) + ((r_h / 1000) ^ 2) + ((L / 1000) ^ 2)) ^ 0.5
End Function

Function GetPlenumAngleTheta(L As Long, R As Single) 'L is in millimetres
Dim PlenumL As Single
PlenumL = L / 1000
    If PlenumL / R >= -1 And PlenumL / R <= 1 Then 'between -1 and 1
    GetPlenumAngleTheta = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(PlenumL / R))
    Else
    GetPlenumAngleTheta = 0
    End If
End Function

Function GetCutoffFrequency(L As Single, w As Single) As Single 'all units are input in metres
GetCutoffFrequency = 343 / (2 * Application.Max(L, w)) 'cutoff frequency
End Function

Function GetPlenumSurfaceArea(L As Long, w As Long, H As Long, InletArea As Single, OutletArea As Single) As Single
GetPlenumSurfaceArea = (2 * L * w / 1000000) + (2 * w * H / 1000000) + (2 * H * L / 1000000) - InletArea - OutletArea 'inputs are in mm, which are squared => correction is 1000x1000 = 1million
End Function

Function GetPlenumAreaCoefficient(f_input As Integer, Vol As Single)

'From ASHRAE Table 13
SmallPlenum = Array(1.4, 1#, 1.1, 1.1, 2.3, 2.4, 2, 1#, 2.2, 0.7, 0.7, 1.1) '50 to 500Hz
LargePlenum = Array(0.3, 0.3, 0.3, 0.3, 0.4, 0.4, 0.3, 0.4, 0.3, 0.2, 0.2) '50 to 500Hz
i = GetArrayIndex_TO_50Hz(f_input)

    If i <= 10 Then
        If Vol < 1.5 Then 'm^3
        GetPlenumAreaCoefficient = SmallPlenum(i)
        Else 'Vol>1.5m^3
        GetPlenumAreaCoefficient = LargePlenum(i)
        End If
    Else
    GetPlenumAreaCoefficient = 0
    End If


End Function

Function GetPlenumOAE(f_input As Integer, f_co As Single, Angle_input As Single)

    'from Table 14 of ASHRAE
    If f_input <= f_co Then
        Select Case f_input
        Case Is = 50
        OAEtable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 63
        OAEtable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 80
        OAEtable = Array(0, 0, -1, -3, -4, -6)
        Case Is = 100
        OAEtable = Array(0, 1, 0, -2, -3, -6)
        Case Is = 125
        OAEtable = Array(0, 1, 0, -2, -4, -6)
        Case Is = 160
        OAEtable = Array(0, 0, -1, -2, -3, -4)
        Case Is = 200
        OAEtable = Array(0, 0, -1, -2, -3, -5)
        Case Is = 250
        OAEtable = Array(0, 1, 2, 3, 5, 7)
        Case Is = 315
        OAEtable = Array(0, 4, 6, 8, 10, 14)
        Case Is = 400
        OAEtable = Array(0, 2, 4, 6, 9, 13)
        Case Is = 500
        OAEtable = Array(0, 1, 3, 6, 10, 15)
        Case Is = 630
        OAEtable = Array(0, 0, 0, 0, 0, 0)
        End Select
    Else 'f_input>f_co
        Select Case f_input
        Case Is = 200
        OAEtable = Array(0, 1, 4, 9, 14, 20)
        Case Is = 250
        OAEtable = Array(0, 2, 4, 8, 13, 19)
        Case Is = 315
        OAEtable = Array(0, 1, 2, 3, 4, 5)
        Case Is = 400
        OAEtable = Array(0, 1, 2, 3, 4, 6)
        Case Is = 500
        OAEtable = Array(0, 0, 1, 2, 4, 5)
        Case Is = 630
        OAEtable = Array(0, 1, 2, 3, 5, 7)
        Case Is = 800
        OAEtable = Array(0, 1, 2, 2, 3, 3)
        Case Is = 1000
        OAEtable = Array(0, 1, 2, 4, 6, 9)
        Case Is = 1250
        OAEtable = Array(0, 0, 2, 4, 6, 9)
        Case Is = 1600
        OAEtable = Array(0, 0, 1, 1, 2, 3)
        Case Is = 2000
        OAEtable = Array(0, 1, 2, 4, 7, 10)
        Case Is = 2500
        OAEtable = Array(0, 1, 2, 3, 5, 8)
        Case Is = 3150
        OAEtable = Array(0, 0, 2, 4, 6, 9)
        Case Is = 4000
        OAEtable = Array(0, 0, 2, 5, 8, 12)
        Case Is = 5000
        OAEtable = Array(0, 0, 3, 6, 10, 15)
        End Select
    End If

    'get x values for interpolation
    Select Case Angle_input
    Case Is = 0
    x1 = 0
    x2 = 0
    Case Is < 15
    x1 = 0
    x2 = 15
    i = 0
    Case Is < 22.5
    x1 = 15
    x2 = 22.5
    i = 1
    Case Is < 30
    x1 = 22.5
    x2 = 30
    i = 2
    Case Is < 37.5
    x1 = 30
    x2 = 37.5
    i = 3
    Case Is < 45
    x1 = 37.5
    x2 = 45
    i = 4
    Case Is >= 45
    x1 = 0
    x2 = 0
    i = 5
    End Select
    
    'calculate OAE
    If (x1 = 0 And x2 = 0) Or y1 = 0 Or y2 = 0 Or y1 = Empty Or y2 = Empty Then
    GetPlenumOAE = 0
    Else
    y1 = OAEtable(i)
    y2 = OAEtable(i + 1)
    'interpolate things
    Slope = (y2 - y1) / (x2 - x1)
    GetPlenumOAE = y1 + (Slope * (Angle_input - x1))
    End If


End Function

Function GetPlenumWallEffect(f As Integer, WallType As Integer)
'Table 15 of ASHRAE - Low Frequency Characdteristics of Plenum TL
i = GetArrayIndex_TO_50Hz(f)
WType1 = Array(1, 1, 2, 2, 2, 3, 4, 5, 6, 8, 9)
WType2 = Array(1, 2, 2, 2, 3, 4, 10, 9, 12, 13, 13)
WType3 = Array(0, 3, 3, 4, 6, 11, 16, 13, 14, 13, 12)
WType4 = Array(1, 7, 9, 12, 11, 15, 12, 14, 14, 13)
WType5 = Array(0, 1, 2, 1, 1, 0, 4, 1, 5, 7, 8)
WType6 = Array(0, 3, 7, 6, 4, 2, 3, 1, 2, 1, 0)

    If i <= 10 Then 'maximum 10 elements in WType arrays
        Select Case WallType
        Case Is = 0
        GetPlenumWallEffect = 0
        Case Is = 1
        GetPlenumWallEffect = WType1(i)
        Case Is = 2
        GetPlenumWallEffect = WType2(i)
        Case Is = 3
        GetPlenumWallEffect = WType3(i)
        Case Is = 4
        GetPlenumWallEffect = WType4(i)
        Case Is = 5
        GetPlenumWallEffect = WType5(i)
        Case Is = 6
        GetPlenumWallEffect = WType6(i)
        End Select
    Else
    GetPlenumWallEffect = 0
    End If


End Function

Function GetPlenumElbowEffect(f As Integer, f_c As Single)
'From Table 15 of ASHRAE - Elbow Effect, from 50Hz One-third octave band
BelowFc = Array(2, 3, 6, 5, 3, 0, -2, -3, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
AboveFc = Array(0, 0, 0, 0, 0, 0, 3, 6, 3, 3, 2, 3, 3, 2, 2, 2, 2, 2, 2, 2, 1)
i = GetArrayIndex_TO_50Hz(f)

    If i <= 20 Then
        If f > f_c Then
        GetPlenumElbowEffect = AboveFc(i)
        Else 'f<=f_c
        GetPlenumElbowEffect = BelowFc(i)
        End If
    Else
    GetPlenumElbowEffect = 0
    End If
    
End Function

Function GetArrayIndex_TO_50Hz(f As Integer)
    Select Case f
    Case Is = 50
    GetArrayIndex_TO_50Hz = 0
    Case Is = 63
    GetArrayIndex_TO_50Hz = 1
    Case Is = 80
    GetArrayIndex_TO_50Hz = 2
    Case Is = 100
    GetArrayIndex_TO_50Hz = 3
    Case Is = 125
    GetArrayIndex_TO_50Hz = 4
    Case Is = 160
    GetArrayIndex_TO_50Hz = 5
    Case Is = 200
    GetArrayIndex_TO_50Hz = 6
    Case Is = 250
    GetArrayIndex_TO_50Hz = 7
    Case Is = 315
    GetArrayIndex_TO_50Hz = 8
    Case Is = 400
    GetArrayIndex_TO_50Hz = 9
    Case Is = 500
    GetArrayIndex_TO_50Hz = 10
    Case Is = 630
    GetArrayIndex_TO_50Hz = 11
    Case Is = 800
    GetArrayIndex_TO_50Hz = 12
    Case Is = 1000
    GetArrayIndex_TO_50Hz = 13
    Case Is = 1250
    GetArrayIndex_TO_50Hz = 14
    Case Is = 1600
    GetArrayIndex_TO_50Hz = 15
    Case Is = 2000
    GetArrayIndex_TO_50Hz = 16
    Case Is = 2500
    GetArrayIndex_TO_50Hz = 17
    Case Is = 3150
    GetArrayIndex_TO_50Hz = 18
    Case Is = 4000
    GetArrayIndex_TO_50Hz = 19
    Case Is = 5000
    GetArrayIndex_TO_50Hz = 20
    End Select
End Function

Function GetOctaveBandIndex(f_input As Integer) 'map a 1/3 octave centre frequency to the relevant 1/1 octave band centre frequency
    Select Case f_input
    Case Is = 50
    GetOctaveBandIndex = 0
    Case Is = 63
    GetOctaveBandIndex = 0
    Case Is = 80
    GetOctaveBandIndex = 0
    Case Is = 100
    GetOctaveBandIndex = 1
    Case Is = 125
    GetOctaveBandIndex = 1
    Case Is = 160
    GetOctaveBandIndex = 1
    Case Is = 200
    GetOctaveBandIndex = 2
    Case Is = 250
    GetOctaveBandIndex = 2
    Case Is = 315
    GetOctaveBandIndex = 2
    Case Is = 400
    GetOctaveBandIndex = 3
    Case Is = 500
    GetOctaveBandIndex = 3
    Case Is = 630
    GetOctaveBandIndex = 3
    Case Is = 800
    GetOctaveBandIndex = 4
    Case Is = 1000
    GetOctaveBandIndex = 4
    Case Is = 1250
    GetOctaveBandIndex = 4
    Case Is = 1600
    GetOctaveBandIndex = 5
    Case Is = 2000
    GetOctaveBandIndex = 5
    Case Is = 2500
    GetOctaveBandIndex = 5
    Case Is = 3150
    GetOctaveBandIndex = 6
    Case Is = 4000
    GetOctaveBandIndex = 6
    Case Is = 5000
    GetOctaveBandIndex = 6
    End Select
End Function

Function GetAdjacentFrequency(f_input As Integer, AdjMode As String)
Dim adjustIndex  As Integer
f_ref = Array(50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000)
    If AdjMode = "Up" Then
    adjustIndex = 1
    ElseIf AdjMode = "Down" Then
    adjustIndex = -1
    Else
    adjustIndex = 0
    End If
    
    For i = LBound(f_ref) To UBound(f_ref)
        If f_ref(i) = f_input Then
        GetAdjacentFrequency = f_ref(i + adjustIndex)
        End If
    Next i

End Function

Function GetDuctArea(inputStr As String)
Dim SplitStr() As String
Dim L As Double
Dim w As Double
SplitStr = Split(inputStr, ",", Len(inputStr), vbTextCompare)
L = CDbl(SplitStr(1))
w = CDbl(SplitStr(2))
GetDuctArea = (L / 1000) * (w / 1000) 'because millimetres
End Function

Function GetDuctParameter(inputStr As String, Parameter As String)
Dim SplitStr() As String
Dim L As Single
Dim w As Single
Dim Area As Single
SplitStr = Split(inputStr, ",", Len(inputStr), vbTextCompare)
L = CSng(SplitStr(1))
w = CSng(SplitStr(2))
Area = (L / 1000) * (w / 1000) 'because millimetres
    Select Case Parameter
    Case Is = "Area"
    GetDuctParameter = Area
    Case Is = "L"
    GetDuctParameter = L
    Case Is = "W"
    GetDuctParameter = w
    End Select
End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Distance(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "Distance Attenuation - point"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=10*LOG($O" & Selection.Row & "/(4*PI()*$N" & Selection.Row & "^2))"
    ParamCol1 = 14
    ParamCol2 = 15
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=10*LOG($AA" & Selection.Row & "/(4*PI()*$Z" & Selection.Row & "^2))"
    ParamCol1 = 26
    ParamCol2 = 27
    End If

ExtendFunction (SheetType)

fmtUserInput SheetType, True

Call ParameterUnmerge(Selection.Row, SheetType)

Cells(Selection.Row, ParamCol1) = 10 'default to 10 metres
Cells(Selection.Row, ParamCol2) = 2 'default to half spherical
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
Cells(Selection.Row, ParamCol2).NumberFormat = "Q=0"

    With Cells(Selection.Row, ParamCol2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="1,2,4,8"
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

Sub DistanceLine(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "Distance Attenuation - line"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=10*LOG($O" & Selection.Row & "/(2*PI()*$N" & Selection.Row & "))"
    ParamCol1 = 14
    ParamCol2 = 15
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=10*LOG($AA" & Selection.Row & "/(2*PI()*$Z" & Selection.Row & "))"
    ParamCol1 = 26
    ParamCol2 = 27
    End If

ExtendFunction (SheetType)

fmtUserInput SheetType, True

Call ParameterUnmerge(Selection.Row, SheetType)

Cells(Selection.Row, ParamCol1) = 10 'default to 10 metres
Cells(Selection.Row, ParamCol2) = 2 'default to half cylindrical
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
Cells(Selection.Row, ParamCol2).NumberFormat = "Q=0"

    With Cells(Selection.Row, ParamCol2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="1,2,4,8"
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


Sub DistancePlane(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'In the near field (approximately `r < a//pi`), the sound level can be approximated as:
'`L_p=L_W-10log_10 S+DI`
'In the line-source intermediate region (approximately `a//pi < r < b//pi`) , the sound level can be approximated as:
'`L_p=L_W-10log_10 S-10log_10(d/(a//pi))+DI`
'In the point-source far region (approximately `r > b//pi`), the sound level can be approximated as:
'`L_p=L_W-10log_10 S-10log_10(a/b)-20log(d/(b//pi))+DI`
'where `r` is the distance from the source, `H` and `L` are the minor and major source dimensions (m), `S=H*L` is the area of the source (m?) and `DI` is the directivity index of the source (dB).
'B&H formula
'=-10*LOG($N9*$O9)+10*LOG(ATAN(($N9*$O9)/(2*$P9*SQRT(($N9^2)+($O9^2)+(4*$P9^2)))))-2

frmPlaneSource.Show

If btnOkPressed = False Then End 'catch cancel

Cells(Selection.Row, 2).Value = "Distance Attenuation - plane"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=-10*LOG(" & PlaneH & "*" & PlaneL & ")+10*LOG(ATAN((" & PlaneH & "*" & PlaneL & ")/(2*$N" & Selection.Row & "*SQRT((" & PlaneH & "^2)+(" & PlaneL & "^2)+(4*$N" & Selection.Row & "^2)))))-2"
    ParamCol1 = 14
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=-10*LOG(" & PlaneH & "*" & PlaneL & ")+10*LOG(ATAN((" & PlaneH & "*" & PlaneL & ")/(2*$Z" & Selection.Row & "*SQRT((" & PlaneH & "^2)+(" & PlaneL & "^2)+(4*$Z" & Selection.Row & "^2)))))-2"
    ParamCol1 = 26
    End If

ExtendFunction (SheetType)

fmtUserInput SheetType, True

Call ParameterMerge(Selection.Row, SheetType)

Cells(Selection.Row, ParamCol1) = PlaneDist
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
    
End Sub

Sub AirAbsorption(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "Air Absorption"
If Left(SheetType, 3) = "OCT" Then
Cells(Selection.Row, 5).Value = "=AirAbsorb(E$6,$N" & Selection.Row & ",$O" & Selection.Row & ")"
ParamCol1 = 14
ParamCol2 = 15
ElseIf Left(SheetType, 2) = "TO" Then
Cells(Selection.Row, 5).Value = "=AirAbsorb(E$6,$Z" & Selection.Row & ",$AA" & Selection.Row & ")"
ParamCol1 = 26
ParamCol2 = 27
End If
ExtendFunction (SheetType)
fmtUserInput SheetType, True
Call ParameterUnmerge(Selection.Row, SheetType)
Cells(Selection.Row, ParamCol1) = 150
Cells(Selection.Row, ParamCol2) = 20
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
Cells(Selection.Row, ParamCol2).NumberFormat = "0""" & chr(176) & "C"""
End Sub



Sub DuctAtten(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmDuctAtten.Show

    If btnOkPressed = False Then
    End
    End If

Cells(Selection.Row, 2).Value = "Duct Attenuation-" & ductMethod

Call ParameterUnmerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Then
    
        If ductMethod = "Reynolds" Then
        'Function GetReynoldsDuct(freq As String, H As Double, W As Double, thickness As Double, L As Double)
        Cells(Selection.Row, 14) = ductLiningThickness
        Cells(Selection.Row, 14).NumberFormat = "0 ""mm"""
        Cells(Selection.Row, 5).Value = "=GetReynoldsDuct(E$6," & ductH & ", " & ductW & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
        ElseIf ductMethod = "ASHRAE" Then
        'Function GetASHRAEDuct(Freq As String, L As Integer, W As Integer, DuctType As String)
        Cells(Selection.Row, 5).Value = "=GetASHRAEDuct(E$6," & ductH & ", " & ductW & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
        Cells(Selection.Row, 14) = ductShape 'from public variable
        Cells(Selection.Row, 14).NumberFormat = xlGeneral
        End If
        
    'same for both methods
    Cells(Selection.Row, 15) = ductL
    Cells(Selection.Row, 15).NumberFormat = "0.0 ""m"""
    
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=GetASHRAEDuct(E$6,$Z" & Selection.Row & ",$AA" & Selection.Row & ")"
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True


    If ductMethod = "ASHRAE" Then
        With Cells(Selection.Row, 14).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="0 R,0 C,25 R,50 R,25 C,50 C"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End If

End Sub

Sub FlexDuct(SheetType As String)

Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "Flex Duct - ASHRAE"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=GetFlexDuct(E$6,$N" & Selection.Row & ",$O" & Selection.Row & ")"
    ParamCol1 = 14
    ParamCol2 = 15
    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
Call ParameterUnmerge(Selection.Row, SheetType)
Cells(Selection.Row, ParamCol1) = 200
Cells(Selection.Row, ParamCol2) = 0.9
Cells(Selection.Row, ParamCol1).NumberFormat = "0 Ø"
Cells(Selection.Row, ParamCol2).NumberFormat = "0.0 ""m"""

With Cells(Selection.Row, ParamCol1).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="100,125,150,175,200,225,250,300,350,400"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With

With Cells(Selection.Row, ParamCol2).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="0.9,1.8,2.7,3.7"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With

End Sub

Sub Area(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "Area Correction: 10log(A)"

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 14).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = 2
    Cells(Selection.Row, 14).NumberFormat = "0 ""m" & chr(178) & """"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 2
    Cells(Selection.Row, 26).NumberFormat = "0 ""m" & chr(178) & """"
    
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True 'IsParamCol
End Sub

Sub TenLogN(SheetType As String)
CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "Multiple sources: 10log(n)"

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 14).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = 2
    Cells(Selection.Row, 14).NumberFormat = """n = ""0"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 2
    Cells(Selection.Row, 26).NumberFormat = """n = ""0"
    
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
End Sub

Sub TenLogOneOnT(SheetType As String)
CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "Time Correction: 10log(1/t)"

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 14).Address(False, True) & "/" & Cells(Selection.Row, 15).Address(False, True) & ")"
    Call ParameterUnmerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = 1
    Cells(Selection.Row, 14).NumberFormat = """t = ""0"
    Cells(Selection.Row, 15) = 2
    Cells(Selection.Row, 15).NumberFormat = """t0 = ""0"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(1/" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterUnmerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 1
    Cells(Selection.Row, 26).NumberFormat = """t = ""0"
    Cells(Selection.Row, 27) = 2
    Cells(Selection.Row, 27).NumberFormat = """t0 = ""0"
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
End Sub

Sub DuctSplit(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmDuctSplit.Show

    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    'Debug.Print "=10*LOG($O" & Selection.Row & "/($O" & Selection.Row & "+$N" & Selection.Row & "))"
    
        Select Case ductSplitType
        
        Case Is = "Area"
        Call ParameterUnmerge(Selection.Row, SheetType)
        Cells(Selection.Row, 14).NumberFormat = "0.0""m" & chr(178) & """"
        Cells(Selection.Row, 15).NumberFormat = "0.0""m" & chr(178) & """"
        Cells(Selection.Row, 14) = ductA1
        Cells(Selection.Row, 15) = ductA2
        Cells(Selection.Row, 5).Value = "=10*LOG($O" & Selection.Row & "/($O" & Selection.Row & "+$N" & Selection.Row & "))"
        Cells(Selection.Row, 2).Value = "Duct Split: 10LOG(A2/(A1+A2))"
        
        Case Is = "Ratio"
        Call ParameterMerge(Selection.Row, SheetType)
        Cells(Selection.Row, 14) = ductA1
        Cells(Selection.Row, 14).NumberFormat = "0"":1"""
        Cells(Selection.Row, 2).Value = "Duct Split: 10LOG(1/R)"
        Cells(Selection.Row, 5).Value = "=10*LOG(1/$N" & Selection.Row & ")"
        
        Case Is = "Percent"
        Call ParameterMerge(Selection.Row, SheetType)
        Cells(Selection.Row, 14).NumberFormat = "0%"
        Cells(Selection.Row, 14) = ductA1
        Cells(Selection.Row, 5).Value = "=10*LOG($N" & Selection.Row & ")"
        Cells(Selection.Row, 2).Value = "Duct Split: 10LOG(P)"
        End Select


    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
End Sub

Sub ERL(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If InStr(1, Cells(Selection.Row - 1, 10).Formula, "GetASHRAEDuct", vbTextCompare) > 0 Or _
    InStr(1, Cells(Selection.Row - 1, 10).Formula, "GetreynoldsDuct", vbTextCompare) > 0 Then
    frmERL.txtL.Value = GetDuctParameter(Cells(Selection.Row - 1, 10).Formula, "L") '1kHz band formula
    frmERL.txtW.Value = GetDuctParameter(Cells(Selection.Row - 1, 10).Formula, "W") '1kHz band formula
    End If

frmERL.Show

    If btnOkPressed = False Then
    End
    End If

Cells(Selection.Row, 2).Value = "End Reflection Loss - " & ERL_Mode
Call ParameterUnmerge(Selection.Row, SheetType)
    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
        If ERL_Mode = "ASHRAE" Then
        Cells(Selection.Row, 5).Value = "=GetERL_ASHRAE($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$O" & Selection.Row & ")"
        ElseIf ERL_Mode = "NEBB" Then
        Cells(Selection.Row, 5).Value = "=GetERL_NEBB($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$O" & Selection.Row & ")"
        End If
        
    ParamCol1 = 14
    ParamCol2 = 15
'    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
'    Cells(Selection.Row, 5).Value = "=GetERL($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    Else
    ErrorOctOnly
    End If
    
Cells(Selection.Row, ParamCol1) = ERL_Termination
Cells(Selection.Row, ParamCol1).NumberFormat = xlGeneral
Cells(Selection.Row, ParamCol2).Value = ERL_Area
Cells(Selection.Row, ParamCol2).NumberFormat = "0.00""m" & chr(178) & """"
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True

    With Cells(Selection.Row, ParamCol1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Flush,Free"
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

Sub ElbowLoss(SheetType As String)
Dim paramcol As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmElbows.Show

    If btnOkPressed = False Then
    End
    End If

Call ParameterUnmerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Then 'OCT or OCTA
    
    Cells(Selection.Row, 14) = ductW 'public variable
    Cells(Selection.Row, 14).NumberFormat = "##0""mm"""
    Cells(Selection.Row, 15) = elbowLining
    Cells(Selection.Row, 15).NumberFormat = xlGeneral
    'Debug.Print "=GetElbowLoss(" & Cells(6, 5).Address(True, False) & ",$N" & Selection.Row & ",""" & elbowShape & """,$O" & Selection.Row & ",""" & elbowLining & """)"
    Cells(Selection.Row, 5).Value = "=GetElbowLoss(" & Cells(6, 5).Address(True, False) & ",$N" & Selection.Row & ",""" & elbowShape & """,$O" & Selection.Row & ",""" & elbowVanes & """)"
    paramcol = 15
'    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
'    Cells(Selection.Row, 5).Value = "=GetERL($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
    
    With Cells(Selection.Row, paramcol).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="Lined,Unlined"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
    End With
    
Cells(Selection.Row, 2).Value = "Elbow Loss - " & elbowShape
    
End Sub

Sub Silencer(SheetType As String)

Dim CheckRng As Range

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'send to public variable
SolverRow = Selection.Row

'msg = MsgBox("This tool is in beta and may not function as intended.", vbOKOnly, "WARNING!")
frmSilencer.Show

If btnOkPressed = False Then End

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(SolverRow, 5).ClearContents 'clear 31.5Hz octave band
        For Col = 0 To 7 '8 columns
        Cells(SolverRow, 6 + Col).Value = SilencerIL(Col)
        Next Col
        
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(SolverRow, 14).Value = SilLength
    Cells(SolverRow, 14).NumberFormat = "0 ""mm"""
    'add comment with more detail
    Set CheckRng = Cells(SolverRow, 14)
        If CheckRng.Comment Is Nothing Then
        Else
        CheckRng.Comment.Delete
        End If
    CheckRng.AddComment SilSeries & chr(10) & "Length: " & SilLength & "mm" & chr(10) & "Free Area: " & CStr(SilFA) & "%"
    Else
    ErrorOctOnly
    End If
    
Cells(Selection.Row, 2).Value = "Silencer: " & SilencerModel

fmtSilencer (SheetType)

End Sub


Sub Louvres(SheetType As String)

Dim CheckRng As Range

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmLouvres.Show

Set CheckRng = Cells(Selection.Row, 14)

If btnOkPressed = False Then End

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
        For Col = 0 To 7 '8 columns
        Cells(Selection.Row, 6 + Col).Value = LouvreIL(Col)
        Next Col
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = LouvreLength
    Cells(Selection.Row, 14).NumberFormat = "0 ""mm"""
    'add comment with more detail
    Set CheckRng = Cells(Selection.Row, 14)
        If CheckRng.Comment Is Nothing Then
        Else
        CheckRng.Comment.Delete
        End If
    CheckRng.AddComment LouvreSeries & chr(10) & "Length: " & LouvreLength & "mm" & chr(10) & "Free Area: " & LouvreFA
    CheckRng.Comment.Shape.Width = 150
    Else
    ErrorOctOnly
    End If
  
Cells(Selection.Row, 2).Value = "Acoustic Louvres: " & LouvreModel

'apply style - silencer
fmtSilencer (SheetType)

End Sub

Sub Plenum(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmPlenum.Show

If btnOkPressed = False Then End

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetASHRAEPlenumLoss(E$6," & PlenumL & "," & PlenumW & "," & PlenumH & "," & DuctInL & "," & DuctInW & "," & DuctOutL & "," & DuctOutW & "," & PlenumQ & "," & r_h & "," & r_v & ",""" & PlenumLiningType & """,""" & UnlinedType & """,""" & PlenumWallEffect & """," & PlenumElbowEffect & ")"
    ElseIf Left(SheetType, 2) = "TO" Then 'one thirds mode
    Cells(Selection.Row, 5).Value = "=GetASHRAEPlenumLoss(E$6," & PlenumL & "," & PlenumW & "," & PlenumH & "," & DuctInL & "," & DuctInW & "," & DuctOutL & "," & DuctOutW & "," & PlenumQ & "," & r_h & "," & r_v & ",""" & PlenumLiningType & """,""" & UnlinedType & """,""" & PlenumWallEffect & """," & PlenumElbowEffect & ",TRUE)"
    End If

ExtendFunction (SheetType)

Cells(Selection.Row, 2).Value = "Plenum loss"

End Sub


Sub DuctBreakout(SheetType As String)

CheckRow (Selection.Row)

frmBreakOut.Show

    If btnOkPressed = False Then End

Call ParameterUnmerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    
    Cells(Selection.Row, 14).Value = ductL 'global variable
    Cells(Selection.Row, 14).NumberFormat = "0""m"""
    Cells(Selection.Row, 15).Value = DuctWallThickness 'global variable
    Cells(Selection.Row, 15).NumberFormat = "0.0""mm"""
    fmtUserInput SheetType, True
    
    Cells(Selection.Row, 5).Value = "=GetDuctBreakout(E$6," & ductW & "," & ductH & ",$N" & Selection.Row & "," & MaterialDensity & ",$O" & Selection.Row & ")"
    
    ExtendFunction (SheetType)
    
    Else
    
    ErrorOctOnly
    
    End If
Cells(Selection.Row, 2).Value = "Duct breakout"



End Sub

Sub RoomLoss(SheetType As String)
Dim SplitStr() As String
Dim paramcol As Integer
On Error GoTo errorCatch:
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'populate frmRoomLoss
SplitStr = Split(Cells(Selection.Row, 5).Formula, ",", Len(Cells(Selection.Row, 5).Formula), vbTextCompare)
roomL = CLng(SplitStr(1))
roomW = CLng(SplitStr(2))
roomH = CLng(SplitStr(3))
roomType = Cells(Selection.Row, 14).Value
Call frmRoomLoss.Populate_frmRoomLoss

errorCatch:

frmRoomLoss.Show

Call ParameterMerge(Selection.Row, SheetType)

    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetRoomLoss(" & Cells(6, 5).Address(True, False) & "," & roomL & "," & roomW & "," & roomH & ",$N" & Selection.Row & ")"
    Cells(Selection.Row, 14) = roomType
    paramcol = 14
    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True

With Cells(Selection.Row, paramcol).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="Dead, Av. Dead, Average, Av. Live, Live"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

Cells(Selection.Row, 2).Value = "Room Loss"

End Sub

Sub RoomLossRC(SheetType As String)
Dim DefaultArray() As Variant

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

DefaultArray = Array(17, 19, 22, 24, 31, 39, 43) 'Some bullshit Rc, based on a 0.5sec RT

    If Left(SheetType, 3) = "OCT" Or Left(SheetType, 2) = "TO" Then 'OCT, OCTA, TO, or TOA
    Cells(Selection.Row + 1, 5).Value = "=10*LOG(4/E" & Selection.Row & ")" 'next row down
    Else
    ErrorOctOnly
    End If

fmtUserInput SheetType
Cells(Selection.Row, 2).Value = "Room Constant"

'move one row down
Cells(Selection.Row + 1, 5).Select
ExtendFunction (SheetType)

Cells(Selection.Row, 2).Value = "Room Loss - 10LOG(4/Rc)"


    If Left(SheetType, 3) = "OCT" Then 'delete 31.5 and 8k octave bands
    Range(Cells(Selection.Row - 1, 6), Cells(Selection.Row - 1, 12)).Value = DefaultArray
    Cells(Selection.Row, 5).ClearContents
    Cells(Selection.Row, 13).ClearContents
    End If


End Sub



Sub RoomLossRT(SheetType As String)
Dim SplitStr() As String
Dim paramcol As Integer
On Error GoTo errorCatch:
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'populate frmRoomLoss
SplitStr = Split(Cells(Selection.Row, 5).Formula, ",", Len(Cells(Selection.Row, 5).Formula), vbTextCompare)
roomL = CLng(SplitStr(1))
roomW = CLng(SplitStr(2))
roomH = CLng(SplitStr(3))
roomType = Cells(Selection.Row, 14).Value
Call frmRoomLoss.Populate_frmRoomLoss

errorCatch:

frmRoomLossRT.Show

Call ParameterMerge(Selection.Row, SheetType)

    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetRoomLossRT(" & Cells(6, 5).Address(True, False) & "," & roomL & "," & roomW & "," & roomH & ",$N" & Selection.Row & ")"
    Cells(Selection.Row, 14) = roomType
    paramcol = 14
    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True

With Cells(Selection.Row, paramcol).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="<0.2 sec,0.2 to 0.5 sec,0.5 to 1 sec,1.5 to 2 sec,>2 sec"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

Cells(Selection.Row, 2).Value = "Room Loss - RT"

End Sub



Sub RegenNoise(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmRegenNoise.Show
    
    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=GetRegenNoise(E$6,$N" & Selection.Row & ",$O" & Selection.Row & ",""" & regenNoiseElement & """)"
    ParamCol1 = 14
    ParamCol2 = 15
    ElseIf Left(SheetType, 2) = "TO" Then
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
fmtUserInput SheetType, True
Call ParameterUnmerge(Selection.Row, SheetType)

    Select Case regenNoiseElement
    Case Is = "Elbow"
    Cells(Selection.Row, ParamCol1) = "Vanes"
    Cells(Selection.Row, ParamCol2) = "15"
    Case Is = "Transition"
    Cells(Selection.Row, ParamCol1) = "Gradual"
    Cells(Selection.Row, ParamCol2) = "15"
    Case Is = "Damper"
    Cells(Selection.Row, ParamCol1) = ""
    Cells(Selection.Row, ParamCol2) = "11"
    End Select
Cells(Selection.Row, ParamCol1).NumberFormat = "General"
Cells(Selection.Row, ParamCol2).NumberFormat = "0""m/s"""

With Cells(Selection.Row, ParamCol1).Validation
    .Delete
        Select Case regenNoiseElement
        Case Is = "Elbow"
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Vanes, No Vanes"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        Case Is = "Transition"
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Abrupt,Gradual"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        Case Is = "Damper"
        'do nothing
        End Select
End With

With Cells(Selection.Row, ParamCol2).Validation
    .Delete
        Select Case regenNoiseElement
        Case Is = "Elbow"
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="10,15,17.5,20,25,30"
        Case Is = "Transition"
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="7.5,10,15,20"
        Case Is = "Damper"
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="3.5,5.5,8.75,11,14.5"
        End Select
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With

Cells(Selection.Row, 2).Value = "Regen. noise -" & regenNoiseElement

End Sub

Sub DirRevSum(SheetType As String)
Dim SpareRow As Integer
Dim SpareCol As Integer
Dim isSpace As Boolean
Dim StartRw As Integer
Dim EndRw As Integer
Dim ScanCol As Integer
Dim TopOfSheet As Boolean

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'code requires 3 free rows
isSpace = True
    If Left(SheetType, 3) = "OCT" Then
        For SpareRow = Selection.Row To Selection.Row + 2
            For SpareCol = 5 To 13
                If Cells(SpareRow, SpareCol).Value <> "" Then 'column D
                isSpace = False
                End If
            Next SpareCol
        Next SpareRow
    ElseIf Left(SheetType, 2) = "TO" Then
        For SpareRow = Selection.Row To Selection.Row + 2
            For SpareCol = 5 To 25
                If Cells(SpareRow, SpareCol).Value <> "" Then 'column D
                isSpace = False
                End If
            Next SpareCol
        Next SpareRow
    Else
    SheetTypeUnknownError
    End If
    
    
    If isSpace = False Then
    msg = MsgBox("Not enough space", vbOKOnly, "SQUISH!")
    End
    End If

'find sum range
StartRw = Selection.Row - 1 'one above StartRw
ScanCol = Selection.Column
    While Cells(StartRw, ScanCol).Value <> ""
    StartRw = StartRw - 1
        If StartRw < 7 Then
        TopOfSheet = True
        'msg = MsgBox("AutoSum Error", vbOKOnly, "ERROR")
        'End
        End If
    Wend
    
If TopOfSheet = True Then StartRw = 7

EndRw = Selection.Row - 1 'for reveberant sum

'<----------------------------------------------------------------------------------------TODO: Show form and let the user see the range to be summed

'distance correction
Distance (SheetType)
Cells(Selection.Row, 14).Value = 1  'COL N ; 1m by default
'move down
Cells(Selection.Row + 1, Selection.Column).Select
    'Sum direct
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & StartRw + 1 & ":E" & Selection.Row - 1 & ")"
    Else
    ErrorOctOnly
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "Direct component"
fmtSubtotal (SheetType)

'move cursor
Cells(Selection.Row + 1, Selection.Column).Select

'Room loss
RoomLoss (SheetType)

'move down
Cells(Selection.Row + 1, Selection.Column).Select

'number of reverberant sources
TenLogN (SheetType)
Cells(Selection.Row, 14).Value = 1 'default to 1 source ie 10log(n)=0

'move down
Cells(Selection.Row + 1, Selection.Column).Select


    'Sum reverb
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & StartRw + 1 & ":E" & EndRw & ",E" & Selection.Row - 2 & ":E" & Selection.Row - 1 & ")" '=SUM(H19:H19,H22:H23)
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "Reverberant component"
fmtSubtotal (SheetType)

'move down
Cells(Selection.Row + 1, Selection.Column).Select

    'Sum TOTAL
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SPLSUM(E" & Selection.Row - 1 & ",E" & Selection.Row - 4 & ")" 'extra line for number of sources correction
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "TOTAL"
fmtTotal (SheetType)

'Colour highlight
Range(Cells(Selection.Row, 2), Cells(Selection.Row, 18)).Font.Color = RGB(68, 114, 196)

End Sub
