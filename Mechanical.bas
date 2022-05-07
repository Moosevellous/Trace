Attribute VB_Name = "Mechanical"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================

'ducts
Public ductL As Single 'duct length from form in mm
Public ductW As Single 'duct width from form in mm
Public ductH As Single 'duct height from form in mm
Public ductShape As String 'Rectangular or Circular from form
Public ductMethod As String 'ASHRAE or Reynolds, set in form
Public ductLiningThickness As Single 'internal insulation thickness in mm

'duct splits
Public ductA1 As Double 'mm
Public ductA2 As Double 'mm
Public ductSplitType As String

'bend/elbows
Public elbowLining As String
Public elbowShape As String
Public ElbowVanes As String

'End Reflection Loss
Public ERL_Area As Single
Public ERL_Mode As String
Public ERL_Termination As String
Public ERL_Circular As Boolean

'Splitter-silencers
Public SilencerModel As String
Public SilencerIL() As Double
Public SilLength As Double
Public SilFA As Double 'percentage free area, between 0 and 100
Public SilSeries As String
Public SolverRow As Integer

'Acoustic louvres
Public LouvreModel As String
Public LouvreIL() As Double
Public LouvreLength As Double
Public LouvreFA As String 'percentage free area, between 0 and 100
Public LouvreSeries As String

'duct break-out / break-in
Public MaterialDensity As Single 'kg/m3
Public DuctWallThickness As Single 'mm

'plenums....Plena?
Public PlenumL As Long 'mm
Public PlenumW As Long 'mm
Public PlenumH As Long 'mm
Public DuctInL As Single 'mm
Public DuctInW As Single 'mm
Public DuctOutL As Single 'mm
Public DuctOutW As Single 'mm
Public PlenumQ As Integer
Public R_H As Long 'mm
Public r_v As Long 'mm
Public PlenumLiningType As String 'Type of absorptive lining in plenum, from ASHRAE
Public UnlinedType As String 'Type of non-absorptive lining in plenum, from ASHRAE
Public PlenumWallEffectStr As String
Public ApplyPlenumElbowEffect As Boolean
Public PlenumPercentUnlined As Long 'percentage

'Regen (note that regen also uses some variables from ducts group)
Public CalcRegen As Boolean 'set to TRUE to calculate regenerated noise from element
Public PressureLoss As Double 'in Pascalls
Public DamperMultiBlade As Boolean 'set to TRUE for multi-blade
Public RegenMode As String 'NEBB or ASHRAE or other?
Public regenNoiseElement As String 'legacy variable, from frmRegenNoiseASHRAE
Public ElementH As Double 'Silencer height in mm
Public ElementW As Double 'Silencer width in mm
Public BendH As Double 'bend height in mm
Public BendW As Double 'bend width in mm
Public BendCordLength As Double 'bend chord length in mm
Public FlowRate As Double 'air flow rate in m3/s
Public FlowUnitsM3ps As Boolean 'set to TRUE for m3/s, otherwise it's L/s
Public PFA As Double 'percentage free area, between 0 and 100
Public numModules As Integer
Public DuctVelocity As Double 'speed of air in duct in m2/s
Public ElbowHasVanes As Boolean 'set to true if the elbow has vanes
Public ElbowNumVanes As Integer 'number of vanes in an elbow
Public ElbowRadius As Double 'Radius of elbow for regen
Public IncludeTurbulence As Boolean 'set to TRUE for extra juice
Public MainDuctCircular As Boolean 'set to TRUE to calculate areas of circular ducts
Public BranchDuctCircular As Boolean 'set to TRUE to calculate areas of circular branches

''==============================================================================
'' Name:     GetASHRAE
'' Author:   PS
'' Desc:     Legacy duct attenuation function was ASHRAE, renamed at some
''           point
'' Args:     freq - frequency band centre frequency
''           H - duct height in mm
''           W - duct width in mm
''           DuctType - R or C for Rectangular and Circular
''           Length - Duct length in metres
'' Comments: (1) here for legacy reasons, keep the old function and forward on
''               to the new function
''           (2) Do we still need it? Maybe one day we'll ditch it
''==============================================================================
'Function GetASHRAE(freq As String, H As Long, W As Long, DuctType As String, _
'Length As Double)
'GetASHRAE = GetASHRAEDuct(freq, H, W, DuctType, Length)
'End Function

'==============================================================================
' Name:     DuctAtten_ASHRAE
' Author:   PS
' Desc:     Looks up down duct attenuation from ASHRAE table and matches to
'           input dimensions
' Args:     freq - octave band centre frequency in Hz
'           H - duct height in mm
'           W - duct width in mm
'           DuctType - R or C for Rectangular and Circular, then 0, 25 or 50
'           for lining thickness
'           Length - Duct length in metres
' Comments: (1) Values in text file are sorted by cross sectional area, so the
'           input area is matched to the first
'           (2) Changed ASHRAE_DUCTS.txt file, now has consistent columns
'==============================================================================
Function DuctAtten_ASHRAE(freq As String, H As Long, W As Long, DuctType As String, _
Length As Double)

Dim ReadStr() As String 'holds lines of data from text file as strings
Dim i As Integer '<-line number for reading text file
Dim SplitStr() As String 'holds each value data from text file as strings
Dim SplitVal() As Double 'holds data converted into numbers
Dim CurrentType As String 'same rules as DuctType, set at lines which contain '*'
Dim InputArea As Double 'cross sectional area of duct in mm^2
Dim ReadArea As Double 'cross sectional area of duct from file in mm^2
Dim found As Boolean 'trigger to escape the read loop
Dim col As Integer 'index for each column
Dim bandNo As Integer 'which column is the frequency band in
Dim f As Double 'holds number

'Get Array from text
Close #1

Call GetSettings

Open ASHRAE_DUCT For Input As #1  'public variable points to te file

i = 0
found = False

f = freqStr2Num(freq) 'convert to number
fStr = CStr(f)

    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) = "*" Then
        TXT_HEAD = Replace(ReadStr(i), vbTab, " ")
        'set CurrentType
        CurrentType = Right(SplitStr(0), Len(SplitStr(0)) - 1)
        'Debug.Print "TYPE: " & currentType
        
        bandNo = -1 'for catching errors
            'find matching column header
            For j = LBound(SplitStr) To UBound(SplitStr)
                'look for matching column headers
                If fStr = SplitStr(j) Then bandNo = j
            Next j
        
        Else
            
            'convert SplitString from Strings to Values
            For col = 0 To UBound(SplitStr)
                If SplitStr(col) <> "" And IsNumeric(SplitStr(col)) Then
                ReDim Preserve SplitVal(col)
                SplitVal(col) = CDbl(SplitStr(col))
                End If
            Next col
            
            'resize array
            ReDim Preserve SplitVal(col + 1)
            
                If Right(DuctType, 1) = "R" Then 'RECTANGULAR DUCT
                ReadArea = SplitVal(0) * SplitVal(1)
                InputArea = H * W
                ElseIf Right(DuctType, 1) = "C" Then 'CIRCULAR DUCT
                ReadArea = WorksheetFunction.Pi * ((SplitVal(0) / 2) ^ 2)
                InputArea = WorksheetFunction.Pi * ((H / 2) ^ 2)
                Else 'error in text file, skip!
                End If
            
            If InputArea <= ReadArea And CurrentType = DuctType Then
            'set public variable
            TXT_RAW = Replace(ReadStr(i), vbTab, " ")
            
                If bandNo < 0 Then
                DuctAtten_ASHRAE = "-"
                Exit Function
                End If
                
            'read value and multiply by length
            DuctAtten_ASHRAE = SplitVal(bandNo) * -Length
                
                'Floor the value, duct attenuation shouldn't be above 40dB
                If DuctAtten_ASHRAE < -40 Then
                DuctAtten_ASHRAE = -40
                End If
                
            found = True '<-this will end the loop
            End If
            
        End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function


'==============================================================================
' Name:     DuctAtten_Reynolds
' Author:   PS
' Desc:     Calculates down-duct attenuation from input parameters
' Args:     fStr - octave band centre frequency in Hz
'           H - duct height in mm
'           W - duct width in mm
'           thickness - internal insulation thickness, in mm
'           L - Duct length in metres
' Comments: (1) For rectangular ducts only!
'==============================================================================
Function DuctAtten_Reynolds(fStr As String, H As Double, W As Double, _
thickness As Double, L As Double)

Dim PonA As Double 'Perimeter divided by area
Dim Attn As Double 'Attenuation in dB, up to 250Hz
Dim IL As Double 'Insertion Loss in dB, a part of Attn above 250Hz
Dim f As Double 'frequency as value
Dim a As Double 'cross sectional area in m^2
Dim P As Double 'perimeter in metres

'Static Values from NEBB book for various coefficients
'bands:    63      125     250    500     1k    2k     4k     8k
b = Array(0.0133, 0.0574, 0.271, 1.0147, 1.77, 1.392, 1.518, 1.581)
c = Array(1.959, 1.41, 0.824, 0.5, 0.695, 0.802, 0.451, 0.219)
D = Array(0.917, 0.941, 1.079, 1.087, 0, 0, 0, 0)

'convert to millimetres to metres
H = H / 1000
W = W / 1000

P = (H * 2) + (W * 2) 'perimeter
a = H * W 'area
    If P = 0 Or a = 0 Then
    PonA = 0
    Else
    PonA = P / a
    End If

f = freqStr2Num(fStr)
i = GetArrayIndex_OCT(fStr)

    'catch error: frequency bands not defined
    If i = 999 Or i = -1 Then
    DuctAtten_Reynolds = "-"
    Else
    
        'equation 5.16
        If thickness = 0 Then 'don't apply this correction
        IL = 0
        Else
        IL = (3.281 * b(i)) * ((0.305 * PonA) ^ c(i)) * _
            ((0.039 * thickness) ^ D(i)) * L
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
        DuctAtten_Reynolds = -40
        Else
        DuctAtten_Reynolds = Round((IL + Attn) * -1, 1)
        End If
    
    End If

End Function

'==============================================================================
' Name:     DuctAttenCircular_Reynolds
' Author:   PS
' Desc:     Calculates down-duct attenuation from input parameters
' Args:     fStr - octave band centre frequency in Hz
'           dia - duct diameter in mm
'           thickness - internal insulation thickness, in mm
'           L - Duct length in metres
' Comments: (1) Beware of versions of the table with incorrect values!
'           (2) Updated for unliend ducts, there's a different table for that!
'==============================================================================
Function DuctAttenCircular_Reynolds(fStr As String, dia As Double, _
thickness As Double, L As Double)

'declare Arrays from Table 5.7 of NEBB book
D1 = Array(0.1, 0.1, 0.16, 0.16, 0.33, 0.33, 0.33)
D2 = Array(0.1, 0.1, 0.1, 0.16, 0.23, 0.23, 0.23)
D3 = Array(0.07, 0.07, 0.07, 0.1, 0.16, 0.16, 0.16)
D4 = Array(0.03, 0.03, 0.03, 0.07, 0.07, 0.07, 0.07)

'declare Arrays from Table 5.8 of NEBB book
'bands     63      125     250     500     1k     2k    4k   8k
a = Array(0.2825, 0.5237, 0.3652, 0.1333, 1.933, 2.73, 2.8, 1.545)
b = Array(0.3447, 0.2234, 0.79, 1.845, 0, 0, 0, 0)
c = Array(-0.05251, -0.004936, -0.1157, -0.3735, 0, 0, 0, 0)
D = Array(-0.03837, -0.02724, -0.01834, -0.01293, 0.06135, -0.07341, -0.1467, -0.05452)
e = Array(0.00091315, 0.0003377, -0.0001211, 0.00008624, -0.003891, 0.0004428, 0.003404, 0.00129)
f = Array(-0.000008294, -0.00000249, 0.000002681, -0.000004986, 0.00003934, 0.000001006, -0.00002851, -0.00001318)

i = GetArrayIndex_OCT(fStr) '63Hz is element 0

    If i < 0 Or i > 6 Then
    DuctAttenCircular_Reynolds = "-"
    Exit Function
    End If

    If thickness = 0 Then 'use table 5.7 method
    
        'choose diameter in mm
        Select Case dia
        Case Is <= 180
        DuctAttenCircular_Reynolds = D1(i) * L * -1
        Case Is <= 381
        DuctAttenCircular_Reynolds = D2(i) * L * -1
        Case Is <= 762
        DuctAttenCircular_Reynolds = D3(i) * L * -1
        Case Is <= 1524
        DuctAttenCircular_Reynolds = D4(i) * L * -1
        Case Is > 1524
        DuctAttenCircular_Reynolds = "-"
        End Select
        
        'top out at 40dB
        If DuctAttenCircular_Reynolds < -40 Then
        DuctAttenCircular_Reynolds = -40
        End If
        
    Else 'some thickness - use equation method
    
    'convert thickness and inside diameter to imperial units for some reason
    thickness = 0.039 * thickness
    dia = 0.039 * dia
    
        If i = 999 Or i = -1 Then
        DuctAttenCircular_Reynolds = ""
        Else
    
        DuctAttenCircular_Reynolds = 3.281 * (a(i) + (b(i) * thickness) + _
        (c(i) * thickness ^ 2) + (D(i) * dia) + (e(i) * dia ^ 2) + _
        (f(i) * dia ^ 3)) * L
        
            'top out at 40dB
            If DuctAttenCircular_Reynolds > 40 Then
            DuctAttenCircular_Reynolds = -40
            Else
            DuctAttenCircular_Reynolds = DuctAttenCircular_Reynolds * -1
            End If
            
        End If
    End If
End Function


'==============================================================================
' Name:     DuctBendAtten_SRL
' Author:   PS
' Desc:     Looks up down duct/bend attenuation from SRL table and matches to
'           input dimensions
' Args:     freq - octave band centre frequency in Hz
'           H - duct height in mm
'           W - duct width in mm
'           DuctType - R or C for Rectangular and Circular
'           Length - Duct length in metres
' Comments: (1) Values in text file are sorted by cross sectional area, so the
'           input area is matched to the first
'==============================================================================
Function DuctBendAtten_SRL(fStr As String, DuctWidth As Long, DuctType As String, _
    Optional Length As Double)

Dim ReadStr() As String 'holds lines of data from text file as strings
Dim i As Integer '<-line number for reading text file
Dim SplitStr() As String 'holds each value data from text file as strings
Dim SplitVal() As Double 'holds data converted into numbers
Dim InputArea As Double 'cross sectional area of duct in mm^2
Dim ReadArea As Double 'cross sectional area of duct from file in mm^2
Dim found As Boolean 'trigger to escape the read loop
Dim col As Integer 'index for each column

'Get Array from text
Close #1

Call GetSettings

Open SRL_DUCTS For Input As #1  'public variable points to the file

i = 0
found = False
'freq = freqStr2Num(fStr)
fCol = GetArrayIndex_OCT(fStr, 3) '63Hz is array element 3
    
    'check for optional input
    If Length = 0 Then
    Length = 1
    End If
    
    '63 to 4k only
    If fCol < 3 Or fCol > 9 Then
    DuctBendAtten_SRL = "-"
    Exit Function
    End If
    
    'Main loop!
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) = "*" Then '* is the type identifier
        
        CurrentTable = Mid(SplitStr(0), 2, 12)
        TXT_HEAD = Replace(ReadStr(i), vbTab, " ")
        
        Else
        CurrentType = SplitStr(0) 'first element
        
            If IsNumeric(SplitStr(1)) And IsNumeric(SplitStr(2)) Then
            
                'check type and widths
                If DuctType = CurrentType And _
                    DuctWidth >= SplitStr(1) And DuctWidth < SplitStr(2) Then
                
                TXT_RAW = Replace(ReadStr(i), vbTab, " ")
                
                    If IsNumeric(SplitStr(fCol)) Then
                    DuctBendAtten_SRL = CDbl(SplitStr(fCol)) * -Length
                        'Floor the value, duct attenuation shouldn't be above 40dB
                        If DuctBendAtten_SRL < -40 Then
                        DuctBendAtten_SRL = -40
                        End If
                    Else
                    DuctBendAtten_SRL = "-"
                    End If
                    
                found = True '<-this will end the loop
                End If
                
            End If
        End If
        
    i = i + 1
    Loop

    'if the loop finds nothing the output dash
    If found = False Then
    DuctBendAtten_SRL = "-"
    End If
    
closefile: '<-on errors, closes text file
Close #1
End Function

'==============================================================================
' Name:     DuctBreakOut_NEBB
' Author:   NI
' Desc:     Calculates sound energy breaking out of ducts, according to NEBB
'           method
' Args:     fStr - octave band centre frequency in Hz
'           H and W - Height and Width of duct in mm
'           L - Duct length in m
'           MaterialDensity - for duct wall material, steel, PVC etc in kg/m3
'           DuctWallThickness - thickness of wall in mm
' Comments: (1) NEBB method
'==============================================================================
Function DuctBreakOut_NEBB(fStr As String, H As Single, W As Single, L As Single, _
MaterialDensity As Single, DuctWallThickness As Single)

Dim SurfaceMass As Single 'Surface mass of duct wall in kg/m^2
Dim fL As Long 'Limiting frequency in Hz
Dim f As Double 'frequency in Hz
Dim TLoutMin As Single 'in dB
Dim TLout As Single 'in dB

f = freqStr2Num(fStr)
fL = 613000# / ((W * H) ^ 0.5)

SurfaceMass = MaterialDensity * DuctWallThickness / 1000 'convert to metres

'convert length to millimetres, works out the same!
TLoutMin = 10 * Application.WorksheetFunction.Log10(2 * L * 1000 * ((1 / W) + (1 / H)))

    If SurfaceMass <> 0 And W <> 0 And H <> 0 Then
        If f < fL Then
        TLout = 10 * Application.WorksheetFunction.Log10((f * (SurfaceMass ^ 2)) / _
                (W + H)) + 17 'equation 6.11
        Else
        'equation 6.12
        TLout = 20 * Application.WorksheetFunction.Log10(f * SurfaceMass) - 45
        End If
        
        'TLout can't be greater than 45dB (but why?)
        If TLout > 45 Then TLout = 45
        
        If TLout > TLoutMin Then
        DuctBreakOut_NEBB = TLoutMin - TLout 'comes out as negative
        Else
        DuctBreakOut_NEBB = 0
        End If
    End If
    
End Function

'==============================================================================
' Name:     DuctBreakIn_NEBB
' Author:   NI & PS
' Desc:     Calculates sound breaking into a duct, according to NEBB method
' Args:     fStr - octave band centre frequency in Hz
'           H and W - Height and Width of duct in mm
'           L - Duct length in m
'           MaterialDensity - for duct wall material, steel, PVC etc in kg/m3
'           DuctWallThickness - thickness of wall in mm
' Comments: (1)
'==============================================================================
Function DuctBreakIn_NEBB(fStr As String, H As Single, W As Single, L As Single, _
MaterialDensity As Single, DuctWallThickness As Single)

Dim SurfaceMass As Single 'Surface mass of duct wall in kg/m^2
Dim f As Double 'octave band centre frequency band
Dim F1 As Double 'octave band centre frequency band
Dim a As Single 'Larger duct dimensions in mm
Dim b As Single 'Smaller duct dimensions in mm
Dim TLin_A As Single 'TL_in for larger dimension
Dim TLin_B As Single 'TL_in for smaller dimension
Dim TLin1 As Single 'larger of the two TL_in values
Dim TLout As Single 'used to calculate

f = freqStr2Num(fStr) 'convert to values

    'set A as the larger dimension
    If H > W Then
    a = H
    b = W
    Else
    a = W
    b = H
    End If

F1 = (1.718 * 10 ^ 5) / a

'Method relies on breakout number
'call trace function for breakout, but make it positive
TLout = DuctBreakOut_NEBB(fStr, H, W, L, MaterialDensity, DuctWallThickness) * -1 + _
    10 * Application.WorksheetFunction.Log(2 * (L * 1000) * ((H + W) / (H * W)))

    If F1 > f Then
    'equation 6.15a
    TLin_A = TLout - 4 - (10 * Application.WorksheetFunction.Log(a / b)) + _
        (20 * Application.WorksheetFunction.Log(f / F1))
    'equation 6.15b
    TLin_B = 10 * Application.WorksheetFunction.Log((L * 1000) * ((1 / a) + (1 / b)))
        
        If TLin_A > TLin_B Then
        TLin1 = TLin_A
        Else
        TLin1 = TLin_B
        End If
    
    Else 'f1<=f
    TLin1 = TLout - 3 'equation 6.16
    End If

DuctBreakIn_NEBB = (TLin1 * -1) - 3  'Trace convention is negative
    
End Function


'Function GetDuctLaggingIL(freq As String, H As Single, w As Single, DuctMass As Single, LaggingMass As Single, LaggingThickness As Single)
'Dim P1 As Single
'Dim P2 As Single
'Dim S As Single
'Dim f_res As Single
'Dim IL_LF As Single
'
'f = freqStr2Num(freq)
'P1 = 2 * (w + H)
'P2 = 2 * (L + w + 4 * LaggingThickness)
'S = 2 * LaggingThickness * (w + H + 2 * LaggingThickness)
'f_res = 156 * ((((P2 / P1) + (LaggingMass / DuctMass)) * (P1 * S / LaggingMass)) ^ 0.5)
'
''Low frequency insertion Loss
'IL_LF = 20 * Application.WorksheetFunction.Log(1 + (LaggingMass / DuctMass) * (P1 / P2))
'
'    'check for octave band containing f_res
'    If IsInOctaveBand(f, freq) Then
'    GetDuctLaggingIL = IL_LF - 5
'    Else
'        If f < f_res Then
'        GetDuctLaggingIL = IL_LF
'        Else
'        GetDuctLaggingIL = IL_LF + 29.9 * Application.WorksheetFunction.Log(f / (1.41 * f_res))
'        End If
'    End If
'
'End Function

'==============================================================================
' Name:     FlexDuctAtten_ASHRAE
' Author:   PS
' Desc:     Gets down-duct attenuation for flexible ducts ASHRAE table, stored
'           in text file ASHRAE_FLEX (public variable)
' Args:     freq - octave band centre frequency to match to data
'           dia - duct diameter in mm
'           L - duct length in m
' Comments: (1) locked to values in text file, no interpolation
'==============================================================================
Function FlexDuctAtten_ASHRAE(freq As String, dia As Integer, L As Double)

On Error GoTo closefile

Dim ReadStr() As String 'array for holding daat from text file
Dim i As Integer 'counter for each line
Dim SplitStr() As String 'splits text file into substrings
Dim SplitVal() As Double 'converted to values
Dim col As Integer 'counter for each column
Dim found As Boolean 'switch for when match found

Call GetSettings

Open ASHRAE_FLEX For Input As #1  'public

i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then 'asterisk detotes titles
        
            'convert to values
            For col = 0 To UBound(SplitStr)
                If SplitStr(col) <> "" Then
                ReDim Preserve SplitVal(col)
                SplitVal(col) = CDbl(SplitStr(col))
                End If
            Next col
            
            ReDim Preserve SplitVal(col + 1)
            
                If SplitVal(0) = dia And SplitVal(1) = L Then
                    Select Case freq
                    Case Is = "63"
                    FlexDuctAtten_ASHRAE = -SplitVal(2)
                    Case Is = "125"
                    FlexDuctAtten_ASHRAE = -SplitVal(3)
                    Case Is = "250"
                    FlexDuctAtten_ASHRAE = -SplitVal(4)
                    Case Is = "500"
                    FlexDuctAtten_ASHRAE = -SplitVal(5)
                    Case Is = "1k"
                    FlexDuctAtten_ASHRAE = -SplitVal(6)
                    Case Is = "2k"
                    FlexDuctAtten_ASHRAE = -SplitVal(7)
                    Case Is = "4k"
                    FlexDuctAtten_ASHRAE = -SplitVal(8)
                    Case Is = "1000"
                    FlexDuctAtten_ASHRAE = -SplitVal(6)
                    Case Is = "2000"
                    FlexDuctAtten_ASHRAE = -SplitVal(7)
                    Case Is = "4000"
                    FlexDuctAtten_ASHRAE = -SplitVal(8)
                    Case Else
                    FlexDuctAtten_ASHRAE = ""
                    End Select
                End If
        End If
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function

'==============================================================================
' Name:     GetERL
' Author:   PS
' Desc:     Legacy ERL function was ASHRAE, now renamed
' Args:     TerminationType - "Flush" or "Free"
'           freq - octave band centre frequency
'           DuctArea - Cross sectional area of duct in m^2
' Comments: (1) here for legacy reasons, keep the old function and forward on
'               to the new function
'           (2) Do we still need it? Maybe one day we'll ditch it
''==============================================================================
'Function GetERL(TerminationType As String, freq As String, DuctArea As Double)
'GetERL = GetERL_ASHRAE(TerminationType, freq, DuctArea)
'End Function

'==============================================================================
' Name:     ERL_ASHRAE
' Author:   PS
' Desc:     Calculates End Reflection Loss using the ASHRAE method
' Args:     TerminationType - Flush or Free (string)
'           fStr - octave band centre frequency
'           DuctArea - Cross sectional area of duct, in m^2
' Comments: (1)
'==============================================================================
Function ERL_ASHRAE(TerminationType As String, fStr As String, DuctArea As Double)

Dim dia As Double 'duct diameter in mm
Dim A1 As Double 'variable in ASHRAE method (dimensionless)
Dim A2 As Double 'variable in ASHRAE method (dimensionless)
Dim f As Double 'frequency as value
Dim c0 As Double 'speed of sound

    If DuctArea <> 0 Then
    'eqn 11 of ASHRAE - same for rectangles and circles!
    dia = (4 * DuctArea / Application.WorksheetFunction.Pi) ^ 0.5
    
    f = freqStr2Num(fStr) 'convert to a value
    c0 = 343
    
        'table 28 of ASHRAE
        If TerminationType = "Flush" Then
        A1 = 0.7
        A2 = 2
        ElseIf TerminationType = "Free" Then
        A1 = 1
        A2 = 2
        End If
        
    ERL_ASHRAE = -10 * Application.WorksheetFunction.Log10(1 + ((A1 * c0) / _
        (f * dia * Application.WorksheetFunction.Pi)) ^ A2)
    Else
    ERL_ASHRAE = 0
    End If
    
End Function

'==============================================================================
' Name:     ERL_NEBB
' Author:   PS
' Desc:     Calculates End Reflection Loss using the ASHRAE method
' Args:     TerminationType - Flush or Free (string)
'           fStr - octave band centre frequency
'           DuctArea - Cross sectional area of duct, in m^2
' Comments: (1)
'==============================================================================
Function ERL_NEBB(TerminationType As String, fStr As String, DuctArea As Double)
Dim dia As Double 'duct diameter in mm
Dim A1 As Double 'variable in ASHRAE method (dimensionless)
Dim A2 As Double 'variable in ASHRAE method (dimensionless)
Dim f As Double 'frequency as value
Dim c0 As Double 'speed of sound
    
    If DuctArea <> 0 Then
    'eqn 5.40 of NEBB method
    dia = (4 * DuctArea / Application.WorksheetFunction.Pi) ^ 0.5
    
    f = freqStr2Num(fStr)
    c0 = 343
    
        If TerminationType = "Flush" Then
        A1 = 0.8
        A2 = 1.88
        ElseIf TerminationType = "Free" Then
        A1 = 1
        A2 = 1.88
        End If
    'Eqn 5.38 or 5.39 of NEBB method, using A1 and A2 to switch between
    ERL_NEBB = -10 * Application.WorksheetFunction.Log10(1 + ((A1 * c0) / _
        (f * dia * Application.WorksheetFunction.Pi)) ^ A2)
    Else
    ERL_NEBB = 0
    End If
    
End Function

''==============================================================================
'' Name:     GetRegenNoise
'' Author:   PS
'' Desc:     Legacy Regen Function
'' Args:     freq - octave band centre frequency in Hz
''           Element - Transition, Elbows or Dampers as string
''           Condition - Vanes/Not, Gradual/Abrupt as string
''           Velocity - Air speed in m/s
'' Comments: (1) here for legacy reasons, keep the old function and forward on
''               to the new function
''           (2) Do we still need it? Maybe one day we'll ditch it
''==============================================================================
'Function GetRegenNoise(freq As String, Condition As String, Velocity As Double, _
'Element As String)
'GetRegenNoise = GetRegenNoise_ASHRAE(freq, Element, Condition, Velocity)
'End Function


'==============================================================================
' Name:     RegenNoise_ASHRAE
' Author:   PS
' Desc:     Returns the values from regenerated noise graphs from ASHRAE
' Args:     freq - octave band centre frequency in Hz
'           ElementType - Transition, Elbows or Dampers as string
'           Condition - Vanes/Not, Gradual/Abrupt as string
'           Velocity - Air speed in m/s
' Comments: (1)
'==============================================================================
Function RegenNoise_ASHRAE(freq As String, ElementType As String, _
Condition As String, Velocity As Double)

On Error GoTo closefile

Dim ReadStr() As String 'string for holding data from text file
Dim SplitStr() As String 'Array for holding sub strings from text file
Dim SplitVal() As Double 'converted to values
Dim col As Integer 'counter for each column of data
Dim found As Boolean 'switch for matching data
Dim i As Integer 'counter for each line of text file

Call GetSettings

Open ASHRAE_REGEN For Input As #1  'public variable

    i = 0
    found = False
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then
            
            If CurrentType = ElementType Then
            'match condition and velocity to referenced values
                If SplitStr(0) = Condition And CDbl(SplitStr(1)) = Velocity Then
                
                    'convert to values
                    For col = 1 To UBound(SplitStr)
                        If SplitStr(col) <> "" Then
                        ReDim Preserve SplitVal(col)
                        SplitVal(col) = CDbl(SplitStr(col))
                        End If
                    Next col
                
                    Select Case freq
                    Case Is = "63"
                    RegenNoise_ASHRAE = SplitVal(2)
                    Case Is = "125"
                    RegenNoise_ASHRAE = SplitVal(3)
                    Case Is = "250"
                    RegenNoise_ASHRAE = SplitVal(4)
                    Case Is = "500"
                    RegenNoise_ASHRAE = SplitVal(5)
                    Case Is = "1k"
                    RegenNoise_ASHRAE = SplitVal(6)
                    Case Is = "2k"
                    RegenNoise_ASHRAE = SplitVal(7)
                    Case Is = "4k"
                    RegenNoise_ASHRAE = SplitVal(8)
                    Case Is = "1000"
                    RegenNoise_ASHRAE = SplitVal(6)
                    Case Is = "2000"
                    RegenNoise_ASHRAE = SplitVal(7)
                    Case Is = "4000"
                    RegenNoise_ASHRAE = SplitVal(8)
                    Case Else
                    RegenNoise_ASHRAE = "-"
                    End Select
                    
                found = True
                    
                End If
                
            End If
            
        Else '* is the type identifier
        CurrentType = Right(SplitStr(0), Len(SplitStr(0)) - 1)
        End If
        
            'catch for 0 or blank strings
            If RegenNoise_ASHRAE = 0 Or RegenNoise_ASHRAE = "" Then
            RegenNoise_ASHRAE = "-"
            End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1

End Function

''==============================================================================
'' Name:     GetElbowLoss
'' Author:   PS
'' Desc:     Legacy function, forwards to ASHRAE method
'' Args:     fStr - octave band centre frequency in Hz
''           W - duct width in mm
''           elbowShape - Square or Radius
''           ductLining - Lined or Unlined
''           VaneType - Vanes or No Vanes
'' Comments: (1)
''==============================================================================
'Function GetElbowLoss(fStr As String, W As Double, elbowShape As String, _
'ductLining As String, VaneType As String)
'GetElbowLoss = GetElbowLossASHRAE(fStr, W, elbowShape, ductLining, VaneType)
'End Function

'==============================================================================
' Name:     ElbowLoss_ASHRAE
' Author:   PS
' Desc:     Calculates loss through an elbow according to the ASHRAE method
' Args:     fStr - octave band centre frequency in Hz
'           W - duct width in mm
'           elbowShape - Square or Radius
'           ductLining - Lined or Unlined
'           VaneType - Vanes or No Vanes
' Comments: (1)
'==============================================================================
Function ElbowLoss_ASHRAE(fStr As String, W As Double, elbowShape As String, _
ductLining As String, VaneType As String)

Dim Unlined() As Variant 'values from ASHRAE table
Dim Lined() As Variant 'values from ASHRAE table
Dim RadiusBend() As Variant 'values from ASHRAE table
Dim freq As Double 'octave band centre frequency
Dim FW As Double 'frequency (kHz) X width (mm)
Dim i As Integer 'array index
Dim linedDuct As Boolean 'true if lined
Dim Vanes As Boolean 'true if has vanes

    If ductLining = "Lined" Then
    linedDuct = True
    ElseIf ductLining = "Unlined" Then
    linedDuct = False
    Else
    linedDuct = "-"
    End If
    
    If VaneType = "Vanes" Then
    Vanes = True
    ElseIf VaneType = "No Vanes" Then
    Vanes = False
    Else
    Vanes = "-"
    End If
    
'from table 22 of ASHRAE
Unlined = Array(0, -1, -5, -8, -4, -3)
Lined = Array(0, -1, -6, -11, -10, -10)
'from table 24 of ASHRAE
UnlinedV = Array(0, -1, -4, -6, -4)
LinedV = Array(0, -1, -4, -7, -7)
'table 23 of ASHRAE
RadiusBend = Array(0, -1, -2, -3)

freq = freqStr2Num(fStr)
FW = (freq / 1000) * W

    Select Case elbowShape
    Case Is = "Square"
        If Vanes = False Then
            Select Case FW
            Case Is < 48
            i = 0
            Case Is < 96
            i = 1
            Case Is < 190
            i = 2
            Case Is < 380
            i = 3
            Case Is < 760
            i = 4
            Case Is >= 760
            i = 5
            End Select
            
                If linedDuct = True Then
                ElbowLoss_ASHRAE = Lined(i)
                Else 'LinedDuct = False
                ElbowLoss_ASHRAE = Unlined(i)
                End If
                
        Else 'vanes=true
            Select Case FW
            Case Is < 48
            i = 0
            Case Is < 96
            i = 1
            Case Is < 190
            i = 2
            Case Is < 380
            i = 3
            Case Is >= 380
            i = 4
            End Select
            
                If linedDuct = True Then
                ElbowLoss_ASHRAE = LinedV(i)
                Else 'LinedDuct = False
                ElbowLoss_ASHRAE = UnlinedV(i)
                End If
            
        End If
        
    Case Is = "Radius"
        Select Case FW
        Case Is < 48
        i = 0
        Case Is < 96
        i = 1
        Case Is < 190
        i = 2
        Case Is >= 190
        i = 3
        End Select
        
    ElbowLoss_ASHRAE = RadiusBend(i)
            
    End Select

End Function

'==============================================================================
' Name:     ElbowLoss_NEBB
' Author:   PS
' Desc:     Calculates loss through an elbow according to the NEBB method
'           (Bodley), for circular radiused bends
' Args:     fStr - octave band centre frequency in Hz
'           W - duct width in mm
'           elbowShape - Square or Radius
'           ductLiningThickness - insulation thickness in mm
' Comments: (1)
'==============================================================================
Function ElbowLoss_NEBB(fStr As String, dia As Double, ductLiningThickness As Integer)
Dim IL_DonRsquared As Double
Dim IL As Double
Dim R As Double

f = freqStr2Num(fStr)
f = 0.039 * f 'imperial unit correction

    Select Case dia
        Case Is < 152 'mm
        IL_DonRsquared = "-"
        Case Is < 457 'mm
        'Equation 5.20
        IL_DonRsquared = 0.485 + _
            2.094 * Application.WorksheetFunction.Log10(f * dia) + _
            3.172 * Application.WorksheetFunction.Log10((f * dia) ^ 2) - _
            1.578 * Application.WorksheetFunction.Log10((f * dia) ^ 4) + _
            0.085 * Application.WorksheetFunction.Log10((f * dia) ^ 7)
        Case Is < 1981 'mm
        'Equation 5.21
        IL_DonRsquared = -1.493 + (0.021 * ductLiningThickness) + _
            1.406 * Application.WorksheetFunction.Log10(f * dia) + _
            2.779 * Application.WorksheetFunction.Log10((f * dia) ^ 2) - _
            0.662 * Application.WorksheetFunction.Log10((f * dia) ^ 4) + _
            0.016 * Application.WorksheetFunction.Log10((f * dia) ^ 7)
        Case Is > 1981
        IL_DonRsquared = "-"
    End Select
    
R = dia / 2
IL = IL_DonRsquared / ((dia / R) ^ 2)

    If IL < 0 Then 'just make it 0
    ElbowLoss_NEBB = 0
    Else 'Trace convention is negative
    ElbowLoss_NEBB = IL * -1
    End If
    
End Function

'==============================================================================
' Name:     PlenumLoss_ASHRAE
' Author:   PS
' Desc:     Calculates Plenum loss according to the ASHRAE method.
'           Hold on, it's gonna get tricky
' Args:     fStr - Octave band centre frequency Hz
'           L - length of plenum in mm
'           W - width of plenum in mm
'           H - Height of plenum in mm
'           DuctInL - Length of inlet in mm
'           DuctInW - Width of inlet in mm
'           DuctOutL - Length of outlet in mm
'           DuctOutW - Width of outlet in mm
'           Q - 2 or 4 for type of spherical spreading
'           r_h - hoizontal offset in mm
'           r_v - vertical offset in mm
'           PlenumLiningType - Concrete, sheet metal or lined sheet metal duct
'           UnlinedType - The unlined plenum material (behind lining)
'           WallEffect - Different types for wall effect options in ASHRAE
'           applyElbowEffect - TRUE if tickbox is checked
'           OneThirdsMode - Optional switch for single one-third band
' Comments: (1)
'==============================================================================
Function PlenumLoss_ASHRAE(fStr As String, L As Long, W As Long, H As Long, _
    DuctInL As Single, DuctInW As Single, DuctOutL As Single, DuctOutW As Single, _
    Q As Integer, R_H As Long, r_v As Long, PlenumLiningType As String, _
    UnlinedType As String, WallEffect As String, applyElbowEffect As Boolean, _
    UnlinedPercent As Long, Optional OneThirdsMode As Boolean)

Dim f_OneUp As Double
Dim f_OneDown As Double
Dim f As Double
Dim Loss1 As Double
Dim Loss2 As Double
Dim Loss3 As Double

f = freqStr2Num(fStr)
    
    'put bounds on the function
    If f < 50 Or f > 4000 Then
    PlenumLoss_ASHRAE = "-"
    Exit Function
    End If

If IsMissing(OneThirdsMode) Then OneThirdsMode = False

    'Switches between a single one-third octave band and a full octave band
    If OneThirdsMode = True Then
    PlenumLoss_ASHRAE = PlenumLossOneThirdOctave_ASHRAE(f, L, W, H, DuctInL, _
        DuctInW, DuctOutL, DuctOutW, Q, R_H, r_v, PlenumLiningType, UnlinedType, _
        WallEffect, applyElbowEffect, UnlinedPercent)
    Else 'octave bands!
    f_OneUp = GetAdjacentFrequency(f, "Up")
    f_OneDown = GetAdjacentFrequency(f, "Down")
    'get for each one third octave and then Tl average them
    Loss1 = PlenumLossOneThirdOctave_ASHRAE(f_OneDown, L, W, H, DuctInL, _
        DuctInW, DuctOutL, DuctOutW, Q, R_H, r_v, PlenumLiningType, UnlinedType, _
        WallEffect, applyElbowEffect, UnlinedPercent)
    Loss2 = PlenumLossOneThirdOctave_ASHRAE(f, L, W, H, DuctInL, _
        DuctInW, DuctOutL, DuctOutW, Q, R_H, r_v, PlenumLiningType, UnlinedType, _
        WallEffect, applyElbowEffect, UnlinedPercent)
    Loss3 = PlenumLossOneThirdOctave_ASHRAE(f_OneUp, L, W, H, DuctInL, _
        DuctInW, DuctOutL, DuctOutW, Q, R_H, r_v, PlenumLiningType, UnlinedType, _
        WallEffect, applyElbowEffect, UnlinedPercent)
    'Note: losses are negative already so no need for negatives sign in formula
    PlenumLoss_ASHRAE = 10 * Application.WorksheetFunction.Log10((1 / 3) * _
        ((10 ^ (Loss1 / 10)) + (10 ^ (Loss2 / 10)) + (10 ^ (Loss3 / 10))))
    End If

End Function

'==============================================================================
' Name:     PlenumLossOneThirdOctave_ASHRAE
' Author:   PS
' Desc:     Calculates Plenum loss according to the ASHRAE method.
'           Hold on, it's gonna get tricky
' Args:     fStr - Octave band centre frequency Hz
'           L - length of plenum in mm
'           W - width of plenum in mm
'           H - Height of plenum in mm
'           DuctInL - Length of inlet in mm
'           DuctInW - Width of inlet in mm
'           DuctOutL - Length of outlet in mm
'           DuctOutW - Width of outlet in mm
'           Q - 2 or 4 for type of spherical spreading
'           r_h - hoizontal offset in mm
'           r_v - vertical offset in mm
'           PlenumLiningType - Concrete, sheet metal or lined sheet metal duct
'           UnlinedType - The unlined plenum material (behind lining)
'           WallEffect - Different types for wall effect options in ASHRAE
'           applyElbowEffect - TRUE if tickbox is checked
'           PercentUnlined - Percent of surface area that's unlined
' Comments: (1)
'==============================================================================
Function PlenumLossOneThirdOctave_ASHRAE(f As Double, L As Long, W As Long, _
H As Long, DuctInL As Single, DuctInW As Single, DuctOutL As Single, _
DuctOutW As Single, Q As Integer, R_H As Long, r_v As Long, _
PlenumLiningType As String, UnlinedType As String, WallEffect As String, _
applyElbowEffect As Boolean, UnlinedPercent As Long)

Dim Stotal As Single 'Total internal surface area of plenum in m^2
Dim InletArea As Single 'area of inlet duct in m^2
Dim OutletArea As Single 'area of outlet duct in m^2
Dim PlenumVolume As Single 'total volume of plenum in m^3
Dim R As Single 'inlet to outlet offset distance
Dim alphaTotal(7) As Single 'absorption coefficient of each octave band
Dim AbsorptionArea(7) As Single 'S*alpha for each octave band
Dim offsetAngle As Single 'angle betweeen inlet and outlet ducts
Dim b As Single 'Constant from ASHRAE
Dim n As Single 'Constant from ASHRAE
Dim f_co As Double 'cutoff frequency
Dim OAE As Single 'offset angle effect
Dim WallEffectIndex As Integer 'index for the defined types in ASHRAE
Dim W_e As Single 'Wall effect in dB
Dim AngleEffect As Single 'Angle effect in dB
Dim DuctInL_m As Single
Dim DuctInW_m As Single
Dim DuctOutL_m As Single
Dim DuctOutW_m As Single
Dim SA_Lined As Single
Dim SA_Unlined As Single

'''''''''''''''''''''''''''''''''''''''''''''''''
'CONSTANTS
'''''''''''''''''''''''''''''''''''''''''''''''''

'Values from ASHRAE equation 5
b = 3.505
n = -0.359
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
DuctInL_m = DuctInL / 1000 'convert to metres
DuctInW_m = DuctInW / 1000 'convert to metres
DuctOutL_m = DuctOutL / 1000 'convert to metres
DuctOutW_m = DuctOutW / 1000 'convert to metres

'''''''''''''''''''''''''''''''''''''''''''''''''
'CALCULATION
'''''''''''''''''''''''''''''''''''''''''''''''''

'cutoff frequency
f_co = PlenumCutoffFrequency(DuctInL_m, DuctInW_m)
'Debug.Print "Cutoff Frequency = " & Round(f_co, 1) & "Hz"

'Areas and Volumes
InletArea = DuctInL_m * DuctInW_m
OutletArea = DuctOutL_m * DuctOutW_m
PlenumVolume = (L / 1000) * (W / 1000) * (H / 1000) 'note: input in mm

'Surface area
'Stotal doesn't include inlet and outlet area
Stotal = PlenumSurfaceArea(L, W, H, InletArea, OutletArea)
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UnlinedType
    Case Is = "Concrete"
    Unlined = Concrete
    Case Is = "Bare Sheet Metal"
    Unlined = Bare_Sheet_Metal
    End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SA_Unlined = Stotal * (UnlinedPercent / 100) 'm2
SA_Lined = Stotal - SA_Unlined 'm2
    For i = 0 To UBound(alphaTotal) - 1 '7th column not used
    alphaTotal(i) = (((InletArea + OutletArea + SA_Unlined) * Unlined(i)) + (SA_Lined * PlenumLining(i))) _
        / (InletArea + OutletArea + Stotal)
    AbsorptionArea(i) = (OutletArea * (1 - alphaTotal(i))) / (Stotal * alphaTotal(i))
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Distance from inlet to outlet
R = PlenumDistanceR(R_H, r_v, L)
'Debug.Print "Offset Distance, R = " & Round(r, 1)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '90-degree bend selected, use the elbow effect method in ASHRAE
    If applyElbowEffect = True Then
    AngleEffect = PlenumElbowEffect(f, f_co)
    Else
    'Offset Angle
    offsetAngle = PlenumAngleTheta(L, R)
    'Debug.Print "Offset Angle = " & Round(OffsetAngle, 2)
    'Offset Angle Effect
    AngleEffect = PlenumOAE(f, f_co, offsetAngle)
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'below critical frequency
    If f < f_co Then
    
        If WallEffect = "False" Or WallEffect = "" Then
        WallEffectIndex = 0
        Else
        WallEffectIndex = CInt(Left(WallEffect, 1))
        End If
    
    'calculate Wall Effect
    W_e = PlenumWallEffect(f, WallEffectIndex)
    'Debug.Print "Wall effect = " & W_e
    
    'calculate A_f
    'note: ASHRAE includes x10.76 in table 13, so there's no need to include twice
    A_f = PlenumAreaCoefficient(f, PlenumVolume)
    'Debug.Print "Area Coefficient, A_f = " & A_f
    
    PlenumLossOneThirdOctave_ASHRAE = _
        -1 * Application.WorksheetFunction.Min((A_f * Stotal) _
        + W_e + AngleEffect, 20) 'limit to 20dB, output is negative
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else 'f>=f_co, above critical frequency
    A_index = MapOneThird2Oct(f)
    
        If A_index = -1 Then 'catch errors with index out of bounds
        PlenumLossOneThirdOctave_ASHRAE = 0
        Else
        'big formula! Eqn 10 from Chapter 48 of ASHRAE
        PlenumLossOneThirdOctave_ASHRAE = -1 * (b * (((OutletArea * Q / _
            (4 * Application.WorksheetFunction.Pi() * (R ^ 2))) _
            + AbsorptionArea(A_index)) ^ n) + AngleEffect)
        End If
    End If
    
'Debug.Print f; "Hz "; PlenumLossOneThirdOctave_ASHRAE

End Function

'==============================================================================
' Name:     PlenumDistanceR
' Author:   PS
' Desc:     Calculates distance between the centres of plenum inlet and outlet
' Args:     r_h - horizontal offset in mm
'           r_v - vertical offset in mm
'           L - length of plenum in mm
' Comments: (1) Pythagoras would be proud
'           (2) Output is in metres
'==============================================================================
Function PlenumDistanceR(R_H As Long, r_v As Long, L As Long) As Single
PlenumDistanceR = (((r_v / 1000) ^ 2) + ((R_H / 1000) ^ 2) + _
                    ((L / 1000) ^ 2)) ^ 0.5
End Function

'==============================================================================
' Name:     PlenumAngleTheta
' Author:   PS
' Desc:     Calculates angle between pleenum inlet and outlet
' Args:     L - length of plenum in mm
'           R - Radial distance in metres
' Comments: (1) Output in degrees
'==============================================================================
Function PlenumAngleTheta(L As Long, R As Single)
Dim PlenumL As Single
PlenumL = L / 1000 'convert to metres
    If PlenumL / R >= -1 And PlenumL / R <= 1 Then 'between -1 and 1
    PlenumAngleTheta = Application.WorksheetFunction.Degrees( _
        Application.WorksheetFunction.Acos(PlenumL / R))
    Else
    PlenumAngleTheta = 0
    End If
End Function

'==============================================================================
' Name:     PlenumCutoffFrequency
' Author:   PS
' Desc:     Calculates the cutoff frequency (Hz), based on the largest
'           dimension of the plenum
' Args:     L - length of plenum in metres
'           W - width of plenum in metres
'           SpeedOfSound - in m/s, defaults to 343
' Comments: (1) Assumes speed of sound is 343m/s
'==============================================================================
Function PlenumCutoffFrequency(L As Single, W As Single, _
    Optional SpeedOfSound As Double) As Single
    
    If SpeedOfSound = 0 Then SpeedOfSound = 343

PlenumCutoffFrequency = SpeedOfSound / (2 * Application.Max(L, W))
End Function

'==============================================================================
' Name:     PlenumSurfaceArea
' Author:   PS
' Desc:     Calculates the cutoff frequency
' Args:     L - length of plenum in mm
'           W - width of plenum in mm
'           H - height of plenum in mm
'           InletArea - area of inlet in m^2
'           OutletArea - area of outlet in m^2
' Comments: (1) inputs are in mm, which are squared
'           => correction is 1000x1000 = 1million
'==============================================================================
Function PlenumSurfaceArea(L As Long, W As Long, H As Long, _
InletArea As Single, OutletArea As Single) As Single
PlenumSurfaceArea = (2 * L * W / 1000000) + (2 * W * H / 1000000) + _
    (2 * H * L / 1000000) - InletArea - OutletArea
End Function


'==============================================================================
' Name:     PlenumAreaCoefficient
' Author:   PS
' Desc:     Returns area coefficient from ASHRAE Table 13
' Args:     f_input - one-third octave band centre frequencyy
'           Vol - voume of plenum in m^3
' Comments: (1)
'==============================================================================
Function PlenumAreaCoefficient(f_input As Double, Vol As Single)

'reference values from Table 13
'bands              50   63  80   100  125 160 200 250  315  400  500
SmallPlenum = Array(1.4, 1#, 1.1, 2.3, 2.4, 2, 1#, 2.2, 0.7, 0.7, 1.1)
'bands              50   63   80   100  125  160  200  250  315  400  500
LargePlenum = Array(0.3, 0.3, 0.3, 0.3, 0.4, 0.4, 0.3, 0.4, 0.3, 0.2, 0.2)

i = GetArrayIndex_TO(f_input)

    If i = -1 Then 'catch errrors
    PlenumAreaCoefficient = 0
    ElseIf i <= 10 Then
        If Vol < 1.5 Then 'm^3
        PlenumAreaCoefficient = SmallPlenum(i)
        Else 'Vol>1.5m^3
        PlenumAreaCoefficient = LargePlenum(i)
        End If
    Else
    PlenumAreaCoefficient = 0
    End If

End Function

'==============================================================================
' Name:     PlenumOAE
' Author:   PS
' Desc:     Gets offset angle effect from Table 14 of ASHRAE and interpolates
'           for Angle_input
' Args:     f_input - one third octave band centre frequency Hz
'           f_co - cutoff one third octave band centre frequency Hz
'           Angle_input - angle in degrees
' Comments: (1) Error catch for y1 and y2 - not needed
'           (2) Changed x1 and x2 to Singles
'==============================================================================
Function PlenumOAE(f_input As Double, f_co As Double, Angle_input As Single)
Dim x1, x2 As Single
Dim y1, y2 As Single
Dim i As Integer 'index of array
Dim Slope As Single

    If f_input <= f_co Then
        Select Case f_input
        Case Is < 50
        OAETable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 50
        OAETable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 63
        OAETable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 80
        OAETable = Array(0, 0, -1, -3, -4, -6)
        Case Is = 100
        OAETable = Array(0, 1, 0, -2, -3, -6)
        Case Is = 125
        OAETable = Array(0, 1, 0, -2, -4, -6)
        Case Is = 160
        OAETable = Array(0, 0, -1, -2, -3, -4)
        Case Is = 200
        OAETable = Array(0, 0, -1, -2, -3, -5)
        Case Is = 250
        OAETable = Array(0, 1, 2, 3, 5, 7)
        Case Is = 315
        OAETable = Array(0, 4, 6, 8, 10, 14)
        Case Is = 400
        OAETable = Array(0, 2, 4, 6, 9, 13)
        Case Is = 500
        OAETable = Array(0, 1, 3, 6, 10, 15)
        Case Is = 630
        OAETable = Array(0, 0, 0, 0, 0, 0)
        End Select
    Else 'f_input>f_co
        Select Case f_input
        Case Is < 200 'catch exception, default to zero?
        OAETable = Array(0, 0, 0, 0, 0, 0)
        Case Is = 200
        OAETable = Array(0, 1, 4, 9, 14, 20)
        Case Is = 250
        OAETable = Array(0, 2, 4, 8, 13, 19)
        Case Is = 315
        OAETable = Array(0, 1, 2, 3, 4, 5)
        Case Is = 400
        OAETable = Array(0, 1, 2, 3, 4, 6)
        Case Is = 500
        OAETable = Array(0, 0, 1, 2, 4, 5)
        Case Is = 630
        OAETable = Array(0, 1, 2, 3, 5, 7)
        Case Is = 800
        OAETable = Array(0, 1, 2, 2, 3, 3)
        Case Is = 1000
        OAETable = Array(0, 1, 2, 4, 6, 9)
        Case Is = 1250
        OAETable = Array(0, 0, 2, 4, 6, 9)
        Case Is = 1600
        OAETable = Array(0, 0, 1, 1, 2, 3)
        Case Is = 2000
        OAETable = Array(0, 1, 2, 4, 7, 10)
        Case Is = 2500
        OAETable = Array(0, 1, 2, 3, 5, 8)
        Case Is = 3150
        OAETable = Array(0, 0, 2, 4, 6, 9)
        Case Is = 4000
        OAETable = Array(0, 0, 2, 5, 8, 12)
        Case Is = 5000
        OAETable = Array(0, 0, 3, 6, 10, 15)
        Case Is > 5000 'catch exception, default to zero?
        OAETable = Array(0, 0, 0, 0, 0, 0)
        End Select
    End If

    'get x values for interpolation
    Select Case Angle_input
    Case Is = 0 'set to 0 for error catching
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
    If (x1 = 0 And x2 = 0) Then
    PlenumOAE = 0
    Else
    y1 = OAETable(i)
    y2 = OAETable(i + 1)
    'interpolate things
    Slope = (y2 - y1) / (x2 - x1)
    PlenumOAE = y1 + (Slope * (Angle_input - x1))
    End If


End Function

'==============================================================================
' Name:     PlenumWallEffect
' Author:   PS
' Desc:     Table 13 of ASHRAE - Low Frequency Characdteristics of Plenum TL
' Args:     f - one-third octave band centre frequency
'           WallType - 1 to 6, from the 6 types defined in the table
' Comments: (1)
'==============================================================================
Function PlenumWallEffect(f As Double, WallType As Integer)
Dim i As Integer

i = GetArrayIndex_TO(f)

    If i = -1 Then 'catch errors
    PlenumWallEffect = 0
    Exit Function
    End If

'reference values from Table 13
'band       50 63 80 100 125 160 200 250 315 400 500
WType1 = Array(1, 1, 2, 2, 2, 3, 4, 5, 6, 8, 9)
WType2 = Array(1, 2, 2, 2, 3, 4, 10, 9, 12, 13, 13)
WType3 = Array(0, 3, 3, 4, 6, 11, 16, 13, 14, 13, 12)
WType4 = Array(1, 7, 9, 12, 12, 11, 15, 12, 14, 14, 13)
WType5 = Array(0, 1, 2, 1, 1, 0, 4, 1, 5, 7, 8)
WType6 = Array(0, 3, 7, 6, 4, 2, 3, 1, 2, 1, 0)


    If i <= UBound(WType1) Then 'maximum 11 elements in WType arrays
        Select Case WallType
        Case Is = 0
        PlenumWallEffect = 0
        Case Is = 1
        PlenumWallEffect = WType1(i)
        Case Is = 2
        PlenumWallEffect = WType2(i)
        Case Is = 3
        PlenumWallEffect = WType3(i)
        Case Is = 4
        PlenumWallEffect = WType4(i)
        Case Is = 5
        PlenumWallEffect = WType5(i)
        Case Is = 6
        PlenumWallEffect = WType6(i)
        End Select
    Else
    PlenumWallEffect = 0
    End If

End Function

'==============================================================================
' Name:     PlenumElbowEffect
' Author:   PS
' Desc:     Returns values from Table 15 of ASHRAE for Elbow effect of plenums
' Args:     f - one-third octave frequency band
'           f_c - critical frequency
' Comments: (1)
'==============================================================================
Function PlenumElbowEffect(f As Double, f_c As Double)
Dim i As Integer
'bands         50 63 80 100 125 160 200 250 315 400 500 630 800 1k 1.25k 1.6k 2k 3.15k 4k
BelowFc = Array(2, 3, 6, 5, 3, 0, -2, -3, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
AboveFc = Array(0, 0, 0, 0, 0, 0, 3, 6, 3, 3, 2, 3, 3, 2, 2, 2, 2, 2, 2, 2, 1)
i = GetArrayIndex_TO(f)

    If i <= 20 Then
        If f > f_c Then
        PlenumElbowEffect = AboveFc(i)
        Else 'f<=f_c
        PlenumElbowEffect = BelowFc(i)
        End If
    Else
    PlenumElbowEffect = 0
    End If
    
End Function

'==============================================================================
' Name:     GetAdjacentFrequency
' Author:   PS
' Desc:     Returns one-third octave band above or below input band
'           Used to generate ASHRAE plenum loss in octave bands
' Args:     f_input - one-third octave band centre frequency Hz
'           AdjMode - "Up" or "Down"
' Comments: (1)
'==============================================================================
Function GetAdjacentFrequency(f_input As Double, AdjMode As String) As Double
Dim adjustIndex  As Integer
f_ref = Array(50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, _
1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300, 8000, 10000)

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
        Exit Function
        End If
    Next i

End Function

'==============================================================================
' Name:     GetDuctArea
' Author:   PS
' Desc:     Returns duct cross sectional area in m^2, given an input string
'           from formulas
' Args:     inputStr - the string to be split up
' Comments: (1) Is this needed anymore?
'==============================================================================
Function GetDuctArea(inputStr As String)
Dim SplitStr() As String
Dim L As Double
Dim W As Double
SplitStr = Split(inputStr, ",", Len(inputStr), vbTextCompare)
L = CDbl(SplitStr(1))
W = CDbl(SplitStr(2))
GetDuctArea = (L / 1000) * (W / 1000) 'because millimetres
End Function

'==============================================================================
' Name:     GetDuctParameter
' Author:   PS
' Desc:     Returns duct parameters from a formula, by splitting up the
'           arguments separated by commas
' Args:     inputStr - the formula string to be split up
'           Parameter - name of the element to be extracted
' Comments: (1) Used to forward on variable to adjacent functions
'           e.g. put duct dimensions from row above in ERL
'==============================================================================
Function GetDuctParameter(inputStr As String, Parameter As String)
Dim SplitStr() As String
Dim L As Single
Dim W As Single
Dim Area As Single
On Error GoTo errCatch
SplitStr = Split(inputStr, ",", Len(inputStr), vbTextCompare)
L = CSng(SplitStr(1))
W = CSng(SplitStr(2))
Area = (L / 1000) * (W / 1000) 'because millimetres
    Select Case Parameter
    Case Is = "Area"
    GetDuctParameter = Area
    Case Is = "L"
    GetDuctParameter = L
    Case Is = "W"
    GetDuctParameter = W
    End Select
Exit Function
errCatch:
    GetDuctParameter = ""
End Function

'==============================================================================
' Name:     FantechAttenRegen
' Author:   AN
' Desc:     Calculates sound power generated from Fantech rectangular silencers
' Args:     fStr - octave band centre frequency
'           airflow - in m^3/s
'           percentage_free_area - value up to 100
'           width - Width of duct in mm
'           SplitterHeight - Height of splitters in a module, in mm
'           numModules - number of modules in silencer
' Comments: (1) Reading from curves in the Fantech book, we've derived the
'           underlying formulas with simultaneous equations!
'==============================================================================
Function FantechAttenRegen(fStr As String, airflow As Double, _
percentage_free_area As Double, Width As Double, SplitterHeight As Double, _
numModules As Integer, Optional litresPerSecond As Boolean)

Dim airwayVelocity As Single 'in m/s
Dim base_sound_power_level As Double 'in dB
Dim SWL As Double 'total SWL for all modules
Dim AV_correction As Integer 'velocity correction

    'Check for errors
    If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Or _
    percentage_free_area <= 0 Or airflow = 0 Or Width = 0 Or SplitterHeight = 0 Then
    FantechAttenRegen = "-"
    Exit Function
    End If
    
'optional switch for L/s
If litresPerSecond = True Then airflow = airflow / 1000

'Main working
airwayVelocity = (airflow * 100) / ((Width / 1000) * (SplitterHeight / 1000) * _
    percentage_free_area)
                
base_sound_power_level = (50.6 * Application.WorksheetFunction.Log(airwayVelocity)) + _
    (10 * Application.WorksheetFunction.Log(SplitterHeight)) - 33.8
    
SWL = base_sound_power_level + (10 * Application.WorksheetFunction.Log(numModules))

'spectrum Corrections, from 63Hz octave band
AV_correction = GetFantechAirwayVelocityCorrection(fStr, airwayVelocity)

'Final output
FantechAttenRegen = SWL + AV_correction

End Function

'==============================================================================
' Name:     NAPAttenRegen
' Author:   PS
' Desc:     Calculates sound power generated from NAP rectangular silencers
' Args:     fStr - octave band centre frequency
'           airflow - in m^3/s
'           percentage_free_area - value up to 100
'           width - Width of duct in mm
'           SplitterHeight - Height of splitters in a module, in mm
'           numModules - number of modules in silencer
' Comments: (1) Reading from curves in the Fantech book, we've derived the
'           underlying formulas with simultaneous equations!
'==============================================================================
Function NAPAttenRegen(fStr As String, airflow As Double, _
percentage_free_area As Double, Width As Double, Height As Double, _
Model As String, Optional litresPerSecond As Boolean)

Dim airwayVelocity As Single 'in m/s
Dim base_sound_power_level As Double 'in dB
Dim SpectrumCorrection As Integer 'as provided by NAP
Dim ModelCorrection As Integer 'as provided by NAP
Dim FaceArea As Double

    'Check for errors
    If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Or _
        percentage_free_area <= 0 Or percentage_free_area > 100 Or airflow = 0 Or _
        Width = 0 Or Height = 0 Then
    NAPAttenRegen = "-"
    Exit Function
    End If

'optional switch for L/s
If litresPerSecond = True Then airflow = airflow / 1000

FaceArea = Width * Height / 1000000

airwayVelocity = (airflow) / _
    (FaceArea * (percentage_free_area / 100))

'TOODO: check this formula
base_sound_power_level = (50 * Application.WorksheetFunction.Log(airwayVelocity)) + _
    (10 * Application.WorksheetFunction.Log(FaceArea)) + 2

'spectrum and model corrections
SpectrumCorrection = GetNAPSpectrumCorrection(fStr, Model)
ModelCorrection = GetNAPModelCorrection(Model)

'Final output
NAPAttenRegen = base_sound_power_level + ModelCorrection + SpectrumCorrection

End Function

'==============================================================================
' Name:     GetNAPSpectrumCorrection
' Author:   PS
' Desc:     Returns model correction for NAP rectangular Silencers (from book)
' Args:     Model - Model type string
' Comments: (1) As supplied by supplier, no edits
'==============================================================================
Function GetNAPSpectrumCorrection(fStr As String, Model As String)

Dcorrections = Array(-2, -6, -7, -10, -12, -16, -19, -22)
Ecorrections = Array(-3, -5, -8, -7, -8, -10, -13, -15)
Hcorrections = Array(-3, -6, -10, -7, -7, -8, -10, -12)

    'Check if appropriate column
    If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Then
    GetNAPSpectrumCorrection = "-"
    Exit Function
    End If

i = GetArrayIndex_OCT(fStr)

    If UCase(Left(Model, 1)) = "D" Then
    GetNAPSpectrumCorrection = Dcorrections(i)
    ElseIf UCase(Left(Model, 1)) = "E" Then
    GetNAPSpectrumCorrection = Ecorrections(i)
    ElseIf UCase(Left(Model, 1)) = "H" Then
    GetNAPSpectrumCorrection = Hcorrections(i)
    Else
    GetNAPSpectrumCorrection = 0
    End If
End Function


'==============================================================================
' Name:     GetNAPModelCorrection
' Author:   PS
' Desc:     Returns model correction for NAP rectangular Silencers (from book)
' Args:     Model - Model type string
' Comments: (1) As supplied by supplier, no edits
'==============================================================================
Function GetNAPModelCorrection(Model As String)
Dim CheckValue As Integer
CheckValue = 999

    Select Case UCase(Left(Model, 3))
    Case Is = "D27"
    CheckValue = 0
    Case Is = "D33"
    CheckValue = -3
    Case Is = "D38"
    CheckValue = -5
    Case Is = "D43"
    CheckValue = -6
    Case Is = "D47"
    CheckValue = -8
    Case Is = "D50"
    CheckValue = -9
    Case Is = "E29"
    CheckValue = -1
    Case Is = "E38"
    CheckValue = -4
    Case Is = "E44"
    CheckValue = -7
    Case Is = "E50"
    CheckValue = -9
    Case Is = "H33"
    CheckValue = -3
    Case Is = "H40"
    CheckValue = -5
    Case Is = "H45"
    CheckValue = -7
    Case Is = "H50"
    CheckValue = -9
    End Select

    'check for errors
    If CheckValue = 999 Then
    GetNAPModelCorrection = 0
    End If
End Function

'==============================================================================
' Name:     GetFantechAirwayVelocityCorrection
' Author:   AN
' Desc:     Returns velcity correction for Fantech Silencers (from book)
' Args:     fStr - octave band centre frequency
'           airwayVelocity - air speed in m/s
' Comments: (1) As supplied by supplier, no edits
'==============================================================================
Function GetFantechAirwayVelocityCorrection(fStr As String, airwayVelocity As Single)
Dim i As Integer
LessThan8 = Array(-2, -6, -7, -10, -12, -16, -19, -22)
EightTo32 = Array(-3, -5, -8, -7, -8, -10, -13, -15)
MoreThan32 = Array(-3, -6, -10, -7, -7, -8, -10, -12)

    'Check if appropriate column
    If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Then
    GetFantechAirwayVelocityCorrection = "-"
    Exit Function
    End If

i = GetArrayIndex_OCT(fStr)

    If airwayVelocity < 8 Then
    GetFantechAirwayVelocityCorrection = LessThan8(i)
    ElseIf airwayVelocity >= 8 And airwayVelocity <= 32 Then
    GetFantechAirwayVelocityCorrection = EightTo32(i)
    Else
    GetFantechAirwayVelocityCorrection = MoreThan32(i)
    End If

End Function

'==============================================================================
' Name:     DuctDirectivity_PGD
' Author:   NI
' Desc:
' Args:     freq - octave band centre frequency
'           angle - angle from normal line through centre of duct, in degrees
'           diameter - duct diameter in mm
' Comments: (1) Reads central text file DUCT_DIRLOSS
'           (2) Values are from a technical paper:
'           *Directivity Loss  at Duct Terminaton* by Daniel Potente,
'           Stepehen Gauld and Athol Day which can be found at:
'           https://www.acoustics.asn.au/conference_proceedings/AASNZ2006/
'           papers/p103.pdf
'==============================================================================
Function DuctDirectivity_PGD(freq As String, Angle As Double, diameter As Double)

On Error GoTo closefile

Dim ReadStr() As String 'for holding values from text file
Dim SplitStr() As String 'for splitting into array
Dim SplitVal() As Double 'for holding array as values
Dim i As Integer 'counter for text line number
Dim col As Integer 'counter for each column of results
Dim found As Boolean 'switch for matching inputs to values from text file

Call GetSettings 'set public variables, including text file location

Open DUCT_DIRLOSS For Input As #1

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then
        
            'convert to values
            For col = 0 To UBound(SplitStr)
                If SplitStr(col) <> "" Then
                ReDim Preserve SplitVal(col)
                SplitVal(col) = CDbl(SplitStr(col))
                End If
            Next col
            
            ReDim Preserve SplitVal(col + 1)
            '<-TODO, simplify this lookup with the getoctindex function
               If SplitVal(0) = diameter And SplitVal(1) = Angle Then
                Select Case freq 'catch for both kinds of header
                Case Is = "63"
                DuctDirectivity_PGD = SplitVal(2)
                Case Is = "125"
                DuctDirectivity_PGD = SplitVal(3)
                Case Is = "250"
                DuctDirectivity_PGD = SplitVal(4)
                Case Is = "500"
                DuctDirectivity_PGD = SplitVal(5)
                Case Is = "1k"
                DuctDirectivity_PGD = SplitVal(6)
                Case Is = "2k"
                DuctDirectivity_PGD = SplitVal(7)
                Case Is = "4k"
                DuctDirectivity_PGD = SplitVal(8)
                Case Is = "8k"
                DuctDirectivity_PGD = SplitVal(9)
                Case Is = "1000"
                DuctDirectivity_PGD = SplitVal(6)
                Case Is = "2000"
                DuctDirectivity_PGD = SplitVal(7)
                Case Is = "4000"
                DuctDirectivity_PGD = SplitVal(8)
                Case Is = "8000"
                DuctDirectivity_PGD = SplitVal(9)
                Case Else
                DuctDirectivity_PGD = "-"
                End Select
                
            End If
            
        End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function


'==============================================================================
' Name:     DamperRegen_NEBB
' Author:   IV & PS
' Desc:     Returns regenerated sound power from a duct damper (NEBB method)
' Args:     fStr - Octave band centre frequency Hz
'           flowrate - in L/s
'           PressureLoss - Pressure loss across the damper
'           DuctHeight - in mm
'           DuctWidth - in mm
'           MultiBlade - set to TRUE for multi-blade dampers
'           mCubedPerSection - set to TRUE for m^3/s flow rates
' Comments: (1)
'==============================================================================
Function DamperRegen_NEBB(fStr As String, FlowRate As Double, PressureLoss As Double, _
DuctHeight As Double, DuctWidth As Double, MultiBlade As Boolean, _
Optional mCubedPerSecond As Boolean)

Dim PLCoeff As Double 'pressure loss coefficient
Dim BlockageFactor As Double
Dim Uc As Double 'Flow velocity in the constricted part of the flow field
Dim St As Double 'Strouhal's number - Equation 4.5 of NEBB
Dim Kd As Double 'Characteristic spectrum - Equation 4.6 of NEBB
Dim f As Double 'frequency as number
Dim CrossSectionArea As Double 'area of the duct in m^2

f = freqStr2Num(fStr)
    
    'catch errors
    If FlowRate <= 0 Or PressureLoss <= 0 Or _
        DuctWidth <= 0 Or DuctHeight <= 0 Then
    DamperRegen_NEBB = "-"
    Exit Function
    End If

CrossSectionArea = (DuctWidth * DuctHeight) / 1000000 'area in m2

If mCubedPerSecond = True Then FlowRate = FlowRate * 1000

'Equation 4.2 of NEBB
PLCoeff = 16.4 * 100000 * PressureLoss * (1 / ((FlowRate / CrossSectionArea) ^ 2))

    If MultiBlade = True Then
        'equation 4.3a of NEBB
        If PLCoeff = 1 Then
        BlockageFactor = 0.5
        Else
        BlockageFactor = ((Sqr(PLCoeff)) - 1) / (PLCoeff - 1)
        End If
    Else 'multiblade is false
        'equation 4.3b of NEBB
        If PLCoeff <= 4 Then
        BlockageFactor = ((Sqr(PLCoeff)) - 1) / (PLCoeff - 1)
        Else
        BlockageFactor = (0.68 * (PLCoeff ^ -0.15)) - 0.22
        End If
    End If

'Equation 4.4 of NEBB
Uc = 0.001 * (FlowRate / (CrossSectionArea * BlockageFactor))

'Equation 4.5 of NEBB
St = (f * (DuctHeight / 1000)) / Uc

    'catch error
    If St < 0 Then
    DamperRegen_NEBB = "-"
    Exit Function
    End If
    
    'Equation 4.6 of NEBB
    If St > 25 Then
    Kd = -1.1 - (35.9 * Application.WorksheetFunction.Log10(St))
    Else 'St<=25
    Kd = -36.3 - (10.7 * Application.WorksheetFunction.Log10(St))
    End If

'Equation 4.2 of NEBB
DamperRegen_NEBB = Kd + (10 * Application.WorksheetFunction.Log10(f / 63)) + _
    (50 * Application.WorksheetFunction.Log10(3.28 * Uc)) + _
    (10 * Application.WorksheetFunction.Log10(10.76 * CrossSectionArea)) + _
    (10 * Application.WorksheetFunction.Log10(3.28 * (DuctHeight / 1000)))

End Function

'==============================================================================
' Name:     ElbowWithVanesRegen_NEBB
' Author:   AA
' Desc:     Calculates elbow (with vanes) regenerated noise according to the
'           NEBB method.
' Args:     fstr - Octave band centre frequency (Hz, string)
'           FlowRate - volumetric flow rate (L/s, double)
'           dP - delta Pressure, i.e. drop across damper (Pa, double)
'           DuctWidth - Duct Width normal to vane axis (mm, double)
'           DuctHeight - Duct Height parallel to vane axis (mm, double)
'           CordLength - cord length of a typical vane (mm, double)
'           n - number of turning vanes (integer)
'           mCubedPerSection - set to TRUE for m^3/s flow rates
' Comments: (1)
'==============================================================================
Function ElbowWithVanesRegen_NEBB(fStr As String, FlowRate As Double, _
    dP As Double, DuctWidth As Double, DuctHeight As Double, CordLength As Double, _
    numVanes As Integer, Optional mCubedPerSecond As Boolean)

Dim f As Single 'frequency band (Hz)
Dim DuctArea As Double 'duct cross-sectional area (m^2)
Dim c As Double 'pressure loss Coefficient
Dim BF As Double 'Blockage Factor
Dim Uc As Double 'flow velocity in damper constriction
Dim St As Double 'Strouhal number
Dim Kt As Double 'characteristic spectrum

'General setup and error catching
f = freqStr2Num(fStr)
    If f < 63 Or f > 8000 Then
    ElbowWithVanesRegen_NEBB = "-"
    Exit Function
    End If

DuctWidth = DuctWidth / 1000 'convert to m
DuctHeight = DuctHeight / 1000 'convert to m

DuctArea = DuctWidth * DuctHeight 'calc duct cross-sectional area

    If mCubedPerSecond = True Then FlowRate = FlowRate * 1000

'Step 1: Calculate pressure loss coefficient, C
c = 16.4 * 100000 * dP * (1 / ((FlowRate / DuctArea) ^ 2))
    
    ' Step 2: Calculate Blockage Factor, BF
    If c = 1 Then
    BF = 0.5
    Else
    BF = ((Sqr(c)) - 1) / (c - 1)
    End If

'Step 3: Calculate flow velocity in damper constriction, Uc
Uc = 0.001 * (FlowRate / (DuctArea * BF))

'Step 4: Calculate Strouhal number, St
St = (f * DuctHeight) / Uc

'Step 5: Calculate characteristic spectrum, Kt
    If St > 1 Then
    Kt = -47.4 - (7.69 * (Application.WorksheetFunction.Log10(St) ^ 2.5))
    End If

'FINAL: Calculate elbow (with vanes) regen
ElbowWithVanesRegen_NEBB = Kt + 10 * Application.WorksheetFunction.Log10(f / 63) _
    + 50 * Application.WorksheetFunction.Log10(3.28 * Uc) _
    + 10 * Application.WorksheetFunction.Log10(10.76 * DuctArea) _
    + 10 * Application.WorksheetFunction.Log10(0.039 * CordLength) _
    + 10 * Application.WorksheetFunction.Log10(numVanes)

End Function

'==============================================================================
' Name:     ElbowOrJunctionRegen_NEBB
' Author:   AA
' Desc:     Calculates elbow (without vanes) or junction regenerated noise
'           according to the NEBB method.
' Args:     fstr - Octave band centre frequency (Hz, string)
'           FlowRate - main duct volumetric flow rate (L/s)
'           IsMainCircular - set to TRUE if the main duct is circular
'           DuctWidth - diameter or width of main duct (mm, Double)
'           DuctHeight - height of main duct, if Tm is circular this remains
'               unused (mm, Double)
'           BranchFlowRate - branch duct volumetric flow rate (L/s)
'           IsBranchCircular - set to TRUE if the branch is circular
'           DuctBranchWidth - diameter or width of branch duct(mm, Double)
'           DuctBranchHeight - height of branch duct, if Tb is circular this
'               remains unused(mm, Double)
'           Radius - radius of bend or elbow (mm)
'           IsTurbulent - whether turbulence is present. Correction applied if damper,
'               elbow, or takeoff upstream within 5 main duct diameters of turn
'               TRUE: turbulence present
'               FALSE: no turbulence present
'           JunctionType - type of junction ({1,2,3,4}, integer)
'               1: elbow bend
'               2: 90 degree branch takeoff
'               3: T-Junction
'               4: X-Junction
'           BranchRegen - set to TRUE to predict regenerated noise in branch,
'           instead of main duct
'           mCubedPerSecond - set to TRUE for m^3/s flow rates
' Comments: (1)
'==============================================================================
Function ElbowOrJunctionRegen_NEBB(fStr As String, _
    FlowRate As Double, IsMainCircular As Boolean, DuctWidth As Double, _
    DuctHeight As Double, BranchFlowRate As Double, IsBranchCircular As Boolean, _
    DuctBranchWidth As Double, DuctBranchHeight As Double, Radius As Double, _
    IsTurbulent As Boolean, JunctionType As Integer, BranchRegen As Boolean, _
    Optional mCubedPerSecond As Boolean)

Dim f As Single  'frequency band (Hz)
Dim MainDuctArea As Double 'main duct cross-sectional area (m^2)
Dim BranchDuctArea As Double 'branch duct cross-sectional area (m^2)
Dim Um As Double 'main duct flow velocity (m/s)
Dim Ub As Double 'branch duct flow velocity (m/s)
Dim WidthRatio As Double 'ratio of DuctWidth/DuctBranchWidth, diameters of main and branch ducts
Dim VelocityRatio As Double 'ratio of Um/Ub, flow velocities of main and branch ducts
Dim RD As Double 'rounding parameter for use in calculating RadiusCorrection
Dim RadiusCorrection As Double 'correction term that quantifies the effect of the size of
'                 the radius of the bend/elbow associated with the turn/junction
Dim TurbCorrection As Double 'correction term quantifying effect of turbulence. Refer to
'                 turb in preamble for extra information
Dim Kj As Double 'characteristic spectrum, dB
Dim Lb As Double 'branch SWL result, dB
Dim Lm As Double 'main duct SWL result, dB

' GENERAL SETUP
f = freqStr2Num(fStr)
If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Then
    ElbowOrJunctionRegen_NEBB = "-"
    Exit Function
End If

'convert dimensions to metres
DuctWidth = DuctWidth / 1000
DuctHeight = DuctHeight / 1000
DuctBranchWidth = DuctBranchWidth / 1000
DuctBranchHeight = DuctBranchHeight / 1000
Radius = Radius / 1000

    If mCubedPerSecond = True Then 'allow for m3/s flowrates
    FlowRate = FlowRate * 1000
    BranchFlowRate = BranchFlowRate * 1000
    End If

    'Determine/calc cross-sectional area, MainDuctArea and BranchDuctArea
    If IsMainCircular = True Then 'circular duct
    MainDuctArea = WorksheetFunction.Pi * (DuctWidth / 2) ^ 2
    Else  'rectangular duct
    MainDuctArea = DuctWidth * DuctHeight
    DuctWidth = (4 * MainDuctArea / WorksheetFunction.Pi) ^ 0.5
    End If
    
    'branch parameters
    If IsBranchCircular = True Then  'circular duct
    BranchDuctArea = WorksheetFunction.Pi * (DuctBranchWidth / 2) ^ 2
    Else 'rectangular duct
    BranchDuctArea = DuctBranchWidth * DuctBranchHeight
    DuctBranchWidth = (4 * BranchDuctArea / WorksheetFunction.Pi) ^ 0.5
    End If


' Step 2: Determine Um and Ub for branch and main ducts
Um = 0.001 * FlowRate / MainDuctArea
Ub = 0.001 * BranchFlowRate / BranchDuctArea

' Step 3: Determine ratios WidthRatio and m_U
WidthRatio = DuctWidth / DuctBranchWidth
VelocityRatio = Um / Ub

' Step 4: Determine rounding parameter, RD
RD = Radius / DuctBranchWidth

' Step 5: Determine Strouhal number, St
St = f * DuctBranchWidth / Ub

' Step 6: Determine radius correction term, RadiusCorrection
RadiusCorrection = (1 - RD / 0.15) * _
    (6.793 - 1.86 * Application.WorksheetFunction.Log10(St))

' Step 7: If turbulence is present, determine, TurbCorrection
If IsTurbulent = True Then
    TurbCorrection = -1.667 + 1.8 * VelocityRatio - 0.1333 * VelocityRatio ^ 2
Else
    TurbCorrection = 0
End If

' Step 8: Determine characteristic spectrum, Kj
Kj = -21.6 + 12.388 * VelocityRatio ^ 0.673 _
    - 16.482 * VelocityRatio ^ -0.303 * Application.WorksheetFunction.Log10(St) _
    - 5.047 * VelocityRatio ^ -0.254 * (Application.WorksheetFunction.Log10(St)) ^ 2

' Step 9: Determine the branch SWL, Lb
Lb = Kj + 10 * Application.WorksheetFunction.Log10(f / 63) _
    + 50 * Application.WorksheetFunction.Log10(3.28 * Ub) _
    + 10 * Application.WorksheetFunction.Log10(10.76 * BranchDuctArea) _
    + 10 * Application.WorksheetFunction.Log10(3.28 * DuctBranchWidth) _
    + RadiusCorrection + TurbCorrection

' Step 10: (Optional) Specify junction type, and determine main duct SWL, Lm.
'           If only the branch is desired, just return Lb
    If BranchRegen = True Then
    ElbowOrJunctionRegen_NEBB = Lb
    Else 'default to main
        Select Case JunctionType
        Case 1 'elbow
        Lm = Lb
        Case 2  '90 degree branch takeoff
        Lm = Lb + 20 * Application.WorksheetFunction.Log10(WidthRatio)
        Case 3 'T-Junction
        Lm = Lb + 3
        Case 4 'X-Junction
        Lm = Lb + 20 * Application.WorksheetFunction.Log10(WidthRatio) + 3
        End Select
    ElbowOrJunctionRegen_NEBB = Lm
    End If
    
End Function

'==============================================================================
' Name:     RegenDiffuser_NEBB
' Author:   AA
' Desc:     Calculates diffuser regenerated noise according to the NEBB method.
' Args:     fstr - Octave band centre frequency (Hz, string)
'           dP - pressure drop across a diffuser (Pa)
'           Q - volume flow rate (L/s)
'           DW - width (mm, double).
'               for generic diffusers this is duct width
'               for slot diffusers this is diffuser width
'           DH - height (mm, double)
'               for generic diffusers this is duct height
'               for slot diffusers this is diffuser height
'           Shape - shape of diffuser (integer)
'               1: Rectangular
'               2: Circular
' Comments: (1)
'==============================================================================
Function RegenDiffuser_NEBB(fStr As String, dP As Double, Q As Double, _
    Dw As Double, Dh As Double, Shape As Integer)

Dim f As Single     'usable freqency number
Dim rho As Single   'density of air (1.2 kg/m3)
Dim S As Double     'duct cross sectional area (m2)
Dim U As Double     'mean air flow velocity (m/s)
Dim z As Double     'normalised pressure drop coefficient
Dim Lw As Double    'overall Sound Power Level before correction
Dim fp As Double    'peak frequency (Hz)
Dim F1 As Double    'spectrum parameter
Dim F2 As Double    'spectrum parameter
Dim a As Double     'spectrum parameter
Dim c As Double     'shape of octave band sound spectrum

' GENERAL SETUP
f = freqStr2Num(fStr)
If freqStr2Num(fStr) < 63 Or freqStr2Num(fStr) > 8000 Then
    RegenDiffuser_NEBB = "-"
    Exit Function
End If
rho = 1.2
Dw = Dw / 1000
Dh = Dh / 1000

'Calculate cross-sectional area, S
S = Dw * Dh

'Calculate airflow velocity, U
U = 0.001 * Q / S

'Calculate normalised pressure drop coefficient, z
 z = 2 * dP / (rho * U ^ 2)

'Calculate overall SWL before correction, Lw
Lw = 10 * Application.WorksheetFunction.Log10(10.76 * S) _
    + 30 * Application.WorksheetFunction.Log10(z) _
    + 60 * Application.WorksheetFunction.Log10(3.28 * U) _
    - 31.3

'Calculate peak frequency, fp
fp = 160.1 * U

'Determine spectrum parameter, F1
F1 = GetArrayIndex_OCT(fStr, 1)

'Determine spectrum parameter, F2
Select Case fp
    Case Is < 44
        F2 = 0
    Case Is < 88
        F2 = 1
    Case Is < 177
        F2 = 2
    Case Is < 355
        F2 = 3
    Case Is < 710
        F2 = 4
    Case Is < 1420
        F2 = 5
    Case Is < 2840
        F2 = 6
    Case Is < 5680
        F2 = 7
    Case Is < 11360
        F2 = 8
End Select

'Calculate spectrum parameter, A
a = F1 - F2

'Calculate octave band sound spectrum, C
If Shape = 1 Then
    c = -11.82 - 0.15 * a - 1.13 * a ^ 2
ElseIf Shape = 2 Then
    c = -5.82 - 0.15 * a - 1.13 * a ^ 2
Else
    Debug.Print "IF error at calc of C"
    Exit Function
End If

'Calculate octave band Sound Power Level
RegenDiffuser_NEBB = Lw + c

End Function

'==============================================================================
' Name:     LouvreGrilleDirectivity_SRL
' Author:   PS
' Desc:     Inserts directivity for a louvre or grille according to Noise
'           Control In Building Services - Sound Research Laboratories
' Args:     fstr - Octave band centre frequency (Hz, string)
'           WidthOrHeight - of louvre in metres
'           Angle - Theta in degrees
' Comments: (1)
'==============================================================================
Function LouvreGrilleDirectivity_SRL(fStr As String, WidthOrHeight As String, _
    Angle As Long)

On Error GoTo closefile

Dim ReadStr As String
Dim i As Integer
Dim foundValue As Boolean
Dim SplitStr() As String
Dim ThetaCol As Integer
Dim fCol As Integer
Dim CurrentTable As String
Dim DirectivityAtZero As Double

Close #1

Call GetSettings

Open SRL_LG_DIRECTIVITY For Input As #1

foundValue = False
foundDir = False
fCol = GetArrayIndex_OCT(fStr, 1) 'array index, col 0 is W/H, 63Hz is col 1

'pick which column for theta
ThetaOptions = Array(0, 20, 40, 60, 80, 100, 120, 140)
ThetaCol = 999 'for checking
    For x = 0 To UBound(ThetaOptions)
        If ThetaOptions(x) = Angle Then ThetaCol = x
    Next x
    
    'catch mismatching columns or frequency range error
    If ThetaCol = 999 Or fCol < 1 Or fCol > 8 Then
    LouvreGrilleDirectivity_SRL = "-"
    Exit Function
    End If

    Do Until EOF(1) Or foundValue = True
    Line Input #1, ReadStr
    SplitStr = Split(ReadStr, vbTab, Len(ReadStr), vbTextCompare)

        If Left(SplitStr(0), 1) = "*" Then
        'extract the table number
        CurrentTable = Mid(SplitStr(0), 2, 7)
        Else
        'look for matching values
            If CurrentTable = "Table 1" Then
                If SplitStr(0) = WidthOrHeight Then
                'Debug.Print "Dir value: "; DirectivityAtZero
                DirectivityAtZero = SplitStr(fCol)
                foundDir = True
                End If
            ElseIf CurrentTable = "Table 2" Then
            'look for other angles
                If foundDir = True And DirectivityAtZero = SplitStr(0) Then
                    'Debug.Print "Value found! "; SplitStr(ThetaCol)
                    LouvreGrilleDirectivity_SRL = CDbl(SplitStr(ThetaCol))
                    foundValue = True
                End If
            Else
            End If
        End If
    Loop

If foundValue = False Then LouvreGrilleDirectivity_SRL = "-"
    
closefile: '<- On erroes, closes text file
Close #1
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     PutDuctAtten
' Author:   PS
' Desc:     Down-duct attenuation by ASHRAE or Reynolds methods
' Args:     None
' Comments: (1) Calls custom functions DuctAtten_ASHRAE and DuctAtten_Reynolds
'==============================================================================
Sub PutDuctAtten()

frmDuctAtten.Show

    If btnOkPressed = False Then End

    If T_BandType <> "oct" Then ErrorOctOnly
    
SetDescription "Duct Attenuation - " & ductMethod
ParameterUnmerge (Selection.Row)

    If ductMethod = "Reynolds" Then
    Cells(Selection.Row, T_ParamStart) = ductLiningThickness
    InsertComment "Duct lining thickness", T_ParamStart, False
    SetUnits "mm", T_ParamStart, 0
        If InStr(1, ductShape, "R", vbTextCompare) > 0 Then 'rectangular duct
        BuildFormula "DuctAtten_Reynolds(" & _
            T_FreqStartRng & "," & ductH & "," & ductW & "," & T_ParamRng(0) & _
            "," & T_ParamRng(1) & ")"
        ElseIf InStr(1, ductShape, "C", vbTextCompare) > 0 Then 'circular duct
        BuildFormula "DuctAttenCircular_Reynolds(" & _
            T_FreqStartRng & "," & ductH & "," & T_ParamRng(0) & _
            "," & T_ParamRng(1) & ")"
        Else
        ErrorUnexpectedValue
        End If
    ElseIf ductMethod = "ASHRAE" Then
    SetDataValidation T_ParamStart, "0 R,0 C,25 R,50 R,25 C,50 C"
    BuildFormula "DuctAtten_ASHRAE(" & _
        T_FreqStartRng & " ," & ductH & ", " & ductW & "," & T_ParamRng(0) & _
        "," & T_ParamRng(1) & ")"
    Cells(Selection.Row, T_ParamStart) = ductShape
    Cells(Selection.Row, T_ParamStart).NumberFormat = xlGeneral
    InsertComment TXT_HEAD & chr(10) & TXT_RAW, T_ParamStart, False
    ElseIf ductMethod = "SRL" Then
    SetDataValidation T_ParamStart, "R,C"
    'note, ductH hold diameter
    BuildFormula "DuctBendAtten_SRL(" & _
        T_FreqStartRng & " ," & ductH & "," & T_ParamRng(0) & _
        "," & T_ParamRng(1) & ")"
    Cells(Selection.Row, T_ParamStart) = Right(ductShape, 1)
    InsertComment TXT_HEAD & chr(10) & TXT_RAW, T_ParamStart, False
    Else
    'unrecognised method
    End If

        
'same for both methods
Cells(Selection.Row, T_ParamStart + 1) = ductL
SetUnits "m", T_ParamStart + 1, 1
SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutFlexDuctAtten
' Author:   PS
' Desc:     Down-duct attenuation flexxible duct, from ASHRAE
' Args:     None
' Comments: (1) Calls custom function FlexDuctAtten_ASHRAE
'==============================================================================
Sub PutFlexDuctAtten()

Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

SetDescription "Flex Duct - ASHRAE"

    If T_BandType <> "oct" Then ErrorOctOnly
    
BuildFormula "FlexDuctAtten_ASHRAE(" & _
    T_FreqStartRng & "," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"

SetTraceStyle "Input", True
ParameterUnmerge (Selection.Row)

Cells(Selection.Row, T_ParamStart) = 200
Cells(Selection.Row, T_ParamStart + 1) = 0.9
Cells(Selection.Row, T_ParamStart).NumberFormat = "0 Ø"
SetUnits "m", T_ParamStart + 1, 1

SetDataValidation T_ParamStart, "100,125,150,175,200,225,250,300,350,400"
SetDataValidation T_ParamStart + 1, "0.9,1.8,2.7,3.7"
        
End Sub


'==============================================================================
' Name:     DuctSplit
' Author:   PS
' Desc:     Energy split from duct junctions
' Args:     None
' Comments: (1) Three different modes (Area / Ratio / Percent), set in
'           frmDuctSplit
'==============================================================================
Sub PutDuctSplit()

frmDuctSplit.Show

    If btnOkPressed = False Then
    End
    End If

    If T_BandType <> "oct" Then ErrorOctOnly
    
    Select Case ductSplitType
    
    Case Is = "Area"
    ParameterUnmerge (Selection.Row)
    Cells(Selection.Row, T_ParamStart) = ductA1
    Cells(Selection.Row, T_ParamStart + 1) = ductA2
    SetUnits "m2", T_ParamStart, 1, T_ParamStart + 1
    
    BuildFormula "10*LOG(" & T_ParamRng(1) & _
        "/(" & T_ParamRng(1) & "+" & T_ParamRng(0) & "))"
        
    SetDescription "Duct Split: 10LOG(A2/(A1+A2))"
    
    Case Is = "Ratio"
    ParameterMerge (Selection.Row)
    Cells(Selection.Row, T_ParamStart) = ductA1
    Cells(Selection.Row, T_ParamStart).NumberFormat = "0"":1"""
    
    BuildFormula "10*LOG(1/" & T_ParamRng(0) & ")"
    
    SetDescription "Duct Split: 10LOG(1/R)"
    
    Case Is = "Percent"
    ParameterMerge (Selection.Row)
    Cells(Selection.Row, T_ParamStart).NumberFormat = "0%"
    Cells(Selection.Row, T_ParamStart) = ductA1
    
    BuildFormula "10*LOG(" & T_ParamRng(0) & ")"
    
    SetDescription "Duct Split: 10LOG(P)"
    
    End Select

SetTraceStyle "Input", True

End Sub

'==============================================================================
' Name:     PutERL
' Author:   PS
' Desc:     End reflection loss, from ASHRAE or NEBB method
' Args:     None
' Comments: (1) Calls custom functions ERL_ASHRAE and ERL_NEBB
'==============================================================================
Sub PutERL()
    
    'set default values if there's a ductAtten in the row above
    If Left(Cells(Selection.Row - 1, T_LossGainStart + 5).Formula, 10) = _
        "=DuctAtten" Then
    'Get parameters from row above
    frmERL.txtL.Value = GetDuctParameter(Cells(Selection.Row - 1, _
        T_LossGainStart + 5).Formula, "L") '1kHz band formula
    frmERL.txtW.Value = GetDuctParameter(Cells(Selection.Row - 1, _
        T_LossGainStart + 5).Formula, "W") '1kHz band formula
    End If

frmERL.Show

    If btnOkPressed = False Then End

SetDescription "End Reflection Loss - " & ERL_Mode
ParameterUnmerge (Selection.Row)

    If T_BandType <> "oct" Then ErrorOctOnly

    If ERL_Mode = "ASHRAE" Then
    BuildFormula "ERL_ASHRAE(" & _
        T_ParamRng(0) & "," & T_FreqStartRng & "," & T_ParamRng(1) & ")"
    ElseIf ERL_Mode = "NEBB" Then
    BuildFormula "ERL_NEBB(" & _
        T_ParamRng(0) & "," & T_FreqStartRng & "," & T_ParamRng(1) & ")"
    End If

Cells(Selection.Row, T_ParamStart) = ERL_Termination
Cells(Selection.Row, T_ParamStart).NumberFormat = xlGeneral
Cells(Selection.Row, T_ParamStart + 1).Value = ERL_Area
SetUnits "m2", T_ParamStart + 1, 2
    
SetTraceStyle "Input", True
SetDataValidation T_ParamStart, "Flush,Free"

End Sub

'==============================================================================
' Name:     PutElbowLoss
' Author:   PS
' Desc:     Loss through a duct bend/elbow, according to ASHRAE method
' Args:     None
' Comments: (1) Calls custom function ElbowLoss_ASHRAE
'==============================================================================
Sub PutElbowLoss()

    'get duct width from the row above, if possible
    If InStr(1, Cells(Selection.Row - 1, T_LossGainStart + 5).Formula, _
        "DuctAtten_ASHRAE", vbTextCompare) > 0 Or _
        InStr(1, Cells(Selection.Row - 1, T_LossGainStart + 5).Formula, _
        "DuctAtten_Reynolds", vbTextCompare) > 0 Then
    'Get parameters from row above
    frmElbowBend.txtW.Value = GetDuctParameter(Cells(Selection.Row - 1, _
        T_LossGainStart + 5).Formula, "W") '1kHz band formula
    End If

frmElbowBend.Show

    If btnOkPressed = False Then End

ParameterUnmerge (Selection.Row)

    If T_BandType <> "oct" Then ErrorOctOnly
    
Cells(Selection.Row, T_ParamStart) = ductW 'note: public variable
InsertComment "Duct width", T_ParamStart, False

Cells(Selection.Row, T_ParamStart + 1) = elbowLining
        
    'BUILD FORMULA
    If ductMethod = "ASHRAE" Then
    BuildFormula "ElbowLoss_ASHRAE(" & _
        T_FreqStartRng & "," & T_ParamRng(0) & ",""" & elbowShape & """," & _
        T_ParamRng(1) & ",""" & ElbowVanes & """)"
    SetDescription "Elbow Loss (" & elbowShape & ") - " & ductMethod
    ElseIf ductMethod = "SRL" Then
    '=DuctBendAtten_SRL(H$6,$N14,$O14)
    BuildFormula "DuctBendAtten_SRL(" & _
        T_FreqStartRng & "," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
    SetDescription "Elbow Loss " & " - " & ductMethod
    InsertComment TXT_HEAD & chr(10) & TXT_RAW, T_ParamStart, False
    Else
    msg = MsgBox("Method not recognised!")
    End If

'formatting
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = xlGeneral
SetUnits "mm", T_ParamStart, 0
SetTraceStyle "Input", True
        
SetDataValidation T_ParamStart + 1, "Lined,Unlined"

    'calc regenerated noise from element
    If CalcRegen = True Then
    SelectNextRow 'move down one row
    PutElbowRegen
    End If

End Sub

'==============================================================================
' Name:     PutSilencer
' Author:   PS
' Desc:     Loss through a splitter/silencer
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutSilencer()

'send selected row to public variable, you'll need it later
SolverRow = Selection.Row

frmSilencer.Show

If btnOkPressed = False Then End

    If T_BandType <> "oct" Then ErrorOctOnly
    
Cells(SolverRow, T_LossGainStart).ClearContents 'clear 31.5Hz octave band

    'check first frequency band
    If freqStr2Num(Range(T_FreqStartRng).Value) <> 31.5 Then ErrorFrequencyBand

    For col = 0 To 7 '8 columns
    Cells(SolverRow, T_LossGainStart + 1 + col).Value = SilencerIL(col)
    Next col
    
'parameter column
ParameterMerge (Selection.Row)
Cells(SolverRow, T_ParamStart).Value = SilLength
SetUnits "mm", T_ParamStart, 0

InsertComment SilSeries & chr(10) & "Length: " & SilLength & "mm" & _
    chr(10) & "Free Area: " & CStr(SilFA) & "%", T_ParamStart, True

SetDescription "Silencer: " & SilencerModel
Cells(Selection.Row, 1).Value = ChrW(167) 'or chrw(187) which is >>

SetTraceStyle "Silencer"

    'calc regenerated noise from element
    If CalcRegen = True Then
    SelectNextRow 'move down one row
    PutSilencerRegen
    End If

End Sub

'==============================================================================
' Name:     PutLouvres
' Author:   PS
' Desc:     Puts Louvre insertion loss into the sheet, with parameters in comment
' Args:     None
' Comments: (1) First frequency band must be 31.5Hz
'==============================================================================
Sub PutLouvres()

frmLouvres.Show

    If btnOkPressed = False Then End
    
    'check first frequency band
    If Range(T_FreqStartRng).Value <> 31.5 And _
       Range(T_FreqStartRng).Value <> "31.5*" Then ErrorFrequencyBand

'description
SetDescription "Acoustic Louvres: " & LouvreModel

If T_BandType <> "oct" Then ErrorOctOnly

    For col = 0 To 7 '8 columns
    Cells(Selection.Row, T_LossGainStart + 1 + col).Value = LouvreIL(col)
    Next col
'parameter cells
ParameterMerge (Selection.Row)
Cells(Selection.Row, T_ParamStart) = LouvreLength
SetUnits "mm", T_ParamStart, 0

InsertComment LouvreSeries & chr(10) & "Length: " & LouvreLength & "mm" & chr(10) _
    & "Free Area: " & LouvreFA, T_ParamStart

'apply style
SetTraceStyle "Silencer"

End Sub



'==============================================================================
' Name:     PutPlenumLoss
' Author:   PS
' Desc:     Puts in the formula for loss through a plenum
' Args:     None
' Comments: (1) This one's a doozy
'==============================================================================
Sub PutPlenumLoss()

frmPlenum.Show

If btnOkPressed = False Then End

    If T_BandType = "oct" Then 'oct or OCT
    BuildFormula "PlenumLoss_ASHRAE(" & _
        T_FreqStartRng & "," & PlenumL & "," & PlenumW & "," & PlenumH & "," _
        & DuctInL & "," & DuctInW & "," & DuctOutL & "," & DuctOutW & "," & _
        PlenumQ & "," & R_H & "," & r_v & ",""" & PlenumLiningType & """,""" & _
        UnlinedType & """,""" & PlenumWallEffectStr & """," & ApplyPlenumElbowEffect & _
        "," & PlenumPercentUnlined & ")"
    ElseIf T_BandType = "to" Then 'one thirds mode
    BuildFormula "PlenumLoss_ASHRAE(" & _
        T_FreqStartRng & "," & PlenumL & "," & PlenumW & "," & PlenumH & "," & _
        DuctInL & "," & DuctInW & "," & DuctOutL & "," & DuctOutW & "," & _
        PlenumQ & "," & R_H & "," & r_v & ",""" & PlenumLiningType & """,""" & _
        UnlinedType & """,""" & PlenumWallEffectStr & """," & ApplyPlenumElbowEffect & "," & _
        PlenumPercentUnlined & ",TRUE)"
    Else
        
    End If
'**!**
SetDescription "Plenum loss"

End Sub


'==============================================================================
' Name:     PutDuctBreakout
' Author:   NI
' Desc:     Puts in the formula for duct breakout
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutDuctBreakout()

frmBreakOut.Show

    If btnOkPressed = False Then End

ParameterUnmerge (Selection.Row)

    If T_BandType <> "oct" Then ErrorOctOnly

'values and formatting
Cells(Selection.Row, T_ParamStart).Value = ductL
SetUnits "m", T_ParamStart, 1
Cells(Selection.Row, T_ParamStart + 1).Value = DuctWallThickness
InsertComment "Duct wall thickness", T_ParamStart + 1, False
SetUnits "mm", T_ParamStart + 1, 1
SetTraceStyle "Input", True

BuildFormula "DuctBreakOut_NEBB(" & _
    T_FreqStartRng & "," & ductW & "," & ductH & "," & T_ParamRng(0) & "," & _
    MaterialDensity & "," & T_ParamRng(1) & ")"

    
SetDescription "Duct breakout"

End Sub


'==============================================================================
' Name:     PutDuctBreakin
' Author:   NI
' Desc:     Puts in the formula for duct breakin
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutDuctBreakIn()

frmBreakIn.Show

    If btnOkPressed = False Then End

ParameterUnmerge (Selection.Row)

    If T_BandType <> "oct" Then ErrorOctOnly

'values and formatting
Cells(Selection.Row, T_ParamStart).Value = ductL 'public variable
SetUnits "m", T_ParamStart, 1
InsertComment "Duct length", T_ParamStart, False
Cells(Selection.Row, T_ParamStart + 1).Value = DuctWallThickness 'public variable
SetUnits "mm", T_ParamStart + 1, 1
SetTraceStyle "Input", True
InsertComment "Duct wall thickness", T_ParamStart + 1, False

BuildFormula "DuctBreakIn_NEBB(" & _
    T_FreqStartRng & "," & ductW & "," & ductH & "," & T_ParamRng(0) & "," & _
    MaterialDensity & "," & T_ParamRng(1) & ")"

SetDescription "Duct break-in"

End Sub


'==============================================================================
' Name:     PutDuctDirectivity
' Author:   AN & PS
' Desc:     Puts in the formula for duct directivity
' Args:     None
' Comments: (1) From *Directivity Loss  at Duct Terminaton* by Daniel Potente,
'           Stepehen Gauld and Athol Day
'==============================================================================
Sub PutDuctDirectivity()

    If T_BandType <> "oct" Then ErrorOctOnly
    
BuildFormula "DuctDirectivity_PGD(" & _
    T_FreqStartRng & "," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
    
SetDescription "Duct directivity"

SetTraceStyle "Input", True

ParameterUnmerge (Selection.Row)

Cells(Selection.Row, T_ParamStart) = 0 'default to 0 degrees as angle
InsertComment "Directivity from normal, degrees", T_ParamStart, False
Cells(Selection.Row, T_ParamStart + 1) = 305 'default to 305mm as diameter
InsertComment "Duct width, mm", T_ParamStart + 1, False
Cells(Selection.Row, T_ParamStart).NumberFormat = "0°"
SetUnits "mm", T_ParamStart + 1, 0

'Validation for options from tables
SetDataValidation T_ParamStart, "0,15,30,45,50,75,90,105,120,135,150,165"
SetDataValidation T_ParamStart + 1, "305,610,914"
    
End Sub


'==============================================================================
' Name:     PutLouvreDirectivity
' Author:   PS
' Desc:     Puts in the formula for duct directivity
' Args:     None
' Comments: (1) From *Directivity Loss  at Duct Terminaton* by Daniel Potente,
'           Stepehen Gauld and Athol Day
'==============================================================================
Sub PutLouvreDirectivity()

    If T_BandType <> "oct" Then ErrorOctOnly
    
BuildFormula "LouvreGrilleDirectivity_SRL(" & T_FreqStartRng & "," & _
    T_ParamRng(0) & "," & T_ParamRng(1) & ")"
SetDescription "Louvre/grille directivity"


SetTraceStyle "Input", True

ParameterUnmerge (Selection.Row)

Cells(Selection.Row, T_ParamStart) = 0.5  'default to 0.5m
InsertComment "Width or height, m", T_ParamStart, False
Cells(Selection.Row, T_ParamStart + 1) = 0 'default to 0 degrees as angle
InsertComment "Directivity angle theta, degrees", T_ParamStart + 1, False
SetUnits "m", T_ParamStart, 1
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = "0°"

'Validation for options from tables
SetDataValidation T_ParamStart, "0.5,1,1.5,2,2.5,3.5,4.5,5.5,6,7.5,9,10.5,12,15"
SetDataValidation T_ParamStart + 1, "0,20,40,60,80,100,120,140"

    
End Sub

'==============================================================================
' Name:     PutSilencerRegen
' Author:   AA
' Desc:     Creates attenuator-regenerated noise row
' Args:     None
' Comments: (1) Inserts FantechAttenRegen() or NAPAttenRegen() function
'==============================================================================
Sub PutSilencerRegen()
Dim LineBefore As String

' General housekeeping and error checking
    If T_BandType <> "oct" Then ErrorOctOnly

'set default values if there's a ductAtten in the row above
    If InStr(1, Cells(Selection.Row - 1, T_Description).Formula, _
        "Silencer: ", vbTextCompare) > 0 Then
    'Get parameters from row above
    LineBefore = Cells(Selection.Row - 1, T_Description).Value
    frmSilencerRegen.txtTypeCode.Value = Right(LineBefore, Len(LineBefore) - _
        Len("Silencer: "))
    End If

frmSilencerRegen.Show

If btnOkPressed = False Then End

'description cell
SetDescription "Regen. noise - Attenuator - " & SilencerModel, Selection.Row, True
SetTraceStyle "Regen"

'parameter columns
ParameterUnmerge (Selection.Row)
SetTraceStyle "Input", True

'set units
Cells(Selection.Row, T_ParamStart + 1).Value = numModules
Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """n = ""0"
Cells(Selection.Row, T_ParamStart).Value = FlowRate

    If FlowUnitsM3ps = True Then
    SetUnits "m3ps", T_ParamStart, 1
    Else 'Litres per second
    SetUnits "lps", T_ParamStart, 0
    End If

    'build formula
    If RegenMode = "Fantech" Then
    BuildFormula "FantechAttenRegen(" & T_FreqStartRng & "," & T_ParamRng(0) & _
        "," & PFA & "," & ElementW & "," & ElementH & "," & T_ParamRng(1) & "," & _
        (Not FlowUnitsM3ps) & ")", True
    ElseIf RegenMode = "NAP" Then
    BuildFormula "NAPAttenRegen(" & T_FreqStartRng & "," & T_ParamRng(0) & "," & _
        PFA & "," & ElementW & "," & ElementH & ",""" & SilencerModel & """," & _
        (Not FlowUnitsM3ps) & ")", True
    Else
    End If

End Sub



'==============================================================================
' Name:     PutDamperRegen
' Author:   PS
' Desc:     Inserts regenerated noise for dampers
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutDamperRegen()

frmDamperRegen.Show
    
    If btnOkPressed = False Then End
    
'always start with the regen style, so input style can be applied after
SetTraceStyle "Regen"

'description
SetDescription "Regen. noise - Damper"

'parameter columns
ParameterMerge (Selection.Row)
SetTraceStyle "Input", True

'build formula
    If RegenMode = "NEBB" Then
        If FlowUnitsM3ps = True Then
        SetUnits "m3ps", T_ParamStart, 1
        Else 'litres
        SetUnits "lps", T_ParamStart, 0
        End If
    Cells(Selection.Row, T_ParamStart).Value = FlowRate
    BuildFormula "DamperRegen_NEBB(" & T_FreqStartRng & "," & T_ParamRng(0) & "," _
        & PressureLoss & "," & ElementH & "," & ElementW & "," & _
        DamperMultiBlade & "," & FlowUnitsM3ps & ")", True
    Else 'ASHRAE mode
    SetDataValidation T_ParamStart, "3.5,5.5,8.75,11,14.5"
    Cells(Selection.Row, T_ParamStart).Value = DuctVelocity
    SetUnits "mps", T_ParamStart, 1
    BuildFormula "RegenNoise_ASHRAE(" & T_FreqStartRng & ",""Damper"",""""," & _
        T_ParamRng(0) & ")", True
    End If
    
End Sub

'==============================================================================
' Name:     PutElbowRegen
' Author:   PS
' Desc:     Inserts regenerated noise for Elbows/Bends
' Args:     None
' Comments: (1)
'==============================================================================
Sub PutElbowRegen()

frmElbowRegen.Show
    
    If btnOkPressed = False Then End

'styling
SetTraceStyle "Regen"

'parameter columns
SetTraceStyle "Input", True

    'NEBB method
    If RegenMode = "NEBB" Then
    'description
    SetDescription "Regen. noise - Elbow/Bend (NEBB)"
    ParameterMerge (Selection.Row)
    
    'flowrate
        If FlowUnitsM3ps = True Then
        SetUnits "m3ps", T_ParamStart, 1
        Else 'litres
        SetUnits "lps", T_ParamStart, 0
        End If
    Cells(Selection.Row, T_ParamStart).Value = FlowRate
    InsertComment "Flowrate", T_ParamStart, False
    
        'for elbows with vanes
        If ElbowHasVanes = True Then
        BuildFormula "ElbowWithVanesRegen_NEBB(" & T_FreqStartRng & "," & _
            T_ParamRng(0) & "," & PressureLoss & "," & ElementW & "," & _
            ElementH & "," & BendCordLength & "," & ElbowNumVanes & "," & _
            FlowUnitsM3ps & ")", True
        Else 'no vanes
        '<---------TODO: Allow for circular duct options
        BuildFormula "ElbowOrJunctionRegen_NEBB(" & T_FreqStartRng & "," & _
            T_ParamRng(0) & "," & MainDuctCircular & "," & ElementW & "," & _
            ElementH & "," & T_ParamRng(0) & "," & BranchDuctCircular & "," _
            & ElementW & "," & ElementH & "," & ElbowRadius & "," & _
            IncludeTurbulence & ",1,1," & FlowUnitsM3ps & ")", True
        End If
    Else 'ASHRAE
    
    ParameterUnmerge (Selection.Row)
    
    'vanes
    SetDataValidation T_ParamStart, "Vanes, No Vanes"
        If ElbowHasVanes Then
        Cells(Selection.Row, T_ParamStart) = "Vanes"
        Else
        Cells(Selection.Row, T_ParamStart) = "No Vanes"
        End If
    Cells(Selection.Row, T_ParamStart).NumberFormat = "General"
    
    'velocity
    SetDataValidation T_ParamStart + 1, "10,15,17.5,20,25,30" 'allow all options
    Cells(Selection.Row, T_ParamStart + 1) = DuctVelocity
    SetUnits "mps", T_ParamStart + 1
    
    'build formula
    BuildFormula "RegenNoise_ASHRAE(" & T_FreqStartRng & ",""" & _
        regenNoiseElement & """," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
    SetDescription "Regen. noise - Elbow/Bend (ASHRAE)"
    End If


End Sub


'**************
'Code Graveyard
'**************



''==============================================================================
'' Name:     RegenNoiseNEBB
'' Author:   PS
'' Desc:     Inserts regenerated noise sound power of elements from NEBB
'' Args:     None
'' Comments: (1)
''==============================================================================
'Sub RegenNoiseNEBB()
'ErrorDoesNotExist
''dampers
'End Sub

''==============================================================================
'' Name:     PutRegenNoiseASHRAE
'' Author:   PS
'' Desc:     Inserts regenerated noise sound power of elements from ASHRAE
'' Args:     None
'' Comments: (1) May not be needed with new Regen Menu?
''==============================================================================
'Sub PutRegenNoiseASHRAE()
'
'frmRegenNoiseASHRAE.Show
'
'    If btnOkPressed = False Then End
'
'    If T_BandType <> "oct" Then ErrorOctOnly
'
'BuildFormula "RegenNoise_ASHRAE(" & T_FreqStartRng & ",""" & regenNoiseElement _
'    & """," & T_ParamRng(0) & "," & T_ParamRng(1) & ")", True
'
'SetTraceStyle "Input", True
'ParameterUnmerge (Selection.Row)
'
'    Select Case regenNoiseElement
'    Case Is = "Elbow"
'    Cells(Selection.Row, T_ParamStart) = "Vanes"
'    Cells(Selection.Row, T_ParamStart + 1) = "15"
'    Case Is = "Transition"
'    Cells(Selection.Row, T_ParamStart) = "Gradual"
'    Cells(Selection.Row, T_ParamStart + 1) = "15"
'    Case Is = "Damper"
'    Cells(Selection.Row, T_ParamStart) = ""
'    Cells(Selection.Row, T_ParamStart + 1) = "11"
'    End Select
'
'Cells(Selection.Row, T_ParamStart).NumberFormat = "General"
'SetUnits "mps", T_ParamStart + 1
'
'    With Cells(Selection.Row, T_ParamStart).Validation
'    .Delete
'        Select Case regenNoiseElement
'        Case Is = "Elbow"
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="Vanes, No Vanes"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .ShowInput = True
'        .ShowError = True
'        Case Is = "Transition"
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="Abrupt,Gradual"
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .ShowInput = True
'        .ShowError = True
'        Case Is = "Damper"
'        'no validation required, no nothing!
'        End Select
'    End With
'
'    With Cells(Selection.Row, T_ParamStart + 1).Validation
'    .Delete
'        Select Case regenNoiseElement
'        Case Is = "Elbow"
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="10,15,17.5,20,25,30"
'        Case Is = "Transition"
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="7.5,10,15,20"
'        Case Is = "Damper"
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:="3.5,5.5,8.75,11,14.5"
'        End Select
'    .IgnoreBlank = True
'    .InCellDropdown = True
'    .ShowInput = True
'    .ShowError = True
'    End With
'
'SetDescription "Regen. noise - " & regenNoiseElement
'
'End Sub

