Attribute VB_Name = "NoiseFunctions"
Public ductL As Integer
Public ductW As Integer
Public ductShape As String
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
Public SilencerModel As String
Public SilencerIL() As Double
Public SilLength As Double
Public SilFA As Double
Public SolverRow As Integer

''''''''''
'FUNCTIONS
''''''''''
Private Function AirAbsorb(freq As String, Distance As Integer, temp As Integer)
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

Private Function DuctAtten(freq As String, Distance As Integer)
    Select Case freq
    Case Is = "63"
    DuctAtten = -0.05 * Distance
    Case Is = "125"
    DuctAtten = -0.21 * Distance
    Case Is = "250"
    DuctAtten = -0.85 * Distance
    Case Is = "500"
    DuctAtten = -3.4 * Distance
    Case Is = "1k"
    DuctAtten = -9.4 * Distance
    Case Is = "2k"
    DuctAtten = -13.2 * Distance
    Case Is = "4k"
    DuctAtten = -8.6 * Distance
    Case Is = "8k"
    DuctAtten = -4# * Distance
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
    
    For i = LBound(freqTitles) To UBound(freqTitles)
        If freq = freqTitles(i) Then
        ArrayIndex = i
        End If
    Next i
    
    If ArrayIndex <> 999 Then 'error
    AWeightCorrections = dBCAdjustment(ArrayIndex)
    Else
    AWeightCorrections = "-"
    End If
    
End Function


Function GetASHRAE(freq As String, L As Long, W As Long, DuctType As String, Distance As Double)
'On Error GoTo closefile
Dim ReadStr() As String
Dim i As Integer
Dim splitStr() As String
Dim splitVal() As Double
Dim CurrentType As String
Dim InputArea As Double
'Get Array from text
Close #1

Call GetSettings

Open ASHRAE_DUCT_TXT For Input As #1  'global

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    
    splitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(splitStr(0), 1) <> "*" Then
        
            'convert to values
            For Col = 0 To UBound(splitStr)
            If splitStr(Col) <> "" Then
            ReDim Preserve splitVal(Col)
            splitVal(Col) = CDbl(splitStr(Col))
            End If
            Next Col
            
            ReDim Preserve splitVal(Col + 1)
            
                If Right(DuctType, 1) = "R" Then 'RECTANGULAR DUCT
                ReadArea = splitVal(0) * splitVal(1)
                InputArea = L * W
                ElseIf Right(DuctType, 1) = "C" Then 'CIRCULAR DUCT
                ReadArea = WorksheetFunction.Pi * ((splitVal(0) / 2) ^ 2)
                InputArea = WorksheetFunction.Pi * ((L / 2) ^ 2)
                Else
                'msg = MsgBox("UNKNOWN TYPE", vbOKOnly, "You done f**ked up now.")
                End If
            
            If InputArea <= ReadArea And CurrentType = DuctType Then
            'Debug.Print "AREA found - line " & i
                'select correct frequency band
'                    For x = 0 To 9
'                    Debug.Print splitVal(x)
'                    Next x
                
                Select Case freq
                Case Is = "63"
                    If Right(CurrentType, 1) = "R" Then 'RECTANGULAR DUCT
                    GetASHRAE = splitVal(2) * -Distance / 2
                    ElseIf Right(CurrentType, 1) = "C" Then 'CIRCULAR DUCT
                    GetASHRAE = splitVal(1) * -Distance
                    End If
                Case Is = "125"
                GetASHRAE = splitVal(2) * -Distance
                Case Is = "250"
                GetASHRAE = splitVal(3) * -Distance
                Case Is = "500"
                GetASHRAE = splitVal(4) * -Distance
                Case Is = "1k"
                GetASHRAE = splitVal(5) * -Distance
                Case Is = "2k"
                GetASHRAE = splitVal(6) * -Distance
                Case Is = "4k"
                GetASHRAE = splitVal(7) * -Distance
                Case Else
                GetASHRAE = ""
                End Select
                
                'Floor the value, duct attenuation shouldn't be above 40dB
                If GetASHRAE < -40 Then
                GetASHRAE = -40
                End If
                
            found = True '<-this will end the loop
            End If
            
            
        Else '* is the type identifier
        'ReDim Preserve SplitVal(1)
        CurrentType = Right(splitStr(0), Len(splitStr(0)) - 1)
        'Debug.Print "TYPE: " & currentType
        End If
        
    i = i + 1
    Loop
    
closefile: '<-on errors, closes text file
Close #1
End Function

Function GetFlexDuct(freq As String, dia As Integer, L As Double)
On Error GoTo closefile
Dim ReadStr() As String
Dim i As Integer
Dim splitStr() As String
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
    
    splitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
        If Left(splitStr(0), 1) <> "*" Then 'titles
        
            'convert to values
            For Col = 0 To UBound(splitStr)
                If splitStr(Col) <> "" Then
                ReDim Preserve splitVal(Col)
                splitVal(Col) = CDbl(splitStr(Col))
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

Function GetERL(TerminationType As String, freq As String, DuctArea As Double)
Dim dia As Double
Dim A1 As Double
Dim A2 As Double
Dim f As Double
dia = (4 * DuctArea / Application.WorksheetFunction.Pi) ^ 0.5 'eqn 11
f = freqStr2Num(freq)
c0 = 343
    'table 28 of ASHRAE
    If TerminationType = "Flush" Then
    A1 = 0.7
    A2 = 2
    ElseIf TerminationType = "Free" Then
    A1 = 1
    A2 = 2
    End If
GetERL = -10 * Application.WorksheetFunction.Log10(1 + ((A1 + c0) / (f * dia * Application.WorksheetFunction.Pi)) ^ A2)
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
    
    splitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(splitStr(0), 1) <> "*" Then
            
            If CurrentType = Element Then 'elbow, damper, or transition
                If splitStr(0) = Condition And CDbl(splitStr(1)) = Velocity Then 'vanes/no vanes
                
                'convert to values
                For Col = 1 To UBound(splitStr)
                If splitStr(Col) <> "" Then
                ReDim Preserve splitVal(Col)
                splitVal(Col) = CDbl(splitStr(Col))
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
                    Case Else
                    GetRegenNoise = ""
                    End Select
                    
                End If
            End If
            
            'ReDim Preserve splitVal(Col + 1)
            
        Else '* is the type identifier
        CurrentType = Right(splitStr(0), Len(splitStr(0)) - 1)
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

Function GetRoomLoss(fstr As String, L As Double, W As Double, H As Double, roomType As String)
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
    End Select
        
    S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)
    alpha_av = ((L * W * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) + (W * H * alpha(bandIndex) * 2)) / S_total
    Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)
    'Debug.Print "Room Contant " Rc
        If Rc <> 0 Then
        GetRoomLoss = 10 * Application.WorksheetFunction.Log10(4 / Rc)
        Else
        GetRoomLoss = 0
        End If
End Function


Function GetRoomLossRT(fstr As String, L As Double, W As Double, H As Double, RT_Type As String)
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
    End Select
    
    S_total = (L * W * 2) + (L * H * 2) + (W * H * 2)
    alpha_av = ((L * W * alpha(bandIndex) * 2) + (L * H * alpha(bandIndex) * 2) + (W * H * alpha(bandIndex) * 2)) / S_total
    Rc = (S_total * alpha(bandIndex)) / (1 - alpha_av)

        If Rc <> 0 Then
        GetRoomLossRT = 10 * Application.WorksheetFunction.Log10(4 / Rc)
        Else
        GetRoomLossRT = 0
        End If

End Function

Function GetElbowLoss(fstr As String, W As Double, elbowShape As String, DuctLining As String, VaneType As String)
Dim Unlined() As Variant
Dim Lined() As Variant
Dim RadiusBend() As Variant
Dim freq As Double
Dim FW As Double
Dim ArrayIndex As Integer
Dim linedDuct As Boolean
Dim vanes As Boolean

    If DuctLining = "Lined" Then
    linedDuct = True
    ElseIf DuctLining = "Unlined" Then
    linedDuct = False
    End If
    
    If VaneType = "Vanes" Then
    vanes = True
    ElseIf VaneType = "No Vanes" Then
    vanes = False
    End If
    

Unlined = Array(0, -1, -5, -8, -4, -6) 'table 22 of ASHRAE
Lined = Array(0, -1, -6, -11, -10, -10)

UnlinedV = Array(0, -1, -4, -6, -4) 'table 24 of ASHRAE
LinedV = Array(0, -1, -4, -7, -7)

RadiusBend = Array(0, -1, -2, -3) 'table 23 of ASHRAE

freq = freqStr2Num(fstr)
FW = (freq / 1000) * W

    Select Case elbowShape
    Case Is = "Square"
        If vanes = False Then
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

Function freqStr2Num(fstr)
    Select Case fstr
    Case Is = "31.5"
    freqStr2Num = 31.5
    Case Is = "40"
    freqStr2Num = 40
    Case Is = "50"
    freqStr2Num = 50
    Case Is = "63"
    freqStr2Num = 63
    Case Is = "80"
    freqStr2Num = 80
    Case Is = "100"
    freqStr2Num = 100
    Case Is = "125"
    freqStr2Num = 125
    Case Is = "160"
    freqStr2Num = 160
    Case Is = "200"
    freqStr2Num = 200
    Case Is = "250"
    freqStr2Num = 250
    Case Is = "315"
    freqStr2Num = 315
    Case Is = "400"
    freqStr2Num = 400
    Case Is = "500"
    freqStr2Num = 500
    Case Is = "630"
    freqStr2Num = 630
    Case Is = "800"
    freqStr2Num = 800
    Case Is = "1k"
    freqStr2Num = 1000
    Case Is = "1.25k"
    freqStr2Num = 1250
    Case Is = "1.6k"
    freqStr2Num = 1600
    Case Is = "2k"
    freqStr2Num = 2000
    Case Is = "2.5k"
    freqStr2Num = 2500
    Case Is = "3.15k"
    freqStr2Num = 3150
    Case Is = "4k"
    freqStr2Num = 4000
    Case Is = "5k"
    freqStr2Num = 5000
    Case Is = "6.3k"
    freqStr2Num = 6300
    Case Is = "8k"
    freqStr2Num = 8000
    Case Is = "10k"
    freqStr2Num = 10000
    Case Else
    freqStr2Num = 0
    End Select
End Function

'''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''
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

UserInputFormat_ParamCol (SheetType)

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

UserInputFormat_ParamCol (SheetType)

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
UserInputFormat_ParamCol (SheetType)
Call ParameterUnmerge(Selection.Row, SheetType)
Cells(Selection.Row, ParamCol1) = 150
Cells(Selection.Row, ParamCol2) = 20
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
Cells(Selection.Row, ParamCol2).NumberFormat = "0""" & Chr(176) & "C"""
End Sub

Sub DuctAttenuation(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "Duct Attenuation"

    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=DuctAtten(E$6,$N" & Selection.Row & ")"
    ParamCol1 = 14
    ParamCol2 = 15
    
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=DuctAtten(E$6,$Z" & Selection.Row & ")"
    ParamCol1 = 26
    ParamCol2 = 27
    End If

ExtendFunction (SheetType)
Cells(Selection.Row, ParamCol1) = 1
Cells(Selection.Row, ParamCol2) = 20
Cells(Selection.Row, ParamCol1).NumberFormat = "0 ""m"""
Cells(Selection.Row, ParamCol2).NumberFormat = "0""" & Chr(176) & "C"""
UserInputFormat_ParamCol (SheetType)
End Sub

Sub ASHRAE_DUCT(SheetType As String)

Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmDuctAtten.Show

    If btnOkPressed = False Then
    End
    End If

Cells(Selection.Row, 2).Value = "Duct Attenuation-ASHRAE"

    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=GetASHRAE(E$6," & ductL & ", " & ductW & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
    'GetASHRAE(Freq As String, L As Integer, W As Integer, DuctType As String)
    ParamCol1 = 14
    ParamCol2 = 15
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=GetASHRAE(E$6,$Z" & Selection.Row & ",$AA" & Selection.Row & ")"
    ParamCol1 = 26
    ParamCol2 = 27
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
Call ParameterUnmerge(Selection.Row, SheetType)
Cells(Selection.Row, ParamCol1) = ductShape 'from public variable
Cells(Selection.Row, ParamCol2) = 1
Cells(Selection.Row, ParamCol1).NumberFormat = xlGeneral
Cells(Selection.Row, ParamCol2).NumberFormat = "0.0 ""m"""

With Cells(Selection.Row, ParamCol1).Validation
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

End Sub

Sub FlexDuct(SheetType As String)

Dim ParamCol1 As Integer
Dim ParamCol2 As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "Flex Duct-ASHRAE"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=GetFlexDuct(E$6,$N" & Selection.Row & ",$O" & Selection.Row & ")"
    ParamCol1 = 14
    ParamCol2 = 15
    ElseIf Left(SheetType, 2) = "TO" Then
'    Cells(Selection.Row, 5).Value = "=GetASHRAE(E$6,$Z" & Selection.Row & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
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
    Cells(Selection.Row, 14).NumberFormat = "0 ""m" & Chr(178) & """"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 2
    Cells(Selection.Row, 26).NumberFormat = "0 ""m" & Chr(178) & """"
    
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
End Sub

Sub DuctSplit(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmDuctAreas.Show

    If btnOkPressed = False Then
    End
    End If

Call ParameterUnmerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    'Debug.Print "=10*LOG($O" & Selection.Row & "/($O" & Selection.Row & "+$N" & Selection.Row & "))"
    Cells(Selection.Row, 5).Value = "=10*LOG($O" & Selection.Row & "/($O" & Selection.Row & "+$N" & Selection.Row & "))"
    Cells(Selection.Row, 14) = ductA1
    Cells(Selection.Row, 14).NumberFormat = "0.0""m" & Chr(178) & """"
    Cells(Selection.Row, 15) = ductA2
    Cells(Selection.Row, 15).NumberFormat = "0.0""m" & Chr(178) & """"

    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
'    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 26).Address(False, True) & ")"
'    Cells(Selection.Row, 26) = 0.5
'    Cells(Selection.Row, 26).NumberFormat = "0 %"

    Else
    SheetTypeUnknownError
    End If
Cells(Selection.Row, 2).Value = "Duct Split: 10LOG(A2/(A1+A2))"
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
End Sub

Sub ERLoss(SheetType As String)
Dim ParamCol1 As Integer
Dim ParamCol2 As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Cells(Selection.Row, 2).Value = "End Reflection Loss"
Call ParameterUnmerge(Selection.Row, SheetType)
    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetERL($N" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$O" & Selection.Row & ")"
    ParamCol1 = 14
    ParamCol2 = 15
'    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
'    Cells(Selection.Row, 5).Value = "=GetERL($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    Else
    SheetTypeUnknownError
    End If
    
Cells(Selection.Row, ParamCol1) = "Flush"
Cells(Selection.Row, ParamCol1).NumberFormat = xlGeneral
'Debug.Print Cells(Selection.Row - 1, 10).Formula
    If InStr(1, Cells(Selection.Row - 1, 10).Formula, "GetASHRAE", vbTextCompare) > 0 Then
    Cells(Selection.Row, ParamCol2) = GetDuctArea(Cells(Selection.Row - 1, 10).Formula) '1kHz band formula
    Else
    Cells(Selection.Row, ParamCol2) = "=0.5*0.5"
    End If
Cells(Selection.Row, ParamCol2).NumberFormat = "0.0""m" & Chr(178) & """"
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)

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

Function GetDuctArea(inputStr As String)
Dim splitStr() As String
Dim L As Double
Dim W As Double
splitStr = Split(inputStr, ",", Len(inputStr), vbTextCompare)
L = CDbl(splitStr(1))
W = CDbl(splitStr(2))
GetDuctArea = (L / 1000) * (W / 1000) 'because millimetres
End Function

Sub ElbowLoss(SheetType As String)
Dim ParamCol As Integer

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
    ParamCol = 15
'    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
'    Cells(Selection.Row, 5).Value = "=GetERL($Z" & Selection.Row & "," & Cells(6, 5).Address(True, False) & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    Else
    SheetTypeUnknownError
    End If
    
    ExtendFunction (SheetType)
    UserInputFormat_ParamCol (SheetType)
    
    With Cells(Selection.Row, ParamCol).Validation
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

Set CheckRng = Cells(SolverRow, 14)

msg = MsgBox("This tool is in beta and may not function as intended.", vbOKOnly, "WARNING!")
frmSilencer.Show

If btnOkPressed = False Then End

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(SolverRow, 5).ClearContents 'clear 31.5Hz octave band
        For Col = 0 To 7 '8 columns
        Cells(SolverRow, 6 + Col).Value = SilencerIL(Col)
        Next Col
        Cells(SolverRow, 14).Value = SilLength
        Cells(SolverRow, 14).NumberFormat = "0 ""mm"""
            If CheckRng.Comment Is Nothing Then
            Else
            CheckRng.Comment.Delete
            End If
            CheckRng.AddComment "Free Area: " & CStr(SilFA) & "%"
        
    Else
    SheetTypeUnknownError
    End If
Cells(Selection.Row, 2).Value = "Silencer Model: " & SilencerModel

End Sub


Sub RoomLoss(SheetType As String)
Dim splitStr() As String
Dim ParamCol As Integer
On Error GoTo errorcatch:
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'populate frmRoomLoss
splitStr = Split(Cells(Selection.Row, 5).Formula, ",", Len(Cells(Selection.Row, 5).Formula), vbTextCompare)
roomL = CLng(splitStr(1))
roomW = CLng(splitStr(2))
roomH = CLng(splitStr(3))
roomType = Cells(Selection.Row, 14).Value
Call frmRoomLoss.Populate_frmRoomLoss

errorcatch:

frmRoomLoss.Show

Call ParameterMerge(Selection.Row, SheetType)

    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetRoomLoss(" & Cells(6, 5).Address(True, False) & "," & roomL & "," & roomW & "," & roomH & ",$N" & Selection.Row & ")"
    Cells(Selection.Row, 14) = roomType
    ParamCol = 14
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)

With Cells(Selection.Row, ParamCol).Validation
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

Call ParameterMerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Or Left(SheetType, 2) = "TO" Then 'OCT, OCTA, TO, or TOA
    Cells(Selection.Row + 1, 5).Value = "=10*LOG(4/E" & Selection.Row & ")" 'next row down
    Else
    SheetTypeUnknownError
    End If

UserInputFormat (SheetType)
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
Dim splitStr() As String
Dim ParamCol As Integer
On Error GoTo errorcatch:
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

'populate frmRoomLoss
splitStr = Split(Cells(Selection.Row, 5).Formula, ",", Len(Cells(Selection.Row, 5).Formula), vbTextCompare)
roomL = CLng(splitStr(1))
roomW = CLng(splitStr(2))
roomH = CLng(splitStr(3))
roomType = Cells(Selection.Row, 14).Value
Call frmRoomLoss.Populate_frmRoomLoss

errorcatch:

frmRoomLossRT.Show

Call ParameterMerge(Selection.Row, SheetType)

    If btnOkPressed = False Then
    End
    End If

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=GetRoomLossRT(" & Cells(6, 5).Address(True, False) & "," & roomL & "," & roomW & "," & roomH & ",$N" & Selection.Row & ")"
    Cells(Selection.Row, 14) = roomType
    ParamCol = 14
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)

With Cells(Selection.Row, ParamCol).Validation
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
'    Cells(Selection.Row, 5).Value = "=GetASHRAE(E$6,$Z" & Selection.Row & ",$AA" & Selection.Row & ")"
'    ParamCol1 = 26
'    ParamCol2 = 27
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
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
Dim endrw As Integer
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

endrw = Selection.Row - 1 'for reveberant sum

'distance correction
Distance (SheetType)
Cells(Selection.Row, 14).Value = 1  'COL N ; 1m by default
'move down
Cells(Selection.Row + 1, Selection.Column).Select
    'Sum direct
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & StartRw + 1 & ":E" & Selection.Row - 1 & ")"
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "Direct component"

'move cursor
Cells(Selection.Row + 1, Selection.Column).Select

'Room loss
RoomLoss (SheetType)

'move down
Cells(Selection.Row + 1, Selection.Column).Select

    'Sum reverb
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & StartRw + 1 & ":E" & endrw & ",E" & Selection.Row - 1 & ")"
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "Reverberant component"

'move down
Cells(Selection.Row + 1, Selection.Column).Select

    'Sum TOTAL
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SPLSUM(E" & Selection.Row - 1 & ",E" & Selection.Row - 3 & ")"
    ElseIf Left(SheetType, 2) = "TO" Then
    'Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    End If
    
ExtendFunction (SheetType)
Cells(Selection.Row, 2).Value = "TOTAL"

'Colour highlight
Range(Cells(Selection.Row, 2), Cells(Selection.Row, 18)).Font.Color = RGB(68, 114, 196)

End Sub
