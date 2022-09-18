Attribute VB_Name = "ISO9613"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public ISOFullElements(5) As Boolean 'boolean array for which elements are selected: Adiv Aatm Agr Abar Amisc
Public iso9613_d As Double
Public iso9613_d_ref As Double
Public iso9613_Temperature As Integer
Public iso9613_RelHumidity As Integer
Public iso9613_G_source As Double
Public iso9613_G_middle As Double
Public iso9613_G_receiver As Double
Public iso9613_SourceHeight As Double
Public iso9613_ReceiverHeight As Double
Public iso9613_SourceToBarrier As Double
Public iso9613_SrcToBarrierEdge As Double
Public iso9613_RecToBarrierEdge As Double
Public iso9613_BarrierHeight As Double
Public iso9613_BarrierHeightReceiverSide As Double
Public iso9613_DoubleDiffraction As Boolean
Public iso9613_BarrierThickness As Double
Public iso9613_MultiSource As Boolean

'==============================================================================
' Name:     ISO9613_A_div
' Author:   AN
' Desc:     Divergence correction
' Args:     Distance, D_ref (in metres)
' Comments: (1) the default reference distance is 1m
'==============================================================================
Function ISO9613_Adiv(Distance As Single, Optional D_ref As Single)
    If IsMissing(D_ref) Or D_ref = 0 Then
    D_ref = 1
    End If
'trace convention is negative numbers for losses
ISO9613_Adiv = (20 * Application.WorksheetFunction.Log(Distance / D_ref) + 11) * -1
End Function

'==============================================================================
' Name:     ISO9613_A_atm
' Author:   AN
' Desc:     Atmospheric correction
' Args:     fStr (octave frequency band), Distance (in metres), Temperature
'           (in degrees C), RelHumidity (value our of 100)
' Comments: (1) No interpolation, just straight out of the standard
'==============================================================================
Function ISO9613_Aatm(fStr As String, Distance As Double, Temperature As Integer, _
RelHumidity As Integer)

'These are the values from Table 2 of ISO9613
Dim TenSeventy() As Variant
Dim TwentySeventy() As Variant
Dim ThirtySeventy() As Variant
Dim FifteenTwenty() As Variant
Dim FifteenFifty() As Variant
Dim FifteenEighty() As Variant
Dim elem As Integer

TenSeventy = Array(0.1, 0.4, 1, 1.9, 3.7, 9.7, 32.8, 117)
TwentySeventy = Array(0.1, 0.3, 1.1, 2.8, 5, 9, 22.9, 76.6)
ThirtySeventy = Array(0.1, 0.3, 1, 3.1, 7.4, 12.7, 23.1, 59.3)
FifteenTwenty = Array(0.3, 0.6, 1.2, 2.7, 8.2, 28.2, 88.8, 202)
FifteenFifty = Array(0.1, 0.5, 1.2, 2.2, 4.2, 10.8, 36.2, 129)
FifteenEighty = Array(0.1, 0.3, 1.1, 2.4, 4.1, 8.3, 23.7, 82.8)

elem = GetArrayIndex_OCT(fStr)

    If elem = 999 Or elem = -1 Then 'catch error
    ISO9613_Aatm = "-"
    Else
        If Temperature = 10 And RelHumidity = 70 Then
        ISO9613_Aatm = TenSeventy(elem) * (Distance / 1000) * -1
        ElseIf Temperature = 20 And RelHumidity = 70 Then
        ISO9613_Aatm = TwentySeventy(elem) * (Distance / 1000) * -1
        ElseIf Temperature = 30 And RelHumidity = 70 Then
        ISO9613_Aatm = ThirtySeventy(elem) * (Distance / 1000) * -1
        ElseIf Temperature = 15 And RelHumidity = 20 Then
        ISO9613_Aatm = FifteenTwenty(elem) * (Distance / 1000) * -1
        ElseIf Temperature = 15 And RelHumidity = 50 Then
        ISO9613_Aatm = FifteenFifty(elem) * (Distance / 1000) * -1
        ElseIf Temperature = 15 And RelHumidity = 80 Then
        ISO9613_Aatm = FifteenEighty(elem) * (Distance / 1000) * -1
        Else 'catch all other cases
        ISO9613_Aatm = "-"
        End If
    End If
    
End Function

'==============================================================================
' Name:     ISO9613_A_gr
' Author:   AN
' Desc:     Ground Effect
' Args:     Dp (source to reciever distance as projected on to the ground plane)
'           ReceiverHeight (in metres), SourceHeight (in metres), Grec,Gsrc,Gmid
'           (Ground hardness of the source, reciever and middle zones, which is
'           between 0 And 1), q (defined for the purpose of adding to the Am)
' Comments: (1) As implemented in the standard.
'==============================================================================
Function ISO9613_Agr(fStr As String, SourceHeight As Double, ReceiverHeight As Double, _
dP As Double, Gsrc As Double, Grec As Double, Optional Gmid As Double)

Dim ahs As Double
Dim bhs As Double
Dim chs As Double
Dim dhs As Double

Dim ahr As Double
Dim bhr As Double
Dim chr As Double
Dim dhr As Double

Dim elem As Integer

Dim Q As Double

    If dP < 30 * (SourceHeight + ReceiverHeight) Then
      Q = 0
    Else
      Q = 1 - ((30 * (SourceHeight + ReceiverHeight)) / dP)
    End If
    
    If IsMissing(Gmid) Then Gmid = 0

'Source polynomials
ahs = 1.5 + (3 * Exp(-0.12 * ((SourceHeight - 5) ^ 2)) * (1 - Exp(-dP / 50))) + _
(5.7 * Exp(-0.09 * SourceHeight ^ 2) * (1 - Exp(-2.8 * dP ^ 2 * (10 ^ -6))))
bhs = 1.5 + ((8.6 * Exp(-0.09 * SourceHeight ^ 2)) * (1 - Exp(-dP / 50)))
chs = 1.5 + ((14 * Exp(-0.46 * SourceHeight ^ 2)) * (1 - Exp(-dP / 50)))
dhs = 1.5 + ((5 * Exp(-0.9 * SourceHeight ^ 2)) * (1 - Exp(-dP / 50)))

'Receiver polynomials
ahr = 1.5 + (3 * Exp(-0.12 * (ReceiverHeight - 5) * (ReceiverHeight - 5)) * _
(1 - Exp(-dP / 50))) + (5.7 * Exp(-0.09 * ReceiverHeight * ReceiverHeight) * _
(1 - Exp(-2.8 * dP * dP * (10 ^ -6))))
bhr = 1.5 + ((8.6 * Exp(-0.09 * ReceiverHeight ^ 2)) * (1 - Exp(-dP / 50)))
chr = 1.5 + ((14 * Exp(-0.46 * ReceiverHeight ^ 2)) * (1 - Exp(-dP / 50)))
dhr = 1.5 + ((5 * Exp(-0.9 * ReceiverHeight ^ 2)) * (1 - Exp(-dP / 50)))

elem = GetArrayIndex_OCT(fStr)

'Debug.Print "Gsrc: "; Gsrc
'Debug.Print "Gmid: "; Gmid
'Debug.Print "Grec: "; Grec
    If elem = 999 Or elem = -1 Then
    ISO9613_Agr = "-"
    Else
        Select Case elem
        Case 0 '63Hz
        ISO9613_Agr = -1.5 + -1.5 + (-3 * Q)
        Case 1 '125Hz
        ISO9613_Agr = (-1.5 + Gsrc * ahs) + (-1.5 + Grec * ahr) + (-3 * Q * (1 - Gmid))
        Case 2 '250Hz
        ISO9613_Agr = (-1.5 + Gsrc * bhs) + (-1.5 + Grec * bhr) + (-3 * Q * (1 - Gmid))
        Case 3 '500Hz
        ISO9613_Agr = (-1.5 + Gsrc * chs) + (-1.5 + Grec * chr) + (-3 * Q * (1 - Gmid))
        Case 4 '1kHz
        ISO9613_Agr = (-1.5 + Gsrc * dhs) + (-1.5 + Grec * dhr) + (-3 * Q * (1 - Gmid))
        Case 5 '2kHz
        ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * Q * (1 - Gmid))
        Case 6 '4kHz
        ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * Q * (1 - Gmid))
        Case 7 '8kHz
        ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * Q * (1 - Gmid))
        Case 999 'catch error
        ISO9613_Agr = "-"
        End Select
    End If

'NOTES The Ground Effect formulae return positive values for attenuation
'(as formula is "....-A").
    If IsNumeric(ISO9613_Agr) Then
    ISO9613_Agr = ISO9613_Agr * -1
    End If

End Function

'==============================================================================
' Name:     ISO9613_A_bar
' Author:   AN
' Desc:     Barrier Effect
' Args:     fStr - octave band centre frequency
'           SourceHeight - Source height above ground in metres
'           ReceiverHeight - Receiver height above ground in metres
'           SourceReceiverDistance - distance between them in metres
'           SourceBarrierDistance - distance from source to barrier
'           SrcDistanceEdge - distance from source to side edge of barrier
'           RecDistanceEdge - distance from receiver to side edge of barrier
'           HeightBarrierSource - Height of the barier on  the source side
'                               in metres. Note: For double barriers this is
'                               the first barrier.
'           DoubleDiffraction - switch to enable DD
'           BarrierThickness - distance between double barriers in metres
'           HeightBarrierReceiver - Height of the barier on  the receiver side
'                               in metres. Note: only exists if
'                               DoubleDiffraction = TRUE
'           GroundEffect - Agr from earlier step
' Comments: (1) distances are input as horizontal distances, with the
'           hypotenuse calculated during the function
'==============================================================================
Function ISO9613_Abar(fStr As String, SourceHeight As Double, ReceiverHeight As Double, _
SourceReceiverDistance As Double, SourceBarrierDistance As Double, _
SrcDistanceEdge As Double, RecDistanceEdge As Double, HeightBarrierSource As Double, _
Optional DoubleDiffraction As Boolean, Optional BarrierThickness As Double, _
Optional HeightBarrierReceiver As Double, Optional multisource As Boolean, _
Optional GroundEffect As Variant)

Dim Ctwo As Double
Dim Cthree As Single
Dim topEdge As Single
Dim verticalEdge As Single
Dim Dz As Double
Dim f As Double
Dim dss As Double 'distance from source to the first diffraction edge
Dim dsr As Double 'distance from the (second) diffraction edge to the receiver
Dim DistanceRecBarrier As Double 'distance from Receiver to the barrier (near side)
Dim a As Double 'a is the horizontal offset distance between the source and the receivers
Dim lambda As Double 'wavelength
Dim z As Double 'difference in path lengths of diffracted and direct sound in metres
Dim Kmet As Double
Dim d_standard As Double 'includes vertical component

If IsMissing(BarrierThickness) Or DoubleDiffraction = False Then BarrierThickness = 0

f = freqStr2Num(fStr)
lambda = (343) / f 'as defined in the method
a = Abs(SrcDistanceEdge - RecDistanceEdge)
    
    If IsNumeric(GroundEffect) = False Then
    ISO9613_Abar = "-"
    Exit Function
    End If

    'If the double diffraction is set as FALSE then there's only 1 top edge
    If DoubleDiffraction = False Then
      If HeightBarrierReceiver <> HeightBarrierSource Then
      HeightBarrierReceiver = HeightBarrierSource
      End If
    End If
    
DistanceRecBarrier = SourceReceiverDistance - SourceBarrierDistance - BarrierThickness

'Note to self: pythagoras rules!
dss = ((SourceBarrierDistance ^ 2) + ((HeightBarrierSource - SourceHeight) ^ 2)) ^ (1 / 2)
dsr = ((DistanceRecBarrier ^ 2) + ((HeightBarrierReceiver - ReceiverHeight) ^ 2)) ^ (1 / 2)
d_standard = (((SourceReceiverDistance ^ 2) + ((ReceiverHeight - SourceHeight) ^ 2)) ^ (1 / 2))


    If DoubleDiffraction = True And BarrierThickness > 0 Then
    Cthree = (1 + ((5 * lambda / BarrierThickness) * (5 * lambda / BarrierThickness))) / _
            (1 / 3 + ((5 * lambda / BarrierThickness) * (5 * lambda / BarrierThickness)))
    z = ((((dss + dsr + BarrierThickness) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d_standard
    
    'Here the case of double diffraction is considered it actually says if lambda is << e
    'so here we have considered half of the value
        If lambda < (BarrierThickness / 2) Then
        Cthree = 3
        End If
    Else 'double diffraction is false
    Cthree = 1
    z = ((((dss + dsr) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d_standard
    End If

'calculate the Kmet: the meteorological correction
    If z < 0 Or d_standard < 100 Then
    Kmet = 1
    Else
    Kmet = Exp((-1 / 2000) * (((dss * dsr * d_standard) / (2 * z)) ^ (1 / 2)))
    End If

''Print values for checking
'Debug.Print "dss: "; dss
'Debug.Print "dsr: "; dsr
'Debug.Print "kmet: "; kmet
'Debug.Print "lambda: "; lambda
'Debug.Print "cthree: "; cthree
'Debug.Print "----------------- ";

    'Difference values for multiple sources correction
    If multisource = True Then
    Ctwo = 40
    Else
    Ctwo = 20
    End If

'calculate barrier loss
Dz = 10 * Application.WorksheetFunction.Log(3 + ((Ctwo / lambda) * Cthree * Abs(z) * Kmet))

    'check for maximum value of Dz
    If DoubleDiffraction = True Then
        If Dz > 20 Then Dz = 20
    Else
        If Dz > 25 Then Dz = 25
    End If

    'Calculate final value
    If multisource = True Then
    ISO9613_Abar = Dz * -1 'trace convention is negative!
    Else
    ISO9613_Abar = (Dz - GroundEffect) * -1
    End If
    
End Function

'==============================================================================
' Name:     ISO9613_Cmet
' Author:
' Desc:
' Args:
' Comments: (1)
'==============================================================================
Function ISO9613_Cmet(fStr As String, hs, hr, dP, c0)
'TODO: One day we'll need this for something
End Function

'==============================================================================
' Name:     ISO9613_D_omega_alt
' Author:   AN
' Desc:     Calculates D_omega as per the alternative method (refer to standard)
' Args:     dp (distance source to receiver) hs (height of source)
'           hr (height of receiver)
' Comments: (1)
'==============================================================================
Function ISO9613_D_omega_alt(dP As Double, hs As Double, hr As Double)
ISO9613_D_omega_alt = 10 * Application.WorksheetFunction.Log(1 + (dP ^ 2 + (hs - hr) ^ 2) / _
                                                                (dP ^ 2 + (hs + hr) ^ 2))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     ISO_full
' Author:   PS
' Desc:     Switches on all ISO9613 elements and inserts all of them
' Args:     None
' Comments: (1) Calls Insert_ISO9613_CalcElements to build calc
'==============================================================================
Sub ISO9613_full()

frmISO9613.chkAatm.Value = True
frmISO9613.chkAdiv.Value = True
frmISO9613.chkAgr.Value = True
frmISO9613.chkAbar.Value = True

frmISO9613.Show

Insert_ISO9613_CalcElements

End Sub

'==============================================================================
' Name:     A_div
' Author:   PS
' Desc:     Switches on Adiv in form and sets up calc
' Args:     None
' Comments: (1) Inserts just teh divergence correction from ISO, without
'           calling the form
'           (2) TODO: Make it call the form?
'==============================================================================
Sub A_div()

SetDescription "ISO9613: A_div"
    
    If T_BandType <> "oct" Then ErrorOctOnly
    
Cells(Selection.Row, T_ParamStart).Value = 10
Cells(Selection.Row, T_ParamStart + 1).Value = 1
BuildFormula "=ISO9613_Adiv(" & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
SetUnits "m", T_ParamStart, 1, T_ParamStart + 1
SetTraceStyle "Input", True
InsertComment "Distance from source to receiver", T_ParamStart
InsertComment "Reference distance: 1m", T_ParamStart + 1

Cells(Selection.Row, T_ParamStart).Select 'move to parameter column to set value
End Sub

'==============================================================================
' Name:     A_atm
' Author:   PS
' Desc:     Switches on Aatm in form and sets up calc
' Args:     None
' Comments: (1) Calls Insert_ISO9613_CalcElements to build calc
'==============================================================================
Sub A_atm()

frmISO9613.chkAatm.Value = True
frmISO9613.chkAdiv.Value = False
frmISO9613.chkAbar.Value = False 'do this one first to avoid error messages
frmISO9613.chkAgr.Value = False

frmISO9613.Show

Insert_ISO9613_CalcElements

End Sub


'==============================================================================
' Name:     A_gr
' Author:   PS
' Desc:     Switches on Agr in form and sets up calc
' Args:     None
' Comments: (1) Calls Insert_ISO9613_CalcElements to build calc
'==============================================================================
Sub A_gr()

frmISO9613.chkAdiv.Value = False
frmISO9613.chkAatm.Value = False
frmISO9613.chkAgr.Value = True
frmISO9613.chkAbar.Value = False

frmISO9613.Show

Insert_ISO9613_CalcElements
    
End Sub

'==============================================================================
' Name:     A_bar
' Author:   PS
' Desc:     Switches on Abar in form and sets up calc
' Args:     None
' Comments: (1) Calls Insert_ISO9613_CalcElements to build calc
'==============================================================================
Sub A_bar()


frmISO9613.chkAdiv.Value = False
frmISO9613.chkAatm.Value = False
frmISO9613.chkAgr.Value = True 'Agr required for barrier calc
frmISO9613.chkAbar.Value = True

frmISO9613.Show

Insert_ISO9613_CalcElements

End Sub

'==============================================================================
' Name:     Insert_ISO9613_CalcElements
' Author:   PS
' Desc:     Builds ISO9613 calculation
' Args:     None
' Comments: (1)
'==============================================================================
Sub Insert_ISO9613_CalcElements()

If btnOkPressed = False Then End

If T_BandType <> "oct" Then ErrorOctOnly
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Adiv
    If ISOFullElements(0) = True Then
    SetDescription "ISO9613: A_div"
    BuildFormula "ISO9613_Adiv(" & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
    Cells(Selection.Row, T_ParamStart).Value = iso9613_d
    Cells(Selection.Row, T_ParamStart + 1).Value = iso9613_d_ref
    InsertComment "Distance from source to receiver", T_ParamStart
    InsertComment "Reference distance: 1m", T_ParamStart + 1
    
    SetTraceStyle "Input", True
    SetUnits "m", T_ParamStart, 1, T_ParamStart + 1
    
    SelectNextRow 'move down
    
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Aatm
    If ISOFullElements(1) = True Then
    SetSheetTypeControls
    SetDescription "ISO9613: A_atm"
        'if row above has _div, so we can use the same input for distance!
        If ISOFullElements(0) = True Then
        BuildFormula "ISO9613_Aatm(" & T_FreqStartRng & _
            "," & Cells(Selection.Row - 1, T_ParamStart).Address(False, True) & _
            "," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
        Else 'row above desn't have A_div, get from public variable
        '<----TODO extra inputs for mech sheets
        BuildFormula "ISO9613_Aatm(" & T_FreqStartRng & _
            "," & iso9613_d & "," & T_ParamRng(0) & "," & T_ParamRng(1) & ")"
        End If
        
    Cells(Selection.Row, T_ParamStart).Value = iso9613_Temperature
    Cells(Selection.Row, T_ParamStart + 1).Value = iso9613_RelHumidity
    Cells(Selection.Row, T_ParamStart).NumberFormat = """""0""°C """
    Cells(Selection.Row, T_ParamStart + 1).NumberFormat = "0 ""RH"""
    InsertComment "Temperature in degrees celcius", T_ParamStart
    InsertComment "Relative Humidity, %", T_ParamStart + 1
    
    'data validation
    SetDataValidation T_ParamStart, "10,15,20,30"
    SetDataValidation T_ParamStart + 1, "20,50,70,80"
        

    SetTraceStyle "Input", True
    SelectNextRow
    
    End If 'end of Aatm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Agr
    If ISOFullElements(2) = True Then
    SetDescription "ISO9613: A_gr"
        
        If ISOFullElements(0) = True Then 'Two rows above has A-div, so we can use the same input for distance!
        BuildFormula "ISO9613_Agr(" & T_FreqStartRng & "," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & ",$N" & _
        Selection.Row - 2 & ",$N" & Selection.Row & ",$O" & Selection.Row & "," & iso9613_G_middle & ")"
        Else
        BuildFormula "ISO9613_Agr(" & T_FreqStartRng & "," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & _
        iso9613_d & ",$N" & Selection.Row & ",$O" & Selection.Row & "," & iso9613_G_middle & ")"
        End If
        
        Cells(Selection.Row, T_ParamStart).Value = iso9613_G_source
        Cells(Selection.Row, T_ParamStart).NumberFormat = """Gs:"" 0.0"
        Cells(Selection.Row, T_ParamStart + 1).Value = iso9613_G_receiver
        Cells(Selection.Row, T_ParamStart + 1).NumberFormat = """Gr:"" 0.0"
        InsertComment "Ground effect in source region", T_ParamStart
        InsertComment "Ground effect in receiver region", T_ParamStart + 1
        
    SetTraceStyle "Input", True
    SelectNextRow 'move down
    
    End If 'end of Agr
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Abar
    If ISOFullElements(3) = True Then
    SetDescription "ISO9613: A_bar"
    Cells(Selection.Row, T_ParamStart).Value = iso9613_BarrierHeight
    SetUnits "m", T_ParamStart, 1
        If ISOFullElements(0) = True Then 'Three rows above has A-div, so we can use the same input for distance!
        BuildFormula "ISO9613_Abar(" & T_FreqStartRng & "," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & ",$N" & Selection.Row - 3 & "," & iso9613_SourceToBarrier & "," & _
            iso9613_SrcToBarrierEdge & "," & iso9613_RecToBarrierEdge & "," & "$N" & Selection.Row & "," & iso9613_DoubleDiffraction & "," & iso9613_BarrierThickness & "," & _
            iso9613_BarrierHeightReceiverSide & "," & iso9613_MultiSource & ",E$" & Selection.Row - 1 & ")"
        Else
        BuildFormula "ISO9613_Abar(" & T_FreqStartRng & "," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & iso9613_d & "," & iso9613_SourceToBarrier & "," & _
            iso9613_SrcToBarrierEdge & "," & iso9613_RecToBarrierEdge & "," & "$N" & Selection.Row & "," & iso9613_DoubleDiffraction & "," & iso9613_BarrierThickness & "," & _
            iso9613_BarrierHeightReceiverSide & "," & iso9613_MultiSource & ",E$" & Selection.Row - 1 & ")"
        End If
    
    SetTraceStyle "Input", True
    
    SelectNextRow 'move down
    
    End If 'end of Abar
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Amisc: TODO one day????
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
End Sub


