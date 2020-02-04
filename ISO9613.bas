Attribute VB_Name = "ISO9613"
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

Function ISO9613_Adiv(Distance As Single, D_ref As Single) 'maybe we don't need this?????
ISO9613_Adiv = -20 * Application.WorksheetFunction.Log(Distance / D_ref) + 11
End Function


Function ISO9613_Aatm(fstr As String, Distance As Double, Temperature As Integer, RelHumidity As Integer)

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

elem = GetOctaveColumnIndex(fstr)

    If elem = 999 Then 'catch error
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


Function ISO9613_Agr(fstr As String, SourceHeight As Double, ReceiverHeight As Double, dp As Double, Gsrc As Double, Grec As Double, Optional Gmid As Double)

'Dp- source to reciever distance as projected on to the ground plane
'ReceiverHeight- height of the reciever
'SourceHeight- height of the source
'Grec,Gsrc,Gmid - Ground type of the source, reciever and middle ground which is between 0 and 1
'q - defined for the purpose of adding to the Am

Dim ahs As Double
Dim bhs As Double
Dim chs As Double
Dim dhs As Double

Dim ahr As Double
Dim bhr As Double
Dim chr As Double
Dim dhr As Double

Dim elem As Integer

Dim q As Double

    If dp < 30 * (SourceHeight + ReceiverHeight) Then
      q = 0
    Else
      q = 1 - ((30 * (SourceHeight + ReceiverHeight)) / dp)
    End If
    
    If IsMissing(Gmid) Then Gmid = 0

'Source polynomials
ahs = 1.5 + (3 * Exp(-0.12 * ((SourceHeight - 5) ^ 2)) * (1 - Exp(-dp / 50))) + (5.7 * Exp(-0.09 * SourceHeight ^ 2) * (1 - Exp(-2.8 * dp ^ 2 * (10 ^ -6))))
bhs = 1.5 + ((8.6 * Exp(-0.09 * SourceHeight ^ 2)) * (1 - Exp(-dp / 50)))
chs = 1.5 + ((14 * Exp(-0.46 * SourceHeight ^ 2)) * (1 - Exp(-dp / 50)))
dhs = 1.5 + ((5 * Exp(-0.9 * SourceHeight ^ 2)) * (1 - Exp(-dp / 50)))

'Receiver polynomials
ahr = 1.5 + (3 * Exp(-0.12 * (ReceiverHeight - 5) * (ReceiverHeight - 5)) * (1 - Exp(-dp / 50))) + (5.7 * Exp(-0.09 * ReceiverHeight * ReceiverHeight) * (1 - Exp(-2.8 * dp * dp * (10 ^ -6))))
bhr = 1.5 + ((8.6 * Exp(-0.09 * ReceiverHeight ^ 2)) * (1 - Exp(-dp / 50)))
chr = 1.5 + ((14 * Exp(-0.46 * ReceiverHeight ^ 2)) * (1 - Exp(-dp / 50)))
dhr = 1.5 + ((5 * Exp(-0.9 * ReceiverHeight ^ 2)) * (1 - Exp(-dp / 50)))

elem = GetOctaveColumnIndex(fstr)

'Debug.Print "Gsrc: "; Gsrc
'Debug.Print "Gmid: "; Gmid
'Debug.Print "Grec: "; Grec

    Select Case elem
    Case 0 '63Hz
    ISO9613_Agr = -1.5 + -1.5 + (-3 * q)
    Case 1 '125Hz
    ISO9613_Agr = (-1.5 + Gsrc * ahs) + (-1.5 + Grec * ahr) + (-3 * q * (1 - Gmid))
    Case 2 '250Hz
    ISO9613_Agr = (-1.5 + Gsrc * bhs) + (-1.5 + Grec * bhr) + (-3 * q * (1 - Gmid))
    Case 3 '500Hz
    ISO9613_Agr = (-1.5 + Gsrc * chs) + (-1.5 + Grec * chr) + (-3 * q * (1 - Gmid))
    Case 4 '1kHz
    ISO9613_Agr = (-1.5 + Gsrc * dhs) + (-1.5 + Grec * dhr) + (-3 * q * (1 - Gmid))
    Case 5 '2kHz
    ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * q * (1 - Gmid))
    Case 6 '4kHz
    ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * q * (1 - Gmid))
    Case 7 '8kHz
    ISO9613_Agr = (-1.5 * (1 - Gsrc)) + (-1.5 * (1 - Grec)) + (-3 * q * (1 - Gmid))
    Case 999 'catch error
    ISO9613_Agr = "-"
    End Select

'NOTES The Ground Effect formulae return positive values for attenuation (as formula is "....-A").
    If IsNumeric(ISO9613_Agr) Then
    ISO9613_Agr = ISO9613_Agr * -1
    End If

End Function


Function ISO9613_Abar(fstr As String, SourceHeight As Double, ReceiverHeight As Double, SourceReceiverDistance As Double, SourceBarrierDistance As Double, _
SrcDistanceEdge As Double, RecDistanceEdge As Double, HeightBarrierSource As Double, _
Optional DoubleDiffraction As Boolean, Optional BarrierThickness As Double, Optional HeightBarrierReceiver As Double, Optional multisource As Boolean, Optional GroundEffect As Double)
'NOTE distances are input as horizontal distances, with the hypotenuse calculated during this function
'SourceBarrierDistance = distance from source to barrier (and vice versa)

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
Dim Z As Double 'difference in path lengths of diffracted and direct sound in metres
Dim kmet As Double
Dim d_standard As Double 'includes vertical component

If IsMissing(BarrierThickness) Or DoubleDiffraction = False Then BarrierThickness = 0

f = freqStr2Num(fstr)
lambda = (343) / f 'as defined in the method
a = Abs(SrcDistanceEdge - RecDistanceEdge)

    'If the double diffraction is set as false then there are no 2 edges to the wall
    If DoubleDiffraction = False Then
      If HeightBarrierReceiver <> HeightBarrierSource Then
      HeightBarrierReceiver = HeightBarrierSource
      End If
    End If
    
DistanceRecBarrier = SourceReceiverDistance - SourceBarrierDistance - BarrierThickness

dss = ((SourceBarrierDistance ^ 2) + ((HeightBarrierSource - SourceHeight) ^ 2)) ^ (1 / 2)
dsr = ((DistanceRecBarrier ^ 2) + ((HeightBarrierReceiver - ReceiverHeight) ^ 2)) ^ (1 / 2)
d_standard = (((SourceReceiverDistance ^ 2) + ((ReceiverHeight - SourceHeight) ^ 2)) ^ (1 / 2)) 'pythagoras rules!

'Here we use the a, dss, dsr, d, e from above to calculate z , cthree
    If DoubleDiffraction = True And BarrierThickness > 0 Then
    Cthree = (1 + ((5 * lambda / BarrierThickness) * (5 * lambda / BarrierThickness))) / (1 / 3 + ((5 * lambda / BarrierThickness) * (5 * lambda / BarrierThickness)))
    Z = ((((dss + dsr + BarrierThickness) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d_standard
    'Here the case of double diffraction is considered it actually says if lambda is << e so here we have considered half of the value
        If lambda < (BarrierThickness / 2) Then
        Cthree = 3
        End If
    Else 'double diffraction is false
    Cthree = 1
    Z = ((((dss + dsr) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d_standard
    End If

'calculate the kmet: the meteorological correction

    If Z < 0 Or d_standard < 100 Then
    kmet = 1
    Else
    kmet = Exp((-1 / 2000) * (((dss * dsr * d_standard) / (2 * Z)) ^ (1 / 2)))
    End If

'Debug.Print "dss: "; dss
'Debug.Print "dsr: "; dsr
'Debug.Print "kmet: "; kmet
'Debug.Print "lambda: "; lambda
'Debug.Print "cthree: "; cthree
'Debug.Print "----------------- ";

    'condition for Ctwo
    If multisource = True Then
    Ctwo = 40
    Else
    Ctwo = 20
    End If

Dz = 10 * Application.WorksheetFunction.Log(3 + ((Ctwo / lambda) * Cthree * Abs(Z) * kmet))

    If multisource = True Then
    ISO9613_Abar = Dz * -1 'trace convention is negative!
    Else
    ISO9613_Abar = (Dz - GroundEffect) * -1
    End If
    
End Function


Function ISO9613_Cmet(fstr As String, hs, hr, dp, C0)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ISO_full(SheetType As String)

CheckRow (Selection.Row)

frmISO9613.chkAatm.Value = True
frmISO9613.chkAdiv.Value = True
frmISO9613.chkAgr.Value = True
frmISO9613.chkAbar.Value = True

frmISO9613.Show

Insert_ISO9613_CalcElements (SheetType)

End Sub

Sub A_div(SheetType As String)
CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "ISO9613: A_div"
    
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 14).Value = 10
    Cells(Selection.Row, 14).NumberFormat = "0 ""m"""
    
    Cells(Selection.Row, 15).Value = 1
    Cells(Selection.Row, 15).NumberFormat = "0 ""m"""
    
    Cells(Selection.Row, 5).Value = "=ISO9613_Adiv($N" & Selection.Row & ",$O" & Selection.Row & ")"
    ExtendFunction (SheetType)
    fmtUserInput SheetType, True
    Else 'Catch other SheetTypes
    ErrorOctOnly
    End If
End Sub

Sub A_atm(SheetType As String)

CheckRow (Selection.Row)

frmISO9613.chkAatm.Value = True
frmISO9613.chkAdiv.Value = False
frmISO9613.chkAgr.Value = False
frmISO9613.chkAbar.Value = False

frmISO9613.Show

Insert_ISO9613_CalcElements (SheetType)


'If btnOkPressed = False Then End
'
'Cells(Selection.Row, 2).Value = "ISO9613: A_atm"
'
'    If Left(SheetType, 3) = "OCT" Then
'
'    Cells(Selection.Row, 14).Value = iso9613_Temperature 'degrees
'    Cells(Selection.Row, 15).Value = iso9613_RelHumidity 'Relative Humidity
'
'    Cells(Selection.Row, 14).NumberFormat = "0""" & chr(176) & "C"""
'    Cells(Selection.Row, 15).NumberFormat = "0 ""RH"""
'
'        If InStr(1, Cells(Selection.Row - 1, 10).Formula, "ISO9613_Adiv", vbTextCompare) > 1 Then 'row above has A-div, so we can use the same input for distance!
'        Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E$6,$N" & Selection.Row - 1 & "," & iso9613_Temperature & "," & iso9613_RelHumidity & ")"
'        Else
'        Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E$6," & iso9613_d & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
'        End If
'
'    ExtendFunction (SheetType)
'
'    fmtUserInput SheetType, True
'
'    Else 'Catch other SheetTypes
'    ErrorOctOnly
'    End If
'

End Sub

Sub A_gr(SheetType As String)
CheckRow (Selection.Row)

frmISO9613.chkAdiv.Value = False
frmISO9613.chkAatm.Value = False
frmISO9613.chkAgr.Value = True
frmISO9613.chkAbar.Value = False

frmISO9613.Show

Insert_ISO9613_CalcElements (SheetType)

'Cells(Selection.Row, 2).Value = "ISO9613: A_gr"
'
'    If Left(SheetType, 3) = "OCT" Then
'    Cells(Selection.Row, 5).Value = "=ISO9613_Agr(E$6," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & _
'    iso9613_d & ",$N" & Selection.Row & ",$O" & Selection.Row & "," & iso9613_G_middle & ")"
'    Cells(Selection.Row, 14).Value = iso9613_G_source
'    Cells(Selection.Row, 14).NumberFormat = """Gs:"" 0.0"
'    Cells(Selection.Row, 15).Value = iso9613_G_receiver
'    Cells(Selection.Row, 15).NumberFormat = """Gr:"" 0.0"
'    fmtUserInput SheetType, True
'    Else 'Catch other SheetTypes
'    ErrorOctOnly
'    End If
'
'ExtendFunction (SheetType)
    
End Sub

Sub A_bar(SheetType As String)
CheckRow (Selection.Row)

frmISO9613.chkAdiv.Value = False
frmISO9613.chkAatm.Value = False
frmISO9613.chkAgr.Value = True 'required for barrier calc
frmISO9613.chkAbar.Value = True

frmISO9613.Show

Insert_ISO9613_CalcElements (SheetType)

End Sub

Sub Insert_ISO9613_CalcElements(SheetType As String)

If btnOkPressed = False Then End

    If Left(SheetType, 3) = "OCT" Then
        
        'Adiv
        If ISOFullElements(0) = True Then
        Cells(Selection.Row, 2).Value = "ISO9613: A_div"
        Cells(Selection.Row, 5).Value = "=ISO9613_Adiv($N" & Selection.Row & ",$O" & Selection.Row & ")"
        Cells(Selection.Row, 14).Value = iso9613_d
        Cells(Selection.Row, 15).Value = iso9613_d_ref
        Unit_m 14, 15
        
        ExtendFunction (SheetType)
        fmtUserInput SheetType, True
        Cells(Selection.Row + 1, Selection.Column).Select 'move down
        
        End If
        
        'Aatm
        If ISOFullElements(1) = True Then
        Cells(Selection.Row, 2).Value = "ISO9613: A_atm"
        
            If ISOFullElements(0) = True Then 'row above has A-div, so we can use the same input for distance!
            Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E$6,$N$" & Selection.Row - 1 & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
            Else
            Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E$6," & iso9613_d & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
            End If
            
        Cells(Selection.Row, 14).Value = iso9613_Temperature
        Cells(Selection.Row, 15).Value = iso9613_RelHumidity
        Cells(Selection.Row, 14).NumberFormat = "0""" & chr(176) & "C"""
        Cells(Selection.Row, 15).NumberFormat = "0 ""RH"""
        
            'data validation
            With Cells(Selection.Row, 14).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="10,15,20,30"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
            End With
        
            With Cells(Selection.Row, 15).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="20,50,70,80"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
            End With
            
            
        ExtendFunction (SheetType)
        fmtUserInput SheetType, True
        Cells(Selection.Row + 1, Selection.Column).Select 'move down
        
        End If 'end of Aatm
        
        'Agr
        If ISOFullElements(2) = True Then
        Cells(Selection.Row, 2).Value = "ISO9613: A_gr"
            
            If ISOFullElements(0) = True Then 'Two rows above has A-div, so we can use the same input for distance!
            Cells(Selection.Row, 5).Value = "=ISO9613_Agr(E$6," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & ",$N" & _
            Selection.Row - 2 & ",$N" & Selection.Row & ",$O" & Selection.Row & "," & iso9613_G_middle & ")"
            Else
            Cells(Selection.Row, 5).Value = "=ISO9613_Agr(E$6," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & _
            iso9613_d & ",$N" & Selection.Row & ",$O" & Selection.Row & "," & iso9613_G_middle & ")"
            End If
            
            Cells(Selection.Row, 14).Value = iso9613_G_source
            Cells(Selection.Row, 14).NumberFormat = """Gs:"" 0.0"
            Cells(Selection.Row, 15).Value = iso9613_G_receiver
            Cells(Selection.Row, 15).NumberFormat = """Gr:"" 0.0"
            
        ExtendFunction (SheetType)
        fmtUserInput SheetType, True
        Cells(Selection.Row + 1, Selection.Column).Select 'move down
        
        End If 'end of Agr
        
        'Abar
        If ISOFullElements(3) = True Then
        Cells(Selection.Row, 2).Value = "ISO9613: A_bar"
        Cells(Selection.Row, 14).Value = iso9613_BarrierHeight
        Cells(Selection.Row, 14).NumberFormat = "0.0 ""m"""
            If ISOFullElements(0) = True Then 'Three rows above has A-div, so we can use the same input for distance!
            Cells(Selection.Row, 5).Value = "=ISO9613_Abar(E$6," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & ",$N" & Selection.Row - 3 & "," & iso9613_SourceToBarrier & "," & _
            iso9613_SrcToBarrierEdge & "," & iso9613_RecToBarrierEdge & "," & "$N" & Selection.Row & "," & iso9613_DoubleDiffraction & "," & iso9613_BarrierThickness & "," & _
            iso9613_BarrierHeightReceiverSide & "," & iso9613_MultiSource & ",E$" & Selection.Row - 1 & ")"
            Else
            Cells(Selection.Row, 5).Value = "=ISO9613_Abar(E$6," & iso9613_SourceHeight & "," & iso9613_ReceiverHeight & "," & iso9613_d & "," & iso9613_SourceToBarrier & "," & _
            iso9613_SrcToBarrierEdge & "," & iso9613_RecToBarrierEdge & "," & "$N" & Selection.Row & "," & iso9613_DoubleDiffraction & "," & iso9613_BarrierThickness & "," & _
            iso9613_BarrierHeightReceiverSide & "," & iso9613_MultiSource & ",E$" & Selection.Row - 1 & ")"
            End If
        
        ExtendFunction (SheetType)
        fmtUserInput SheetType, True
        Cells(Selection.Row + 1, Selection.Column).Select 'move down
        
        End If 'end of Abar
        
        'Amisc
        
    Else
    ErrorOctOnly
    End If 'end of sheet type
End Sub
