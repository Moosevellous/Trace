Attribute VB_Name = "ISO9613"
Public ISOFullElements(5) As Boolean 'boolean array for which elements are selected: Adiv Aatm Agr Abar Amisc
Public iso9613_d As Double
Public iso9613_d_ref As Double
Public iso9613_G_source As Double
Public iso9613_G_middle As Double
Public iso9613_G_receiver As Double
Public iso9613_SourceHeight As Double
Public iso9613_ReceiverHeight As Double

Function ISO9613_Adiv(Distance As Single, D_ref As Single) 'maybe we don't need this?????
ISO9613_Adiv = -20 * Application.WorksheetFunction.Log(Distance / D_ref) + 11
End Function


Function ISO9613_Aatm(fstr As String, Distance As Double, Temperature As Integer, RelHumidity As Integer)

'These are the values from Table X of ISO9613
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


Function ISO9613_Agr(fstr As String, SourceHeight As Double, ReceiverHeight As Double, dp As Double, Gsrc As Double, Grec As Double, Gmid As Double)

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



Function ISO9613_Abar(fstr As String, d As Double, ls As Double, lr As Double, a As Double, e As Double, hs As Double, hbs As Double, dbs As Double, _
hr As Double, hbr As Double, dbr As Double, ctwo As Double, Agr As Double, multisource As Double, Optional DoubleDiffraction As Boolean)

'Final calclation variables
'Dim DoubleDiffraction As Double
'Dim multisource As Single
Dim topEdge As Single
Dim verticalEdge As Single
Dim Dz As Double
'Dim Agr As DoubleWelcome2019
Dim f As Double
Dim dss As Double
Dim dsr As Double
'hs is the height of the source
'hbs is the height of the barrier at the source
'dbs perpendicular distance between the source and the barrier at the source
dss = ((dbs ^ 2) + ((hbs - hs) ^ 2)) ^ (1 / 2)

'hr is the height of the reciever
'hbr is the height of the barrier at the reciever
'dbr perpendicular distance between the reciever and the barrier at the source

dsr = ((dbr ^ 2) + ((hbr - hr) ^ 2)) ^ (1 / 2)


Dim lambda As Double
'Dim ctwo As Double
'Dim cthree As Single
Dim z As Double
Dim kmet As Double

'z calculation variables
'Dim dss As Double
'Dim dsr As Double
'Dim d As Double
'Dim e As Double
'Dim a As Double


'User needs to enter ll and lr
'Then a check will be done to see if lr+ll is > the wavelength that can be obtained from the freqquency
'The next check will be done for multisource noices or nigh noise sources using the variable multisource
'if the variable multisource is 1, the use equation 13
'If the variable multisource is not 1 then the function checks for the variables topedge and vertical edge and assigns the equations accordingly
'If the variable topedge is set as 12 then eq 12 is used otherwise eq 13 is used
'If equation 12 is used, then Agr should be available from a certain location- Temporartily we can enter it using a input box

'DoubleDiffraction = InputBox("User please input 1 for single diffraction,2 for double diffraction and 3 for well seperated double diffraction")

'dss = InputBox("User Please Input the Distance of the source to the first diffraction edge(dss)")
'dsr = InputBox("User Please Input the Distance of the second diffraction edge to the reciever(dsr)")
'a = InputBox("Sorry for the lots of inputs user but please input the component distance parallel to the barrier edge between source and reciever in metres(a)")
'e = InputBox("Sorry for the lots of inputs user but please input the distance between the two barrier edges in metres(e) if you have selected double diffraction")
'd = InputBox("User please input the distance between the source and the reciever in metres (d)")
'cthree = InputBox("User please enter 40 if you want to consider the effect of image sources and 20 if not")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'calulating the lambda

f = freqStr2Num(fstr)

lambda = (343) / f

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Here we use the a,dss,dsr,d,e from above to calculate z , c3
'z is the difference in path lengths of diffracted and direct sound in metres
'the value of cthree changes according to the diffraction type entered
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

cthree = 1
z = ((((dss + dsr) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d
    If DoubleDiffraction = True Then
    cthree = (1 + ((5 * lambda / e) * (5 * lambda / e))) / (1 / 3 + ((5 * lambda / e) * (5 * lambda / e)))
    z = ((((dss + dsr + e) ^ 2) + (a ^ 2)) ^ (1 / 2)) - d
    'Here the case of double diffraction is considered it actually says if lambda is << e so here we have considered half of the value
        If lambda < (e / 2) Then
        cthree = 3
        End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Here we try to calculate the kmet which is the meteorological correction


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Debug.Print "z: "; z
    If z < 0 Or d < 100 Then
    kmet = 1
    Else
    kmet = Exp((-1 / 2000) * (((dss * dsr * d) / (2 * z)) ^ (1 / 2)))
    End If


'The code is workig till here
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calculation 0f Dz
Debug.Print "dss: "; dss
Debug.Print "dsr: "; dsr
Debug.Print "kmet: "; kmet
Debug.Print "lambda: "; lambda
Debug.Print "cthree: "; cthree
Debug.Print "----------------- ";
Dz = 10 * Application.WorksheetFunction.Log10(3 + ((ctwo / lambda) * cthree * Abs(z) * kmet))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (lr + ls > lambda) Then
    
      ' multisource = InputBox("Please enter 1 if the environment under consideration has multisource industrial plants or high noise sources, And of course press 0 if they are not")
    
       If multisource = 1 Then
       ISO9613_Abar = Dz
       Else
       'Agr = InputBox("User Please Enter a Reasonable Value for Ground attenuation ")
       ISO9613_Abar = Dz - Agr
       End If
    Else
        MsgBox ("User the lr+ll has to be greater than lambda to consider the attenuation due to barriers")
    
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

frmISO9613.Show
If btnOkPressed = False Then End

If Left(SheetType, 3).Value = "OCT" Then
Else
ErrorOctOnly
End If

ExtendFunction (SheetType)

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

Cells(Selection.Row, 2).Value = "ISO9613: A_atm"
    If Left(SheetType, 3) = "OCT" Then
    
    Cells(Selection.Row, 14).Value = 10 'degrees
    Cells(Selection.Row, 15).Value = 70 'Relative Humidity
    
    Cells(Selection.Row, 14).NumberFormat = "0""" & chr(176) & "C"""
    Cells(Selection.Row, 15).NumberFormat = "0 ""RH"""
    
        If InStr(1, Cells(Selection.Row - 1, 10).Formula, "ISO9613_Adiv", vbTextCompare) > 1 Then 'row above has A-div, so we can use the same input for distance!
        Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E6,$N$" & Selection.Row - 1 & ",$N" & Selection.Row & ",$O" & Selection.Row & ")"
        Else
        Cells(Selection.Row, 5).Value = "=ISO9613_Aatm(E6,10,$N" & Selection.Row & ",$O" & Selection.Row & ")" 'default to 10m for now   <-----TODO, add different options for input
        End If
        
    ExtendFunction (SheetType)

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
    
    fmtUserInput SheetType, True
    
    Else 'Catch other SheetTypes
    ErrorOctOnly
    End If
    

End Sub
