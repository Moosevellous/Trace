VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstDieselEngine 
   Caption         =   "SWL Estimator - Diesel Engine"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   OleObjectBlob   =   "frmEstDieselEngine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstDieselEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LwOverall As Double
Dim K As Double
Dim ExhaustLength As Double
Dim A, B, C, D As Long

Private Sub chkRootsBlower_Click()
CalcLw
End Sub

Private Sub optNone_Click()
CalcLw
End Sub

Private Sub txtPower_Change()
CalcLw
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

'air intake correction
Me.cboxAirIntakeCorrection.Clear
Me.cboxAirIntakeCorrection.AddItem ("")
Me.cboxAirIntakeCorrection.AddItem ("Unducted air inlet to unmuffled roots blower (+3dB)")
Me.cboxAirIntakeCorrection.AddItem ("Ducted air from outside the enclosure (0dB)")
Me.cboxAirIntakeCorrection.AddItem ("Muffled roots blower (0dB)")
Me.cboxAirIntakeCorrection.AddItem ("All other inlets (with or without turbo-charger) (0dB)")

'Me.cboxDir.Clear
'Me.cboxDir.AddItem ("")
End Sub


Private Sub btnOK_Click()

btnOkPressed = True

If IsNumeric(Me.txtInExLength.Value) And IsNumeric(Me.txtPower.Value) Then
DieselPower = Me.txtPower.Value
DieselInExLength = Me.txtInExLength.Value
End If

If Me.optCasing.Value = True Then
    DieselEqn = "Lw=93+10*log(kW)" & Format(A, "0") & Format(B, "+0;0") & Format(C, "+0;0") & Format(D, "+0;0")
Else
   DieselEqn = Me.lblDieselEqn.Caption
End If

If DieselEqn = "" Then 'no formula = error
    End
End If

'DieselTurbo = Me.chkTurbo.Value
'corrections
DieselCorrection(0) = CLng(Me.txt31adj.Value)
DieselCorrection(1) = CLng(Me.txt63adj.Value)
DieselCorrection(2) = CLng(Me.txt125adj.Value)
DieselCorrection(3) = CLng(Me.txt250adj.Value)
DieselCorrection(4) = CLng(Me.txt500adj.Value)
DieselCorrection(5) = CLng(Me.txt1kadj.Value)
DieselCorrection(6) = CLng(Me.txt2kadj.Value)
DieselCorrection(7) = CLng(Me.txt4kadj.Value)
DieselCorrection(8) = CLng(Me.txt8kadj.Value)

EngineMuffler(0) = CLng(Me.txt31muff.Value)
EngineMuffler(1) = CLng(Me.txt63muff.Value)
EngineMuffler(2) = CLng(Me.txt125muff.Value)
EngineMuffler(3) = CLng(Me.txt250muff.Value)
EngineMuffler(4) = CLng(Me.txt500muff.Value)
EngineMuffler(5) = CLng(Me.txt1kmuff.Value)
EngineMuffler(6) = CLng(Me.txt2kmuff.Value)
EngineMuffler(7) = CLng(Me.txt4kmuff.Value)
EngineMuffler(8) = CLng(Me.txt8kmuff.Value)

If Me.optLowPD.Value = True Then
    If Me.optMuffSmall.Value = True Then
    MufflerDescription = "Low PD, small"
    ElseIf Me.optMuffMedium.Value = True Then
    MufflerDescription = "Low PD, medium"
    ElseIf Me.optMuffLarge.Value = True Then
    MufflerDescription = "Low PD, large"
    End If
ElseIf Me.optHighPD.Value = True Then
    If Me.optMuffSmall.Value = True Then
    MufflerDescription = "High PD, small"
    ElseIf Me.optMuffMedium.Value = True Then
    MufflerDescription = "High PD, medium"
    ElseIf Me.optMuffLarge.Value = True Then
    MufflerDescription = "High PD, large"
    End If
Else
MufflerDescription = "None"
End If

Me.Hide
Unload Me
End Sub

Private Sub cboxAirIntakeCorrection_Change()
CalcLw
End Sub

Private Sub chkTurbo_Click()

If Me.chkTurbo.Value = False And Me.optInlet.Value = True Then
    MsgBox "Inlet SWL negligible, in comparison with casing and exhaust noise", vbOKOnly, "S'all good!"
    Me.chkTurbo.Value = True
End If

CalcLw

End Sub

Private Sub opt1500RPM_Click()
Me.chkRootsBlower.Enabled = False
Me.chkRootsBlower.Value = False
CalcLw
End Sub


Private Sub opt600RPM_Click()
Me.chkRootsBlower.Enabled = False
Me.chkRootsBlower.Value = False
CalcLw
End Sub

Private Sub opt600to1500RPM_Click()
Me.chkRootsBlower.Enabled = True
CalcLw
End Sub

Private Sub optCasing_Click()
EnableFrame Me.frameCasing, True
EnableFrame Me.frameExhaust, False
Me.lblDieselEqn.Caption = "Lw=93+10*log(kW)+A+B+C+D"
CalcLw
End Sub

Public Sub EnableFrame(InFrame As Frame, ByVal Flag As Boolean)
Dim Contrl As control
On Error Resume Next

InFrame.Enabled = Flag 'enable or disable the frame that passed as parameter.
'passing over all controls
    For Each Contrl In InFrame.Controls
        If (Contrl.Container.Name = InFrame.Name) Then
        Contrl.Enabled = Flag
        End If
        
        If Flag = True Then 'some radio buttons are not enabled
'            If Me.optCircular.Value = True Then
'            EnableCircularOptions
'            Else
'            EnableRectangularOptions
'            End If
        End If
        
    Next
End Sub

Private Sub optDieselAndNaturalGas_Click()
CalcLw
End Sub

Private Sub optDieselOnly_Click()
CalcLw
End Sub

Private Sub optExhaust_Click()
EnableFrame Me.frameCasing, False
EnableFrame Me.frameExhaust, True
Me.lblInExLength.Caption = "Exhaust length"
Me.lblDieselEqn.Caption = "Lw=120+10*log(kW)-K-(L/1.2)"
Me.chkTurbo.Enabled = True
CalcLw
End Sub

Private Sub optHighPD_Click()
CalcLw
End Sub

Private Sub optInlet_Click()
EnableFrame Me.frameCasing, False
EnableFrame Me.frameExhaust, False
Me.lblInExLength.Caption = "Inlet length"
Me.lblDieselEqn.Caption = "Lw=95+5*log(kW)-(L/1.8)"
Me.chkTurbo.Value = True
'Me.chkTurbo.Enabled = False                'not needed now, message is friendlier
CalcLw
End Sub

Private Sub optInline_Click()
CalcLw
End Sub

Private Sub optLowPD_Click()
CalcLw
End Sub

Private Sub optMuffLarge_Click()
CalcLw
End Sub

Private Sub optMuffMedium_Click()
CalcLw
End Sub

Private Sub optMuffSmall_Click()
CalcLw
End Sub

Private Sub OptNaturalGasOnly_Click()
CalcLw
End Sub

Private Sub optRadial_Click()
CalcLw
End Sub

Private Sub optVtype_Click()
CalcLw
End Sub

Private Sub txtInExLength_Change()
CalcLw
End Sub


Private Sub btnCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#diesel-engines")
End Sub


Sub CalcLw()

If IsNumeric(Me.txtPower.Value) = False Or Me.txtPower.Value < 0 Then
    
    Me.txt31adj.Value = "-"
    Me.txt63adj.Value = "-"
    Me.txt125adj.Value = "-"
    Me.txt250adj.Value = "-"
    Me.txt500adj.Value = "-"
    Me.txt1kadj.Value = "-"
    Me.txt2kadj.Value = "-"
    Me.txt4kadj.Value = "-"
    Me.txt8kadj.Value = "-"
    
    Me.txt31muff.Value = 0
    Me.txt63muff.Value = 0
    Me.txt125muff.Value = 0
    Me.txt250muff.Value = 0
    Me.txt500muff.Value = 0
    Me.txt1kmuff.Value = 0
    Me.txt2kmuff.Value = 0
    Me.txt4kmuff.Value = 0
    Me.txt8kmuff.Value = 0
    
    Me.txt31.Value = "-"
    Me.txt63.Value = "-"
    Me.txt125.Value = "-"
    Me.txt250.Value = "-"
    Me.txt500.Value = "-"
    Me.txt1k.Value = "-"
    Me.txt2k.Value = "-"
    Me.txt4k.Value = "-"
    Me.txt8k.Value = "-"
Exit Sub
End If
    
'set exhaust length
If IsNumeric(Me.txtInExLength.Value) Then
    ExhaustLength = Me.txtInExLength.Value
Else
    ExhaustLength = 0
End If


If Me.optCasing.Value = True Then '<--------------------------------CASING
    
    'Speed Correction A
    If Me.opt600RPM.Value = True Then
        A = -5
    ElseIf Me.opt600to1500RPM.Value = True Then
        A = -2
    ElseIf Me.opt1500RPM.Value = True Then
        A = 0
    End If
    
    'select adjustment, based on speed - table 11.22 from B&H
    If Me.opt600RPM.Value = True Then
        SpectrumCorrection = Array(-12, -12, -6, -5, -7, -9, -12, -18, -28)
    ElseIf Me.opt600to1500RPM.Value = True Then
        If Me.chkRootsBlower.Value = True Then
        SpectrumCorrection = Array(-22, -16, -18, -14, -3, -4, -10, -15, -26)
        Else
        SpectrumCorrection = Array(-14, -9, -7, -8, -7, -7, -9, -13, -19)
        End If
    ElseIf Me.opt1500RPM.Value = True Then
        SpectrumCorrection = Array(-22, -14, -7, -7, -8, -6, -7, -13, -20)
    Else
        SpectrumCorrection = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    End If
    
    'Fuel Correction B
    If Me.optDieselOnly.Value = True Then
        B = 0
    ElseIf Me.optDieselAndNaturalGas.Value = True Then
        B = 0
    ElseIf Me.OptNaturalGasOnly.Value = True Then
        B = -3
    End If
    
    'Cylinder arrangement C
    If Me.optInline.Value = True Then
        C = 0
    ElseIf Me.optVtype.Value = True Then
        C = -1
    ElseIf Me.optRadial.Value = True Then
        C = -1
    End If
    
    'Air intake correction D
    If Me.cboxAirIntakeCorrection.Value = "Unducted air inlet to unmuffled roots blower (+3dB)" Then
        D = 3
    Else 'all other options
        D = 0
    End If
    
    LwOverall = 93 + (10 * Application.WorksheetFunction.Log(Me.txtPower.Value)) + A + B + C + D

ElseIf Me.optInlet.Value = True Then '<--------------------------------INLET

    LwOverall = 95 + 5 * Application.WorksheetFunction.Log(Me.txtPower.Value) - (ExhaustLength / 1.8) 'eqn 11.87
    MufflerCorrection = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    SpectrumCorrection = Array(-4, -11, -13, -13, -12, -9, -8, -9, -17) 'table 11.23 from B&H
    
ElseIf Me.optExhaust.Value = True Then '<--------------------------------EXHAUST
    
    If Me.chkTurbo.Value = True Then
        K = 6
    Else
        K = 0
    End If
        
    LwOverall = 120 + (10 * Application.WorksheetFunction.Log(Me.txtPower.Value)) - K - (ExhaustLength / 1.2)
    SpectrumCorrection = Array(-5, -9, -3, -7, -15, -19, -25, -35, -43)
    
    'muffler options, all from table 11.20
    If Me.optLowPD.Value = True Then
        If Me.optMuffSmall.Value = True Then
        MufflerCorrection = Array(-99, -10, -15, -13, -11, -10, -9, -8, -8)
        ElseIf Me.optMuffMedium.Value = True Then
        MufflerCorrection = Array(-99, -15, -20, -18, -16, -15, -14, -13, -13)
        ElseIf Me.optMuffLarge.Value = True Then
        MufflerCorrection = Array(-99, -20, -25, -23, -21, -20, -19, -18, -18)
        Else
        End If
    ElseIf Me.optHighPD.Value = True Then
        If Me.optMuffSmall.Value = True Then
        MufflerCorrection = Array(-99, -16, -21, -21, -19, -17, -15, -14, -14)
        ElseIf Me.optMuffMedium.Value = True Then
        MufflerCorrection = Array(-99, -20, -25, -24, -22, -20, -19, -18, -17)
        ElseIf Me.optMuffLarge.Value = True Then
        MufflerCorrection = Array(-99, -25, -29, -29, -27, -25, -24, -23, -23)
        Else
        End If
    Else
    MufflerCorrection = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    End If
    
Else 'nothing selected?
SpectrumCorrection = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
MufflerCorrection = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
End If

'put Sound power into text box
Me.txtLw.Value = Round(LwOverall, 1)

'put the values from Corection Array into the text boxes
If IsEmpty(SpectrumCorrection) = False Then
    Me.txt31adj.Value = SpectrumCorrection(0)
    Me.txt63adj.Value = SpectrumCorrection(1)
    Me.txt125adj.Value = SpectrumCorrection(2)
    Me.txt250adj.Value = SpectrumCorrection(3)
    Me.txt500adj.Value = SpectrumCorrection(4)
    Me.txt1kadj.Value = SpectrumCorrection(5)
    Me.txt2kadj.Value = SpectrumCorrection(6)
    Me.txt4kadj.Value = SpectrumCorrection(7)
    Me.txt8kadj.Value = SpectrumCorrection(8)

End If

If IsEmpty(MufflerCorrection) = False Then
    Me.txt31muff.Value = MufflerCorrection(0)
    Me.txt63muff.Value = MufflerCorrection(1)
    Me.txt125muff.Value = MufflerCorrection(2)
    Me.txt250muff.Value = MufflerCorrection(3)
    Me.txt500muff.Value = MufflerCorrection(4)
    Me.txt1kmuff.Value = MufflerCorrection(5)
    Me.txt2kmuff.Value = MufflerCorrection(6)
    Me.txt4kmuff.Value = MufflerCorrection(7)
    Me.txt8kmuff.Value = MufflerCorrection(8)
End If

'calculate spectrum (at last)
If InputsAreNumeric Then
    Me.txt31.Value = Round(LwOverall + Me.txt31adj.Value + Me.txt31muff.Value, 1)
    Me.txt63.Value = Round(LwOverall + Me.txt63adj.Value + Me.txt63muff.Value, 1)
    Me.txt125.Value = Round(LwOverall + Me.txt125adj.Value + Me.txt125muff.Value, 1)
    Me.txt250.Value = Round(LwOverall + Me.txt250adj.Value + Me.txt250muff.Value, 1)
    Me.txt500.Value = Round(LwOverall + Me.txt500adj.Value + Me.txt500muff.Value, 1)
    Me.txt1k.Value = Round(LwOverall + Me.txt1kadj.Value + Me.txt1kmuff.Value, 1)
    Me.txt2k.Value = Round(LwOverall + Me.txt2kadj.Value + Me.txt2kmuff.Value, 1)
    Me.txt4k.Value = Round(LwOverall + Me.txt4kadj.Value + Me.txt4kmuff.Value, 1)
    Me.txt8k.Value = Round(LwOverall + Me.txt8kadj.Value + Me.txt8kmuff.Value, 1)
End If

End Sub

Function InputsAreNumeric()
Dim valuesOk As Boolean
valuesOk = True
    'spectrum adjustments
    If IsNumeric(Me.txt31adj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt63adj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt125adj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt250adj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt500adj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt1kadj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt2kadj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt4kadj.Value) = False Then valuesOk = False
    If IsNumeric(Me.txt8kadj.Value) = False Then valuesOk = False
    
InputsAreNumeric = valuesOk
End Function
