VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDamperRegen 
   Caption         =   "Regenerated noise - Dampers"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "frmDamperRegen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDamperRegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#regenerated-noise")
End Sub

Private Sub btnOK_Click()

PreviewDamperLw

    'set public variables
    If Me.optNEBB.Value = True Then
    RegenMode = "NEBB"
        If Me.optMetresCubed.Value = True Then
        FlowUnitsM3ps = True
        Else
        FlowUnitsM3ps = False
        End If
    FlowRate = Me.txtFlowRate.Value
    PressureLoss = Me.txtPressureLoss.Value
    ElementW = Me.txtW.Value
    ElementH = Me.txtH.Value
    DamperMultiBlade = Me.optMultiBlade.Value
    Else
    RegenMode = "ASHRAE"
        'set velocity from radio buttons
        If Me.optV1.Value = True Then
        DuctVelocity = 3.5
        ElseIf Me.optV2.Value = True Then
        DuctVelocity = 5.5
        ElseIf Me.optV3.Value = True Then
        DuctVelocity = 8.75
        ElseIf Me.optV4.Value = True Then
        DuctVelocity = 11
        ElseIf Me.optV5.Value = True Then
        DuctVelocity = 14.5
        Else
        DuctVelocity = 999
        End If
    End If
    
btnOkPressed = True
Me.Hide
End Sub

Private Sub optASHRAE_Click()

    For i = 0 To Me.FrameASHRAE.Controls.Count - 1
    Me.FrameASHRAE.Controls(i).Enabled = True
    Next i

    For i = 0 To Me.FrameNEBB.Controls.Count - 1
    Me.FrameNEBB.Controls(i).Enabled = False
    Next i

PreviewDamperLw
End Sub

Private Sub optLitres_Click()
PreviewDamperLw
End Sub

Private Sub optMetresCubed_Click()
PreviewDamperLw
End Sub

Private Sub optMultiBlade_Click()
PreviewDamperLw
End Sub

Private Sub optNEBB_Click()

    For i = 0 To Me.FrameASHRAE.Controls.Count - 1
    Me.FrameASHRAE.Controls(i).Enabled = False
    Next i

    For i = 0 To Me.FrameNEBB.Controls.Count - 1
    Me.FrameNEBB.Controls(i).Enabled = True
    Next i
    
PreviewDamperLw
End Sub

Private Sub optSingleBlade_Click()
PreviewDamperLw
End Sub

Private Sub optV1_Click()
PreviewDamperLw
End Sub

Private Sub optV2_Click()
PreviewDamperLw
End Sub

Private Sub optV3_Click()
PreviewDamperLw
End Sub

Private Sub optV4_Click()
PreviewDamperLw
End Sub

Private Sub optV5_Click()
PreviewDamperLw
End Sub

Private Sub txtFlowRate_Change()
PreviewDamperLw
End Sub

Private Sub txtH_Change()
PreviewDamperLw
End Sub

Private Sub txtPressureLoss_Change()
PreviewDamperLw
End Sub

Private Sub txtW_Change()
PreviewDamperLw
End Sub


Sub PreviewDamperLw()
Dim FlowRateLitres As Double
Dim DuctAreaMsq As Double
Dim VelocityMps As Double

    'NEBB MODE
    If Me.optNEBB.Value = True Then
        'calculate area and velocity
        If IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
        DuctAreaMsq = (Me.txtW.Value * Me.txtH.Value) / 1000000 'area in m^2
        Me.txtDuctArea.Value = Round(DuctAreaMsq, 3)
        End If
        
        'Calculate flow rate
        If IsNumeric(Me.txtFlowRate.Value) Then
            If Me.optLitres.Value = True Then
            FlowRateLitres = CDbl(Me.txtFlowRate.Value)
            Me.txtVelocity.Value = Round((Me.txtFlowRate.Value / 1000) / DuctAreaMsq, 1)
            Else 'metres cubed per second
            FlowRateLitres = Me.txtFlowRate.Value * 1000
            Me.txtVelocity.Value = Round(Me.txtFlowRate.Value / DuctAreaMsq, 2)
            End If
        End If
    
        'calculate values
        If IsNumeric(Me.txtFlowRate.Value) And IsNumeric(Me.txtPressureLoss.Value) And IsNumeric(Me.txtH.Value) And IsNumeric(Me.txtW.Value) And Me.txtDuctArea.Value <> "0" Then
        Me.txt63.Value = CheckNumericValue(DamperRegen_NEBB("63", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt125.Value = CheckNumericValue(DamperRegen_NEBB("125", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt250.Value = CheckNumericValue(DamperRegen_NEBB("250", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt500.Value = CheckNumericValue(DamperRegen_NEBB("500", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt1k.Value = CheckNumericValue(DamperRegen_NEBB("1k", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt2k.Value = CheckNumericValue(DamperRegen_NEBB("2k", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt4k.Value = CheckNumericValue(DamperRegen_NEBB("4k", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Me.txt8k.Value = CheckNumericValue(DamperRegen_NEBB("8k", FlowRateLitres, Me.txtPressureLoss.Value, Me.txtH.Value, Me.txtW.Value, Me.optMultiBlade.Value), 1)
        Else 'clear! *beeeeeeep*
        NoSpectrum
        End If
    Else 'ASHRAE MODE
        'set velocity from radio buttons
        If Me.optV1.Value = True Then
        VelocityMps = 3.5
        ElseIf Me.optV2.Value = True Then
        VelocityMps = 5.5
        ElseIf Me.optV3.Value = True Then
        VelocityMps = 8.75
        ElseIf Me.optV4.Value = True Then
        VelocityMps = 11
        ElseIf Me.optV5.Value = True Then
        VelocityMps = 14.5
        Else
        VelocityMps = 999
        End If
        
        'set values
        If VelocityMps = 999 Then
        NoSpectrum
        Else
        Me.txt63.Value = Round(CDbl(RegenNoise_ASHRAE("63", "Damper", "", VelocityMps)), 1)
        Me.txt125.Value = Round(RegenNoise_ASHRAE("125", "Damper", "", VelocityMps), 1)
        Me.txt250.Value = Round(RegenNoise_ASHRAE("250", "Damper", "", VelocityMps), 1)
        Me.txt500.Value = Round(RegenNoise_ASHRAE("500", "Damper", "", VelocityMps), 1)
        Me.txt1k.Value = Round(RegenNoise_ASHRAE("1k", "Damper", "", VelocityMps), 1)
        Me.txt2k.Value = Round(RegenNoise_ASHRAE("2k", "Damper", "", VelocityMps), 1)
        Me.txt4k.Value = Round(RegenNoise_ASHRAE("4k", "Damper", "", VelocityMps), 1)
        Me.txt8k.Value = "-"
        End If
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewDamperLw
End Sub

Sub NoSpectrum()
Me.txt63.Value = "-"
Me.txt125.Value = "-"
Me.txt250.Value = "-"
Me.txt500.Value = "-"
Me.txt1k.Value = "-"
Me.txt2k.Value = "-"
Me.txt4k.Value = "-"
Me.txt8k.Value = "-"
End Sub
