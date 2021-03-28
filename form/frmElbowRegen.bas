VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbowRegen 
   Caption         =   "Elbow Regen."
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   OleObjectBlob   =   "frmElbowRegen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElbowRegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#regenerated-noise")
End Sub

Private Sub btnOK_Click()

PreviewValues

    'set public variables
    If Me.optNEBB.Value = True Then
    RegenMode = "NEBB"
    ElbowHasVanes = Me.optVanes.Value
    ElbowNumVanes = Me.txtNumVanes.Value
    FlowUnitsM3ps = Me.optMetresCubed.Value 'true if m3/s
        
    'set numeric variables
    FlowRate = ScreenInput(Me.txtFlowRate.Value)
    PressureLoss = ScreenInput(Me.txtPressureDrop.Value)
    ElementW = ScreenInput(Me.txtW.Value)
    ElementH = ScreenInput(Me.txtH.Value)
    BendCordLength = ScreenInput(Me.txtCordLength.Value)
    ElbowRadius = ScreenInput(Me.txtRadius.Value)
    'set boolean switches
    IncludeTurbulence = Me.chkTurb.Value
    MainDuctCircular = Me.optCircular.Value
    BranchDuctCircular = Me.optCircular.Value 'same as main for elbows
    
    Else 'ASHRAE
    RegenMode = "ASHRAE"
    regenNoiseElement = "Elbow"
        'set velocity from radio buttons
        If Me.optVanes.Value = True Then
        ElbowHasVanes = True
            'choose velocity
            If Me.optV1_vanes.Value = True Then
            DuctVelocity = 15
            ElseIf Me.optV2_vanes.Value = True Then
            DuctVelocity = 20
            ElseIf Me.optV3_vanes.Value = True Then
            DuctVelocity = 30
            Else
            DuctVelocity = 999
            End If
        ElseIf Me.optNoVanes = True Then 'no vanes
        ElbowHasVanes = False
            'choose velocity
            If Me.optV1_NoVanes.Value = True Then
            DuctVelocity = 10
            ElseIf Me.optV2_NoVanes.Value = True Then
            DuctVelocity = 17.5
            ElseIf Me.optV3_NoVanes.Value = True Then
            DuctVelocity = 20
            ElseIf Me.optV4_NoVanes.Value = True Then
            DuctVelocity = 25
            Else
            DuctVelocity = 999
            End If
        Else
        DuctVelocity = 999
        End If
    End If

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub chkTurb_Click()
PreviewValues
End Sub

Private Sub optASHRAE_Click()
PreviewValues
End Sub

Private Sub optCircular_Click()
PreviewValues
End Sub

Private Sub optLitres_Click()
PreviewValues
End Sub

Private Sub optMetresCubed_Click()
PreviewValues
End Sub

Private Sub optNEBB_Click()
PreviewValues
'    For i = 0 To Me.FrameASHRAE.Controls.Count - 1
'    Me.FrameASHRAE.Controls(i).Enabled = False
'    Next i
'
'    For i = 0 To Me.FrameNEBB.Controls.Count - 1
'    Me.FrameNEBB.Controls(i).Enabled = True
'    Next i
End Sub

Private Sub optNoVanes_Click()
PreviewValues
End Sub

Private Sub optRectangular_Click()
PreviewValues
End Sub

Private Sub optV1_NoVanes_Click()
PreviewValues
End Sub

Private Sub optV1_vanes_Click()
PreviewValues
End Sub

Private Sub optV2_NoVanes_Click()
PreviewValues
End Sub

Private Sub optV2_vanes_Click()
PreviewValues
End Sub

Private Sub optV3_NoVanes_Click()
PreviewValues
End Sub

Private Sub optV3_vanes_Click()
PreviewValues
End Sub

Private Sub optV4_NoVanes_Click()
PreviewValues
End Sub

Private Sub optVanes_Click()
PreviewValues
End Sub

Private Sub sbVanes_Change()
Me.txtNumVanes.Value = Me.sbVanes.Value
PreviewValues
End Sub

Private Sub txtCordLength_Change()
PreviewValues
End Sub

Private Sub txtFlowRate_Change()
PreviewValues
End Sub

Private Sub txtH_Change()
PreviewValues
End Sub

Private Sub txtPressureDrop_Change()
PreviewValues
End Sub

Private Sub txtRadius_Change()
PreviewValues
End Sub

Private Sub txtW_Change()
PreviewValues
End Sub


Sub PreviewValues()
Dim DuctVelocity As Double 'in m/s
Dim Condition As String 'vanes / no vanes
Dim FlowRateLitres As Double 'volumetric flow rate
Dim DuctAreaMsq As Double 'area in m^2

'turn buttons on and off
SelectControls

'-----------------------------------------------------------------------
'Preview values NEBB
'-----------------------------------------------------------------------
    'set public variables
    If Me.optNEBB.Value = True Then
    RegenMode = "NEBB"
    
    'calculate area
        If Me.optRectangular.Value = True Then
            If IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
            DuctAreaMsq = (Me.txtW.Value * Me.txtH.Value) / 1000000 'area in m^2
            End If
        Else 'circular
            If IsNumeric(Me.txtW.Value) Then
            DuctAreaMsq = ((Me.txtW.Value / 1000) / 2) ^ 2 * Application.WorksheetFunction.Pi 'area in m^2
            End If
        End If
    Me.txtDuctArea.Value = Round(DuctAreaMsq, 3)
    
        'check for units
        If IsNumeric(Me.txtFlowRate.Value) And _
            IsNumeric(Me.txtPressureDrop.Value) And _
            IsNumeric(Me.txtW.Value) And _
            IsNumeric(Me.txtH.Value) Then
        
            'calculate air velocity
            If Me.optMetresCubed.Value = True Then
            FlowUnitsM3ps = True
            FlowRateLitres = Me.txtFlowRate.Value * 1000
            Me.txtVelocity.Value = Round(Me.txtFlowRate.Value / DuctAreaMsq, 2)
            Else
            FlowUnitsM3ps = False
            FlowRateLitres = Me.txtFlowRate.Value
            Me.txtVelocity.Value = Round((Me.txtFlowRate.Value / 1000) / DuctAreaMsq, 1)
            End If
        
            'check for vanes and preview values
                
            'vanes
            If Me.optVanes.Value = True And IsNumeric(Me.txtCordLength.Value) Then
            Me.txt63.Value = Round(ElbowWithVanesRegen_NEBB("63", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt125.Value = Round(ElbowWithVanesRegen_NEBB("125", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt250.Value = Round(ElbowWithVanesRegen_NEBB("250", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt500.Value = Round(ElbowWithVanesRegen_NEBB("500", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt1k.Value = Round(ElbowWithVanesRegen_NEBB("1k", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt2k.Value = Round(ElbowWithVanesRegen_NEBB("2k", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt4k.Value = Round(ElbowWithVanesRegen_NEBB("4k", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            Me.txt8k.Value = Round(ElbowWithVanesRegen_NEBB("8k", Me.txtFlowRate.Value, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value, Me.optMetresCubed.Value), 1)
            ElseIf Me.optVanes = False And IsNumeric(Me.txtRadius.Value) Then 'no vanes
            Me.txt63.Value = Round(ElbowOrJunctionRegen_NEBB("63", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt125.Value = Round(ElbowOrJunctionRegen_NEBB("125", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt250.Value = Round(ElbowOrJunctionRegen_NEBB("250", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt500.Value = Round(ElbowOrJunctionRegen_NEBB("500", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt1k.Value = Round(ElbowOrJunctionRegen_NEBB("1k", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt2k.Value = Round(ElbowOrJunctionRegen_NEBB("2k", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt4k.Value = Round(ElbowOrJunctionRegen_NEBB("4k", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            Me.txt8k.Value = Round(ElbowOrJunctionRegen_NEBB("8k", Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtFlowRate.Value, Me.optCircular.Value, Me.txtW.Value, Me.txtH.Value, Me.txtRadius.Value, Me.chkTurb.Value, 1, False, Me.optMetresCubed.Value), 1)
            End If
            
        Else
        Me.txt63.Value = "-"
        Me.txt125.Value = "-"
        Me.txt250.Value = "-"
        Me.txt500.Value = "-"
        Me.txt1k.Value = "-"
        Me.txt2k.Value = "-"
        Me.txt4k.Value = "-"
        Me.txt8k.Value = "-"
        End If 'end of loop for flowrate
'-----------------------------------------------------------------------
'Preview values ASHRAE
'-----------------------------------------------------------------------
    Else 'ASHRAE
    RegenMode = "ASHRAE"
        'no vanes
        If Me.optNoVanes.Value = True Then
        Condition = "No Vanes"
            'set velocity from radio buttons
            If Me.optV1_NoVanes.Value = True Then
            DuctVelocity = 10
            ElseIf Me.optV2_NoVanes.Value = True Then
            DuctVelocity = 17.5
            ElseIf Me.optV3_NoVanes.Value = True Then
            DuctVelocity = 20
            ElseIf Me.optV4_NoVanes.Value = True Then
            DuctVelocity = 25
            Else
            DuctVelocity = 999
            End If
        Else 'vanes!
        Condition = "Vanes"
            'set velocity from radio buttons
            If Me.optV1_vanes.Value = True Then
            DuctVelocity = 15
            ElseIf Me.optV2_vanes.Value = True Then
            DuctVelocity = 20
            ElseIf Me.optV3_vanes.Value = True Then
            DuctVelocity = 30
            Else
            DuctVelocity = 999
            End If
        End If
    'Regen values
    Me.txt63.Value = RegenNoise_ASHRAE("63", "Elbow", Condition, DuctVelocity)
    Me.txt125.Value = RegenNoise_ASHRAE("125", "Elbow", Condition, DuctVelocity)
    Me.txt250.Value = RegenNoise_ASHRAE("250", "Elbow", Condition, DuctVelocity)
    Me.txt500.Value = RegenNoise_ASHRAE("500", "Elbow", Condition, DuctVelocity)
    Me.txt1k.Value = RegenNoise_ASHRAE("1k", "Elbow", Condition, DuctVelocity)
    Me.txt2k.Value = RegenNoise_ASHRAE("2k", "Elbow", Condition, DuctVelocity)
    Me.txt4k.Value = RegenNoise_ASHRAE("4k", "Elbow", Condition, DuctVelocity)
    Me.txt8k.Value = RegenNoise_ASHRAE("8k", "Elbow", Condition, DuctVelocity)
    End If


End Sub

Sub SelectControls()
'-----------------------------------------------------------------------
'Switch buttons on and off
'-----------------------------------------------------------------------
    'enable/disable buttons
    'ASHRAE Frame
    If Me.optASHRAE.Value = True Then
        For i = 0 To Me.FrameASHRAE.Controls.Count - 1
        'Debug.Print TypeName(Me.FrameASHRAE.Controls(i))
            If TypeName(Me.FrameASHRAE.Controls(i)) = "OptionButton" Then
                If Me.FrameASHRAE.Controls(i).GroupName = "A_Vel_Vanes" Then
                Me.FrameASHRAE.Controls(i).Enabled = Me.optVanes.Value
                Else 'no vanes
                Me.FrameASHRAE.Controls(i).Enabled = Me.optNoVanes.Value
                End If
            Else
            Me.FrameASHRAE.Controls(i).Enabled = True
            End If
        Next i
        
        For i = 0 To Me.FrameNEBB.Controls.Count - 1
        Me.FrameNEBB.Controls(i).Enabled = False
        Next i
        
    Else 'NEBB Frame
    
        For i = 0 To Me.FrameASHRAE.Controls.Count - 1
        Me.FrameASHRAE.Controls(i).Enabled = False
        Next i
        
        For i = 0 To Me.FrameNEBB.Controls.Count - 1
        Me.FrameNEBB.Controls(i).Enabled = True
        Next i
        
        'duct shape
        If Me.optRectangular.Value = True Then
        Me.lblDimensions.Caption = "Dimensions (W x H)"
        Me.txtH.Enabled = True
        Else
        Me.lblDimensions.Caption = "Dimensions (diameter)"
        Me.txtH.Enabled = False
        End If
        
        'vanes
        If Me.optVanes.Value = True Then
        Me.chkTurb.Enabled = False
        Me.txtCordLength.Enabled = True
        Me.txtCordLength.BackColor = &HC0FFFF
        Me.txtRadius.Enabled = False
        Me.txtRadius.BackColor = &H8000000F
        Me.txtNumVanes.Enabled = True
        Me.sbVanes.Enabled = True
        Me.optCircular.Enabled = False 'vanes can only be in rectangular ducts!
        Me.optRectangular.Value = True
        Else 'novanes
        Me.chkTurb.Enabled = True
        Me.txtCordLength.Enabled = False
        Me.txtCordLength.BackColor = &H8000000F
        Me.txtRadius.Enabled = True
        Me.txtRadius.BackColor = &HC0FFFF
        Me.txtNumVanes.Enabled = False
        Me.sbVanes.Enabled = False
        Me.optCircular.Enabled = True
        End If
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

