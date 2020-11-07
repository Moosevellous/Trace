VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbowRegen 
   Caption         =   "Elbow Regen."
   ClientHeight    =   6735
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
        If Me.optMetresCubed.Value = True Then
        FlowUnitsM3ps = True
        Else
        FlowUnitsM3ps = False
        End If
    FlowRate = Me.txtFlowRate.Value
    PressureLoss = Me.txtPressureDrop.Value
    BendW = Me.txtW.Value
    BendH = Me.txtH.Value
    BendCordLength = Me.txtCordLength.Value
    Else
    RegenMode = "ASHRAE"
        'set velocity from radio buttons
    End If

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub optASHRAE_Click()
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
Me.lblNumVanes.Enabled = False
Me.txtNumVanes.Enabled = False
Me.sbVanes.Enabled = False
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
Me.lblNumVanes.Enabled = True
Me.txtNumVanes.Enabled = True
Me.sbVanes.Enabled = True
PreviewValues
End Sub

Private Sub sbVanes_Change()
Me.txtNumVanes.Value = Me.sbVanes.Value
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

Private Sub txtW_Change()
PreviewValues
End Sub


Sub PreviewValues()
Dim DuctVelocity As Double
Dim Condition As String
Dim FlowRateLitres As Double
'-----------------------------------------------------------------------
    'enable/disable buttons
    'ASHRAE Frame
    For i = 0 To Me.FrameASHRAE.Controls.Count - 1
    'Debug.Print TypeName(Me.FrameASHRAE.Controls(i))
        If TypeName(Me.FrameASHRAE.Controls(i)) = "OptionButton" Then
            If Me.FrameASHRAE.Controls(i).GroupName = "A_Vel_Vanes" Then
            Me.FrameASHRAE.Controls(i).Enabled = Me.optVanes.Value
            Else 'no vanes
            Me.FrameASHRAE.Controls(i).Enabled = Me.optNoVanes.Value
            End If
        Else
        Me.FrameASHRAE.Controls(i).Enabled = Me.optASHRAE.Value
        End If
    Next i
    
    'NEBB Frame
    For i = 0 To Me.FrameNEBB.Controls.Count - 1
    Me.FrameNEBB.Controls(i).Enabled = Me.optNEBB.Value
    Next i
'-----------------------------------------------------------------------
    'set public variables
    If Me.optNEBB.Value = True Then
    RegenMode = "NEBB"
        'check for units
        If IsNumeric(Me.txtFlowRate.Value) Then
            If Me.optMetresCubed.Value = True Then
            FlowUnitsM3ps = True
            FlowRateLitres = Me.txtFlowRate.Value * 1000
            Else
            FlowUnitsM3ps = False
            FlowRateLitres = Me.txtFlowRate.Value
            End If
        
            'check for vanes and preview values
            If Me.optVanes.Value = True Then
            Me.txt63.Value = RegenElbowWithVanes_NEBB("63", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt125.Value = RegenElbowWithVanes_NEBB("125", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt250.Value = RegenElbowWithVanes_NEBB("250", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt500.Value = RegenElbowWithVanes_NEBB("500", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt1k.Value = RegenElbowWithVanes_NEBB("1k", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt2k.Value = RegenElbowWithVanes_NEBB("2k", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt4k.Value = RegenElbowWithVanes_NEBB("4k", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Me.txt8k.Value = RegenElbowWithVanes_NEBB("8k", FlowRateLitres, Me.txtPressureDrop.Value, Me.txtW.Value, Me.txtH.Value, Me.txtCordLength.Value, Me.txtNumVanes.Value)
            Else 'no vanes
            End If
        End If
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

