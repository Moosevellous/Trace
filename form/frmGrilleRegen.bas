VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrilleRegen 
   Caption         =   "Regenerated noise - Grilles"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   OleObjectBlob   =   "frmGrilleRegen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGrilleRegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnOK_Click()
'set global variables.
FlowUnitsM3ps = Me.optMetresCubed.Value 'true if m3/s
        
'set numeric variables
ElementH = CheckNumericValue(Me.txtH.Value)
ElementW = CheckNumericValue(Me.txtW.Value)
FlowRate = CheckNumericValue(Me.txtFlowRate.Value)
PressureLoss = CheckNumericValue(Me.txtPressureLoss.Value)
btnOkPressed = True
Unload Me
End Sub

Private Sub optLitres_Click()
PreviewResult
End Sub

Private Sub optMetresCubed_Click()
PreviewResult
End Sub

Private Sub txtFlowRate_Change()
PreviewResult
End Sub

Private Sub txtH_Change()
PreviewResult
End Sub

Private Sub txtPressureLoss_Change()
PreviewResult
End Sub

Private Sub txtW_Change()
PreviewResult
End Sub

Private Sub UserForm_Activate()
With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With
PreviewResult
End Sub


Sub PreviewResult()
Dim DuctAreaMsq As Double
Dim OverallLw  As Double

'calculate area and velocity
If IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
    DuctAreaMsq = (Me.txtW.Value * Me.txtH.Value) / 1000000 'area in m^2
    Me.txtDuctArea.Value = Round(DuctAreaMsq, 3)
Else
Me.txtDuctArea.Value = "-"
End If

'Calculate flow rate
If IsNumeric(Me.txtFlowRate.Value) And IsNumeric(Me.txtDuctArea.Value) Then
    If Me.optLitres.Value = True Then
        FlowRateLitres = CDbl(Me.txtFlowRate.Value)
        Me.txtVelocity.Value = Round((Me.txtFlowRate.Value / 1000) / DuctAreaMsq, 1)
    Else 'metres cubed per second
        FlowRateLitres = Me.txtFlowRate.Value * 1000
        Me.txtVelocity.Value = Round(Me.txtFlowRate.Value / DuctAreaMsq, 2)
    End If
End If

If IsNumeric(Me.txtVelocity.Value) Then
    Me.txtFpeak.Value = 160 * Me.txtVelocity.Value
    If IsNumeric(Me.txtPressureLoss.Value) Then
        OverallLw = 10 + 10 * Application.WorksheetFunction.Log(DuctAreaMsq) + 30 * Application.WorksheetFunction.Log(Me.txtPressureLoss.Value) + 5
        Me.txtOverallLw.Value = Round(OverallLw, 1)
    End If
End If

If IsNumeric(Me.txtFpeak.Value) And IsNumeric(Me.txtPressureLoss.Value) Then
    'NearestBand = NearestOctaveBand(Me.txtFpeak.Value)
    'preview the values!
    Me.txt63.Value = Round(RegenGrille_CIBSE("63", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt125.Value = Round(RegenGrille_CIBSE("125", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt250.Value = Round(RegenGrille_CIBSE("250", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt500.Value = Round(RegenGrille_CIBSE("500", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt1k.Value = Round(RegenGrille_CIBSE("1k", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt2k.Value = Round(RegenGrille_CIBSE("2k", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt4k.Value = Round(RegenGrille_CIBSE("4k", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
    Me.txt8k.Value = Round(RegenGrille_CIBSE("8k", Me.txtW.Value, Me.txtH.Value, Me.txtPressureLoss.Value, Me.txtFlowRate.Value, Me.optMetresCubed.Value), 1)
End If
        
End Sub



