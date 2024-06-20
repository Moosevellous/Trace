VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstElectricMotorSmall 
   Caption         =   "SWL Estimator - Electric Motor"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   OleObjectBlob   =   "frmEstElectricMotorSmall.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstElectricMotorSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#electric-motor")
End Sub

Private Sub optDRPR_Click()
Me.txt31adj.Value = -9
Me.txt63adj.Value = -9
Me.txt125adj.Value = -7
Me.txt250adj.Value = -7
Me.txt500adj.Value = -6
Me.txt1kadj.Value = -9
Me.txt2kadj.Value = -12
Me.txt4kadj.Value = -18
Me.txt8kadj.Value = -27
CalcLw
End Sub

Private Sub optTEFC_Click()
Me.txt31adj.Value = -14
Me.txt63adj.Value = -14
Me.txt125adj.Value = -11
Me.txt250adj.Value = -9
Me.txt500adj.Value = -6
Me.txt1kadj.Value = -6
Me.txt2kadj.Value = -7
Me.txt4kadj.Value = -12
Me.txt8kadj.Value = -20
CalcLw
End Sub

Private Sub txtPower_Change()

    If Me.txtPower.Value > 300 Then
    msg = MsgBox("Suitable for small motors (<300kW) only.", vbOKOnly, "Error - motor power")
    Else
    CalcLw
    End If
    
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()

CalcLw

MotorEqn = Me.lblEqn.Caption

    If Me.optDRPR.Value = True Then
    MotorType = "DRPR"
    ElseIf Me.optTEFC = True Then
    MotorType = "TEFC"
    End If

    If Me.txtPower.Value <> "" And Me.txtSpeed.Value <> "" Then
    MotorPower = Me.txtPower.Value
    MotorSpeed = Me.txtSpeed.Value
    Motor_Correction(0) = Me.txt31adj.Value
    Motor_Correction(1) = Me.txt63adj.Value
    Motor_Correction(2) = Me.txt125adj.Value
    Motor_Correction(3) = Me.txt250adj.Value
    Motor_Correction(4) = Me.txt500adj.Value
    Motor_Correction(5) = Me.txt1kadj.Value
    Motor_Correction(6) = Me.txt2kadj.Value
    Motor_Correction(7) = Me.txt4kadj.Value
    Motor_Correction(8) = Me.txt8kadj.Value
    End If
    
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub txtSpeed_Change()
CalcLw
End Sub

Sub CalcLw()

Dim Lp As Double

    If Me.txtPower.Value = "" Or Me.txtSpeed.Value = "" Then
    
    Me.txtLp = "-"
    
    Else
        
        If Me.txtPower.Value < 40 Then
        Me.lblEqn.Caption = "Lp=17+17*log(kW)+15*log(RPM)"
        Me.optUnder40kW = True
        Lp = 17 + 17 * Application.WorksheetFunction.Log10(CDbl(Me.txtPower.Value)) + 15 * Application.WorksheetFunction.Log10(Me.txtSpeed.Value)
        Me.txtLp.Value = Round(Lp, 1)
        
        
        Else
        Me.lblEqn.Caption = "Lp=28+10*log(kW)+15*log(RPM)"
        Me.optAbove40kW = True
        Lp = 28 + 10 * Application.WorksheetFunction.Log10(Me.txtPower.Value) + 15 * Application.WorksheetFunction.Log10(Me.txtSpeed.Value)
        Me.txtLp.Value = Round(Lp, 1)
        End If
    
    'Pressure
    txt31.Value = Round(Lp + Me.txt31adj.Value, 0)
    txt63.Value = Round(Lp + Me.txt63adj.Value, 0)
    txt125.Value = Round(Lp + Me.txt125adj.Value, 0)
    txt250.Value = Round(Lp + Me.txt250adj.Value, 0)
    txt500.Value = Round(Lp + Me.txt500adj.Value, 0)
    txt1k.Value = Round(Lp + Me.txt1kadj.Value, 0)
    txt2k.Value = Round(Lp + Me.txt2kadj.Value, 0)
    txt4k.Value = Round(Lp + Me.txt4kadj.Value, 0)
    txt8k.Value = Round(Lp + Me.txt8kadj.Value, 0)
    
    'Power=pressure+hemispherical spreading @1m = Lp+8
    txt31Lw.Value = txt31.Value + 8
    txt63Lw.Value = txt63.Value + 8
    txt125Lw.Value = txt125.Value + 8
    txt250Lw.Value = txt250.Value + 8
    txt500Lw.Value = txt500.Value + 8
    txt1kLw.Value = txt1k.Value + 8
    txt2kLw.Value = txt2k.Value + 8
    txt4kLw.Value = txt4k.Value + 8
    txt8kLw.Value = txt8k.Value + 8

    End If

End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

