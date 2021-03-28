VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstPumpLw 
   Caption         =   "SWL Estimator - Pump (Simple)"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   OleObjectBlob   =   "frmEstPumpLw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstPumpLw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LpPump As Single

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#pump-simple")
End Sub

Private Sub btnOK_Click()
SelectEquation
PumpEqn = Me.lblEqn.Caption
    If IsNumeric(Me.txtPower.Value) = False Then
    btnOkPressed = False
    Else
    btnOkPressed = True
    PumpPower = Me.txtPower.Value
    End If
Me.Hide
End Sub


Private Sub opt1000_Click()
SelectEquation
End Sub

Private Sub opt1600_Click()
SelectEquation
End Sub

Private Sub opt3000_Click()
SelectEquation
End Sub

Private Sub opt450_Click()
SelectEquation
End Sub

Private Sub txtPower_Change()
SelectEquation
End Sub

Sub SelectEquation()
'equations from Table 11.10 of Beiss and Hansen
' Originally from Army, Air Force and Navy USE 1983a
    If IsNumeric(txtPower.Value) And Me.txtPower.Value <> 0 Then
        If txtPower.Value <= 75 Then
        Me.optUnder75kW.Value = True
            If Me.opt450.Value Then
            Me.lblEqn.Caption = "Lp=68+10*log(kW)"
            LpPump = 68 + 10 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (<75kW 450-900RPM)"
            ElseIf Me.opt1000.Value Then
            Me.lblEqn.Caption = "Lp=70+10*log(kW)"
            LpPump = 70 + 10 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (<75kW 1000-1500RPM)"
            ElseIf Me.opt1600.Value Then
            Me.lblEqn.Caption = "Lp=75+10*log(kW)"
            LpPump = 75 + 10 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (<75kW 1600-1800RPM)"
            ElseIf Me.opt3000.Value Then
            Me.lblEqn.Caption = "Lp=72+10*log(kW)"
            LpPump = 72 + 10 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (<75kW 3000-3600RPM)"
            Else
            Me.lblEqn.Caption = ""
            End If
        ElseIf txtPower.Value > 75 Then
        Me.optAbove75kW.Value = True
            If Me.opt450.Value Then
            Me.lblEqn.Caption = "Lp=82+3*log(kW)"
            LpPump = 82 + 3 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (>75kW 450-900RPM)"
            ElseIf Me.opt1000.Value Then
            Me.lblEqn.Caption = "Lp=84+3*log(kW)"
            LpPump = 84 + 3 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (>75kW 1000-1500RPM)"
            ElseIf Me.opt1600.Value Then
            Me.lblEqn.Caption = "Lp=89+3*log(kW)"
            LpPump = 89 + 3 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (>75kW 1600-1800RPM)"
            ElseIf Me.opt3000.Value Then
            Me.lblEqn.Caption = "Lp=86+3*log(kW)"
            LpPump = 86 + 3 * Application.WorksheetFunction.Log10(txtPower.Value)
            DescriptionString = "Pump SPL Estimate (>75kW 3000-3600RPM)"
            Else
            Me.lblEqn.Caption = ""
            End If
        Else
        
        End If
    Me.txtLp.Value = Round(LpPump, 1)
    Me.txt31.Value = Round(LpPump - 13, 1)
    Me.txt63.Value = Round(LpPump - 12, 1)
    Me.txt125.Value = Round(LpPump - 11, 1)
    Me.txt250.Value = Round(LpPump - 9, 1)
    Me.txt500.Value = Round(LpPump - 9, 1)
    Me.txt1k.Value = Round(LpPump - 6, 1)
    Me.txt2k.Value = Round(LpPump - 9, 1)
    Me.txt4k.Value = Round(LpPump - 13, 1)
    Me.txt8k.Value = Round(LpPump - 19, 1)
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
