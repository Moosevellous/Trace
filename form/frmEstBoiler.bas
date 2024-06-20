VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstBoiler 
   Caption         =   "SWL Estimator - Boiler"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   OleObjectBlob   =   "frmEstBoiler.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstBoiler"
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
GotoWikiPage ("Estimator-Functions#Boiler")
End Sub

Private Sub btnOK_Click()
btnOkPressed = True
'store public variables
If IsNumeric(Me.txtPower.Value) Then BoilerPower = Me.txtPower.Value
BoilerEqn = Me.lblEqn.Caption
BoilerCorrection(0) = CLng(Me.txt31adj.Value)
BoilerCorrection(1) = CLng(Me.txt63adj.Value)
BoilerCorrection(2) = CLng(Me.txt125adj.Value)
BoilerCorrection(3) = CLng(Me.txt250adj.Value)
BoilerCorrection(4) = CLng(Me.txt500adj.Value)
BoilerCorrection(5) = CLng(Me.txt1kadj.Value)
BoilerCorrection(6) = CLng(Me.txt2kadj.Value)
BoilerCorrection(7) = CLng(Me.txt4kadj.Value)
BoilerCorrection(8) = CLng(Me.txt8kadj.Value)
    If Me.optGeneralPurpose.Value = True Then
    BoilerType = "General Purpose"
    ElseIf Me.optLargePowerPlant.Value = True Then
    BoilerType = "Large Power Plant"
    Else
    msg = MsgBox("Error - nothing selected??????", vbOKOnly, "How????")
    End If

Me.Hide
End Sub

Private Sub optGeneralPurpose_Click()
UpdateCalc
End Sub

Private Sub optLargePowerPlant_Click()
UpdateCalc
End Sub

Private Sub txtPower_Change()
UpdateCalc
End Sub


Sub UpdateCalc()

Dim LwBoiler As Single

    'values & captions
    If IsNumeric(Me.txtPower.Value) Then
    
        If Me.optGeneralPurpose.Value = True Then
        Me.lblkWMW.Caption = "kW"
        Me.lblEqn.Caption = "Lw=95+4*log(kW)"
        LwBoiler = 95 + (4 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.txt31adj.Value = -6
        Me.txt63adj.Value = -6
        Me.txt125adj.Value = -7
        Me.txt250adj.Value = -9
        Me.txt500adj.Value = -12
        Me.txt1kadj.Value = -15
        Me.txt2kadj.Value = -18
        Me.txt4kadj.Value = -21
        Me.txt8kadj.Value = -24
        ElseIf Me.optLargePowerPlant.Value = True Then
        Me.lblkWMW.Caption = "MW"
        Me.lblEqn.Caption = "Lw=84+15*log(MW)"
        LwBoiler = 84 + (15 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.txt31adj.Value = -4
        Me.txt63adj.Value = -5
        Me.txt125adj.Value = -10
        Me.txt250adj.Value = -16
        Me.txt500adj.Value = -17
        Me.txt1kadj.Value = -19
        Me.txt2kadj.Value = -21
        Me.txt4kadj.Value = -21
        Me.txt8kadj.Value = -21
        Else
        msg = MsgBox("Error - nothing selected??????", vbOKOnly, "How????")
        End If
    
    'calculate spectrum
    Me.txtLw.Value = Round(LwBoiler, 1)
    Me.txt31.Value = Round(LwBoiler - Me.txt31adj.Value, 1)
    Me.txt63.Value = Round(LwBoiler - Me.txt63adj.Value, 1)
    Me.txt125.Value = Round(LwBoiler - Me.txt125adj.Value, 1)
    Me.txt250.Value = Round(LwBoiler - Me.txt250adj.Value, 1)
    Me.txt500.Value = Round(LwBoiler - Me.txt500adj.Value, 1)
    Me.txt1k.Value = Round(LwBoiler - Me.txt1kadj.Value, 1)
    Me.txt2k.Value = Round(LwBoiler - Me.txt2kadj.Value, 1)
    Me.txt4k.Value = Round(LwBoiler - Me.txt4kadj.Value, 1)
    Me.txt8k.Value = Round(LwBoiler - Me.txt8kadj.Value, 1)
    End If

End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

