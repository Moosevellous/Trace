VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstCoolingTower 
   Caption         =   "SWL Estimator - Cooling Tower"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045
   OleObjectBlob   =   "frmEstCoolingTower.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstCoolingTower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
    If IsNumeric(Me.txtPower.Value) = False Then
    btnOkPressed = False
    Else
    btnOkPressed = True
    CTPower = Me.txtPower.Value
    End If
Me.Hide
End Sub

Private Sub optCentrifugalType_Click()
    If optCentrifugalType.Value = True Then
    SelectCentrifugalType
    Else
    SelectPropellerType
    End If
End Sub

Private Sub optPropellerType_Click()
    If optPropellerType.Value = True Then
    SelectPropellerType
    Else
    SelectCentrifugalType
    End If
End Sub

Private Sub txtPower_Change()
SetEqn
End Sub

Private Sub UserForm_Activate()
SetEqn
    With Me
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub

Sub SelectPropellerType()
Me.optAbove60kW.Value = False
Me.optUnder60kW.Value = False
Me.txt31adj.Value = -8
Me.txt63adj.Value = -5
Me.txt125adj.Value = -5
Me.txt250adj.Value = -8
Me.txt500adj.Value = -11
Me.txt1kadj.Value = -15
Me.txt2kadj.Value = -18
Me.txt4kadj.Value = -21
Me.txt8kadj.Value = -29
End Sub

Sub SelectCentrifugalType()
Me.optAbove75kW.Value = False
Me.optUnder75kW.Value = False
Me.txt31adj.Value = -6
Me.txt63adj.Value = -6
Me.txt125adj.Value = -8
Me.txt250adj.Value = -10
Me.txt500adj.Value = -11
Me.txt1kadj.Value = -13
Me.txt2kadj.Value = -12
Me.txt4kadj.Value = -18
Me.txt8kadj.Value = -25
End Sub

Sub SetEqn()
    If optPropellerType.Value = True Then
        If txtPower.Value > 75 Then
        lblEqn.Caption = "Lw=96+10log(kW)"
        optAbove75kW.Value = True
        optUnder75kW.Value = False
        Else
        lblEqn.Caption = "Lw=100+8log(kW)"
        optAbove75kW.Value = False
        optUnder75kW.Value = True
        End If
    ElseIf optCentrifugalType.Value = True Then
        If txtPower.Value > 60 Then
        lblEqn.Caption = "Lw=85+11log(kW)"
        optAbove60kW.Value = True
        Else
        lblEqn.Caption = "Lw=93+7log(kW)"
        optUnder60kW.Value = True
        End If
    End If
End Sub

Sub CalcLw()
txtLw.Value = 999
End Sub
