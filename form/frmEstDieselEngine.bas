VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstDieselEngine 
   Caption         =   "SWL Estimator - Diesel Engine"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8595
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
Dim l As Double

Private Sub btnOK_Click()
btnOkPressed = True
DieselPower = Me.txtPower.Value
DieselInExLength = Me.txtInExLength.Value
DieselEqn = Me.lblDieselEqn.Caption
DieselTurbo = Me.chkTurbo.Value
Me.Hide
Unload Me
End Sub

Private Sub chkTurbo_Click()
CalcLw
End Sub

Private Sub opt1500RPM_Click()
CalcLw
End Sub

Private Sub opt600RPM_Click()
CalcLw
End Sub

Private Sub opt600to1500RPM_Click()
CalcLw
End Sub

Private Sub optCasing_Click()
CalcLw
End Sub

Private Sub optExhaust_Click()
Me.lblInExLength.Caption = "Exhaust length"
Me.lblDieselEqn.Caption = "Lw=120+10*log(kW)-K-(L/1.2)"
Me.chkTurbo.Caption = "Turbo? (-6dB)"
Me.chkTurbo.Enabled = True
CalcLw
End Sub

Private Sub optInlet_Click()
Me.lblInExLength.Caption = "Inlet length"
Me.lblDieselEqn.Caption = "Lw=95+5*log(kW)-(L/1.8)"
Me.chkTurbo.Caption = "Turbo?"
Me.chkTurbo.Value = True
Me.chkTurbo.Enabled = False
CalcLw
End Sub

Private Sub txtInExLength_Change()
CalcLw
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub btnCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#diesel-engines")
End Sub


Sub CalcLw()

If IsNumeric(Me.txtPower.Value) And Me.txtPower.Value > 0 Then
    
    'set exhaust length
    If IsNumeric(Me.txtInExLength.Value) Then
        l = Me.txtInExLength.Value
    Else
        l = 0
    End If
    
    If Me.optCasing.Value = True Then '<--------CASING
        Me.lblDieselEqn.Caption = "93+10*log(kW)+A+B+C+D"
'        LwOverall = 93 + (10 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.cboxEnclosure.Enabled = True

    ElseIf Me.optInlet.Value = True Then '<--------INLET
        LwOverall = 95 + 5 * Application.WorksheetFunction.Log(Me.txtPower.Value) - (l / 1.8) 'eqn 11.87
        Correction = Array(-4, -11, -13, -13, -12, -9, -8, -9, -17)
'        Me.cboxEnclosure.ListIndex = 0 'no enclosure! no capes!
        Me.cboxEnclosure.Enabled = False
        
    ElseIf Me.optExhaust.Value = True Then '<--------EXHAUST
        
            If Me.chkTurbo.Value = True Then
            K = 6
            Else
            K = 0
            End If
            
        LwOverall = 120 + (10 * Application.WorksheetFunction.Log(Me.txtPower.Value)) - K - (l / 1.2)
        Correction = Array(-5, -9, -3, -7, -15, -19, -25, -35, -43)
'        Me.cboxEnclosure.ListIndex = 0 'no enclosure! no capes!
        Me.cboxEnclosure.Enabled = False
    Else 'nothing selected?
    Correction = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    End If
        
    Me.txtLw.Value = Round(LwOverall, 1)
    
    Me.txt31adj.Value = Correction(0)
    Me.txt63adj.Value = Correction(1)
    Me.txt125adj.Value = Correction(2)
    Me.txt250adj.Value = Correction(3)
    Me.txt500adj.Value = Correction(4)
    Me.txt1kadj.Value = Correction(5)
    Me.txt2kadj.Value = Correction(6)
    Me.txt4kadj.Value = Correction(7)
    Me.txt8kadj.Value = Correction(8)
    
    CalcSpectrum
    
Else 'no power, no values
    
    Me.txt31.Value = "-"
    Me.txt63.Value = "-"
    Me.txt125.Value = "-"
    Me.txt250.Value = "-"
    Me.txt500.Value = "-"
    Me.txt1k.Value = "-"
    Me.txt2k.Value = "-"
    Me.txt4k.Value = "-"
    Me.txt8k.Value = "-"
    
    Me.txt31adj.Value = "-"
    Me.txt63adj.Value = "-"
    Me.txt125adj.Value = "-"
    Me.txt250adj.Value = "-"
    Me.txt500adj.Value = "-"
    Me.txt1kadj.Value = "-"
    Me.txt2kadj.Value = "-"
    Me.txt4kadj.Value = "-"
    Me.txt8kadj.Value = "-"

End If

End Sub

Sub CalcSpectrum()
'overall values
    If InputsAreNumeric Then
    Me.txt31.Value = Round(LwOverall + Me.txt31adj.Value + Me.txt31enc.Value, 1)
    Me.txt63.Value = Round(LwOverall + Me.txt63adj.Value + Me.txt63enc.Value, 1)
    Me.txt125.Value = Round(LwOverall + Me.txt125adj.Value + Me.txt125enc.Value, 1)
    Me.txt250.Value = Round(LwOverall + Me.txt250adj.Value + Me.txt250enc.Value, 1)
    Me.txt500.Value = Round(LwOverall + Me.txt500adj.Value + Me.txt500enc.Value, 1)
    Me.txt1k.Value = Round(LwOverall + Me.txt1kadj.Value + Me.txt1kenc.Value, 1)
    Me.txt2k.Value = Round(LwOverall + Me.txt2kadj.Value + Me.txt2kenc.Value, 1)
    Me.txt4k.Value = Round(LwOverall + Me.txt4kadj.Value + Me.txt4kenc.Value, 1)
    Me.txt8k.Value = Round(LwOverall + Me.txt8kadj.Value + Me.txt8kenc.Value, 1)
    End If
End Sub

Function InputsAreNumeric()
Dim valuesOk As Boolean
valuesOk = True
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
