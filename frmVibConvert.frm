VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVibConvert 
   Caption         =   "Integrate / Differentiate Vibration"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "frmVibConvert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVibConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub


Private Sub btnOK_Click()
CalcConversionFactor
ConversionFactorStr = Me.txtConversionFactor.Value
btnOkPressed = True
Me.Hide
Unload Me
End Sub


Private Sub optAccelIn_Click()
CalcConversionFactor
End Sub

Private Sub OptAccelOut_Click()
CalcConversionFactor
End Sub

Private Sub optdB_Click()
'Me.lblUnits.Caption = "dB"
CalcConversionFactor
End Sub

Private Sub optDispIn_Click()
CalcConversionFactor
End Sub

Private Sub optDispOut_Click()
CalcConversionFactor
End Sub

Private Sub optLinear_Click()
'Me.lblUnits.Caption = "m/s"
CalcConversionFactor
End Sub

Private Sub optVelIn_Click()
CalcConversionFactor
End Sub

Private Sub optVelOut_Click()
CalcConversionFactor
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub



Sub CalcConversionFactor()
    
Dim OptionSelected As Boolean
Dim opButton As control
Dim opGroup As OptionButton

OptionSelected = True

    'check input radio boxes
    If Me.optAccelIn.Value = False And Me.optVelIn.Value = False And Me.optDispIn.Value = False Then
    OptionSelected = False 'nothing selected
    End If
    
    'check output radio boxes
    If Me.OptAccelOut.Value = False And Me.optVelOut.Value = False And Me.optDispOut.Value = False Then
    OptionSelected = False 'nothing selected
    End If
    
    
    If OptionSelected = True Then 'go ahead
    
        If (Me.optAccelIn.Value = True And Me.OptAccelOut.Value = True) Or (Me.optVelIn.Value = True And Me.optVelOut.Value = True) Or (Me.optDispIn.Value = True And Me.optDispOut.Value = True) Then
        SetFactorZero
        End If
        
        'acceleration in
        If Me.optAccelIn.Value = True Then
            If Me.OptAccelOut.Value = True Then
            SetFactorZero
            ElseIf Me.optVelOut.Value = True Then
            SetFactorInt1
            VibConversionDescription = "Acceleration->Velocity"
            ElseIf Me.optDispOut.Value = True Then
            SetFactorInt2
            VibConversionDescription = "Acceleration->Displacement"
            Else
            End If
        ElseIf Me.optVelIn.Value = True Then
            If Me.OptAccelOut.Value = True Then
            SetFactorDif1
            VibConversionDescription = "Velocity->Acceleration"
            ElseIf Me.optVelOut.Value Then
            SetFactorZero
            ElseIf Me.optDispOut.Value = True Then
            SetFactorInt1
            VibConversionDescription = "Velocity->Displacement"
            Else
            End If
        ElseIf Me.optDispIn.Value = True Then
            If Me.OptAccelOut.Value = True Then
            SetFactorDif2
            VibConversionDescription = "Displacement->Acceleration"
            ElseIf Me.optVelOut.Value Then
            SetFactorDif1
            VibConversionDescription = "Displacement->Velocity"
            ElseIf Me.optDispOut.Value = True Then
            SetFactorZero
            Else
            End If
        End If
        
        
    End If
    
End Sub


Sub SetFactorZero()
Me.txtConversionFactor.Value = 0
End Sub

Sub SetFactorDif1()
Me.txtConversionFactor.Value = "2*pi*f"
End Sub


Sub SetFactorDif2()
Me.txtConversionFactor.Value = "4*pi" & chr(178) & "*f" & chr(178)
End Sub


Sub SetFactorInt1()
Me.txtConversionFactor.Value = "1/(2*pi*f)"
End Sub


Sub SetFactorInt2()
Me.txtConversionFactor.Value = "1/(4*pi" & chr(178) & "*f" & chr(178) & ")"
End Sub
