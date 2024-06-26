VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstFanLw 
   Caption         =   "SWL Estimator - Fan (Simple)"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "frmEstFanLw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstFanLw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#fan-simple")
End Sub

Private Sub btnOK_Click()
CalcLw
btnOkPressed = True
FanV = CDbl(txtV.Value)
FanP = CDbl(txtP.Value)
FanType = CStr(cBoxFanType.Value)
Me.Hide
End Sub

Private Sub cBoxFanType_Change()
CalcSpectrum
End Sub

Private Sub txtP_Change()
CalcLw
CalcSpectrum
End Sub

Private Sub txtV_Change()
CalcLw
CalcSpectrum
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
FanTypes
CalcLw
End Sub

Sub FanTypes()
    If Me.cBoxFanType.ListCount = 0 Then
    cBoxFanType.AddItem ("Forward curved centrifugal")
    cBoxFanType.AddItem ("Backward curved centrifugal")
    cBoxFanType.AddItem ("Radial or paddle blade")
    cBoxFanType.AddItem ("Axial")
    cBoxFanType.AddItem ("Bifurcated")
    cBoxFanType.AddItem ("Propeller fan(approx)")
    cBoxFanType.AddItem ("Variable inlet vanes - 100%")
    cBoxFanType.AddItem ("Variable inlet vanes - 80%")
    cBoxFanType.AddItem ("Variable inlet vanes - 60%")
    cBoxFanType.AddItem ("Variable inlet vanes - 40%")
    End If
End Sub

Sub CalcLw()
    If IsNumeric(txtV.Value) And IsNumeric(txtP.Value) Then
    txtLw.Value = CStr(Round(LwFan(CDbl(txtV.Value), CDbl(txtP.Value)), 1))
    End If
End Sub

Sub CalcSpectrum()
    Select Case Me.cBoxFanType.text
    Case Is = "Forward curved centrifugal"
    Correction = Array(-5, -10, -15, -20, -25, -28, -31) 'SRL
    Case Is = "Backward curved centrifugal"
    Correction = Array(-10, -11, -10, -15, -20, -25, -30) 'SRL
    Case Is = "Radial or paddle blade"
    Correction = Array(3, -3, -10, -11, -15, -19, -23) 'SRL
    Case Is = "Axial"
    Correction = Array(-8, -8, -6, -7, -8, -12, -16) 'MDA/Woods
    Case Is = "Bifurcated"
    Correction = Array(-3, -3, -4, -5, -7, -8, -11) 'SRL
    Case Is = "Propeller fan(approx)"
    Correction = Array(-3, -4, -1, -8, -12, -13, -20) 'SRL
    'Variable Inlet Vanes
    Case Is = "Variable inlet vanes - 100%"
    Correction = Array(0, 0, 0, 0, 0, 0, 0) 'RICHDS
    Case Is = "Variable inlet vanes - 80%"
    Correction = Array(8, 5, 4, 4, 4, 4, 4) 'RICHDS
    Case Is = "Variable inlet vanes - 60%"
    Correction = Array(8, 7, 6, 5, 5, 5, 5) 'RICHDS
    Case Is = "Variable inlet vanes - 40%"
    Correction = Array(3, 2, 1, 0, 0, 0, 1) 'RICHDS
    End Select
    
    If Me.cBoxFanType.text <> "" Then
    'Corrections
    txtC63.Value = Correction(0)
    txtC125.Value = Correction(1)
    txtC250.Value = Correction(2)
    txtC500.Value = Correction(3)
    txtC1k.Value = Correction(4)
    txtC2k.Value = Correction(5)
    txtC4k.Value = Correction(6)
    'Spectrum
    txt63.Value = Round(CDbl(txtLw.Value) + Correction(0), 0)
    txt125.Value = Round(CDbl(txtLw.Value) + Correction(1), 0)
    txt250.Value = Round(CDbl(txtLw.Value) + Correction(2), 0)
    txt500.Value = Round(CDbl(txtLw.Value) + Correction(3), 0)
    txt1k.Value = Round(CDbl(txtLw.Value) + Correction(4), 0)
    txt2k.Value = Round(CDbl(txtLw.Value) + Correction(5), 0)
    txt4k.Value = Round(CDbl(txtLw.Value) + Correction(6), 0)
    End If
    
End Sub

Function LwFan(V As Double, P As Double)
    If IsNumeric(txtV.Value) And IsNumeric(txtP.Value) Then
        If V = 0 Or P = 0 Then
        LwFan = 0
        Else
        LwFan = 10 * Application.WorksheetFunction.Log10(V) + 20 * Application.WorksheetFunction.Log10(P) + 40 'v in m^3, p in Pa
        End If
    End If
End Function
