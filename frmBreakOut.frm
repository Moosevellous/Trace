VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBreakOut 
   Caption         =   "Duct Break Out (Rectangular)"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "frmBreakOut.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBreakOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fL As Single
Dim SurfaceMass As Single
Dim TLoutMin As Single

Private Sub opPVC_Click()
Me.txtDensity.Enabled = False
Me.txtDensity.Value = 1467

Me.txtDuctWallThick.Value = 3.5 'standard size
CalcValues
End Sub

Private Sub optCustom_Click()
Me.txtDensity.Enabled = True
End Sub

Private Sub optGalvanisedSteel_Click()
Me.txtDensity.Enabled = False
Me.txtDuctWallThick.Value = 0.6 'standard steel wall thickness
Me.txtDensity.Value = 7482
CalcValues
End Sub



Private Sub txtDensity_Change()
CalcValues
End Sub

Private Sub UserForm_Activate()
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
CalcValues
    
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
ductW = Me.txtW.Value
ductH = Me.txtH.Value
ductL = Me.txtL.Value
MaterialDensity = Me.txtDensity.Value
DuctWallThickness = Me.txtDuctWallThick.Value

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub txtDuctWallThick_Change()
CalcValues
End Sub

Private Sub txtH_Change()
CalcValues
End Sub

Private Sub txtL_Change()
CalcValues
End Sub

Private Sub txtW_Change()
CalcValues
End Sub

Sub CalcValues()

Dim CheckValues

'can't calculate if these fields are no good
CheckValues = True
If IsNumeric(Me.txtH.Value) = False Then CheckValues = False
If IsNumeric(Me.txtW.Value) = False Then CheckValues = False
If IsNumeric(Me.txtL.Value) = False Then CheckValues = False
If IsNumeric(Me.txtDensity.Value) = False Then CheckValues = False
If IsNumeric(Me.txtDuctWallThick.Value) = False Then CheckValues = False

    
    If CheckValues = True Then
    
    'Calculate fL
    fL = 613000# / ((CSng(Me.txtH.Value) * CSng(Me.txtW.Value)) ^ 0.5)
    fL = Round(fL, 1)
    Me.lblFL.Caption = "fL = " & CStr(fL) & " Hz"

    'calculate Surface Mass
    SurfaceMass = Me.txtDensity.Value * (CSng(Me.txtDuctWallThick.Value) / 1000)
    Me.txtSurfaceMass.Value = Round(SurfaceMass, 1)
    
    'calculate minimum TL
    TLoutMin = 10 * Application.WorksheetFunction.Log10(2 * CSng(Me.txtL.Value) * 1000 * ((1 / CSng(Me.txtW.Value)) + (1 / CSng(Me.txtH.Value)))) 'length in metres, needs to X1000
    Me.txtTLoutMin.Value = Round(TLoutMin, 1)
    
    'place in form, using actual function in module NoiseFunctions
    Me.txt31.Value = Round(GetDuctBreakout("31.5", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt63.Value = Round(GetDuctBreakout("63", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt125.Value = Round(GetDuctBreakout("125", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt250.Value = Round(GetDuctBreakout("250", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt500.Value = Round(GetDuctBreakout("500", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt1k.Value = Round(GetDuctBreakout("1k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt2k.Value = Round(GetDuctBreakout("2k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt4k.Value = Round(GetDuctBreakout("4k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt8k.Value = Round(GetDuctBreakout("8k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Else
    Me.lblFL.Caption = "fL = - "
    End If
    
End Sub

Function TLvalue(f As Single) As Single

    If Me.txtW.Value <> 0 And Me.txtH.Value <> 0 And Me.txtSurfaceMass.Value <> 0 Then
    TLvalue = 10 * Application.WorksheetFunction.Log10((f * (CSng(Me.txtSurfaceMass.Value) ^ 2)) / (CSng(Me.txtW.Value) + CSng(Me.txtH.Value))) + 17
    End If

End Function


