VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBreakIn 
   Caption         =   "Duct Break-in (Rectangular)"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6825
   OleObjectBlob   =   "frmBreakIn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBreakIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F1 As Single
Dim SurfaceMass As Single

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#duct-break-in")
End Sub

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
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
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

Dim CheckValues As Boolean
Dim F1 As Single


'can't calculate if these fields are no good
CheckValues = True
If IsNumeric(Me.txtH.Value) = False Then CheckValues = False
If IsNumeric(Me.txtW.Value) = False Then CheckValues = False
If IsNumeric(Me.txtL.Value) = False Then CheckValues = False
If IsNumeric(Me.txtDensity.Value) = False Then CheckValues = False
If IsNumeric(Me.txtDuctWallThick.Value) = False Then CheckValues = False

    
    If CheckValues = True Then
    
    'Calculate fL
    F1 = (1.718 * 10 ^ 5) / Application.WorksheetFunction.Max(Me.txtW.Value, Me.txtH.Value)
    Me.lblF1.Caption = "f1 = " & CStr(Round(F1, 1)) & " Hz"

    'calculate Surface Mass
    SurfaceMass = Me.txtDensity.Value * (CSng(Me.txtDuctWallThick.Value) / 1000)
    Me.txtSurfaceMass.Value = Round(SurfaceMass, 1)
    
'    'calculate minimum TL
'    TLoutMin = 10 * Application.WorksheetFunction.Log10(2 * CSng(Me.txtL.Value) * 1000 * ((1 / CSng(Me.txtW.Value)) + (1 / CSng(Me.txtH.Value)))) 'length in metres, needs to X1000
'    Me.txtTLoutMin.Value = Round(TLoutMin, 1)
    
    'place in form, using actual function in module NoiseFunctions
    Me.txt31.Value = Round(DuctBreakIn_NEBB("31.5", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt63.Value = Round(DuctBreakIn_NEBB("63", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt125.Value = Round(DuctBreakIn_NEBB("125", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt250.Value = Round(DuctBreakIn_NEBB("250", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt500.Value = Round(DuctBreakIn_NEBB("500", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt1k.Value = Round(DuctBreakIn_NEBB("1k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt2k.Value = Round(DuctBreakIn_NEBB("2k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt4k.Value = Round(DuctBreakIn_NEBB("4k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Me.txt8k.Value = Round(DuctBreakIn_NEBB("8k", CSng(Me.txtH.Value), CSng(Me.txtW.Value), CSng(Me.txtL.Value), CSng(Me.txtDensity.Value), CSng(Me.txtDuctWallThick.Value)), 1)
    Else
    Me.lblF1.Caption = "f1 = - "
    End If
    
End Sub

'Function TLvalue(f As Single) As Single
'
'    If Me.txtW.Value <> 0 And Me.txtH.Value <> 0 And Me.txtSurfaceMass.Value <> 0 Then
'    TLvalue = 10 * Application.WorksheetFunction.Log10((f * (CSng(Me.txtSurfaceMass.Value) ^ 2)) / (CSng(Me.txtW.Value) + CSng(Me.txtH.Value))) + 17
'    End If
'
'End Function



