VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstSteamTurbine 
   Caption         =   "SWL Estimator - Steam Turbines"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460.001
   OleObjectBlob   =   "frmEstSteamTurbine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstSteamTurbine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#steam-turbines")
End Sub

Private Sub btnOK_Click()

'store public variables
TurbineEqn = Me.lblSteamEqn.Caption
TurbinePower = Me.txtPower.Value

'spectrum corrections
TurbineCorrection(0) = CLng(Me.txt31_st_cor.Value)
TurbineCorrection(1) = CLng(Me.txt63_st_cor.Value)
TurbineCorrection(2) = CLng(Me.txt125_st_cor.Value)
TurbineCorrection(3) = CLng(Me.txt250_st_cor.Value)
TurbineCorrection(4) = CLng(Me.txt500_st_cor.Value)
TurbineCorrection(5) = CLng(Me.txt1k_st_cor.Value)
TurbineCorrection(6) = CLng(Me.txt2k_st_cor.Value)
TurbineCorrection(7) = CLng(Me.txt4k_st_cor.Value)
TurbineCorrection(8) = CLng(Me.txt8k_st_cor.Value)
 
'enclosures
TurbineEnclosure(0) = CLng(Me.txt31enc.Value)
TurbineEnclosure(1) = CLng(Me.txt63enc.Value)
TurbineEnclosure(2) = CLng(Me.txt125enc.Value)
TurbineEnclosure(3) = CLng(Me.txt250enc.Value)
TurbineEnclosure(4) = CLng(Me.txt500enc.Value)
TurbineEnclosure(5) = CLng(Me.txt1kenc.Value)
TurbineEnclosure(6) = CLng(Me.txt2kenc.Value)
TurbineEnclosure(7) = CLng(Me.txt4kenc.Value)
TurbineEnclosure(8) = CLng(Me.txt8kenc.Value)

EnclosureDescription = Me.cboxEnclosure.Value

btnOkPressed = True
Me.Hide
Unload Me
End Sub


Sub EnclosureTypes()

    If Me.cboxEnclosure.ListCount = 0 Then
        cboxEnclosure.AddItem ("0 - No enclosure")
        cboxEnclosure.AddItem ("1 - Glass fibre / mineral wool with lightweight foil")
        cboxEnclosure.AddItem ("2 - Glass fibre / mineral wool with 20 or 24 gauge aluminium")
        cboxEnclosure.AddItem ("3 - Enclosing metal cabinet with open ventilation holes - no internal lining")
        cboxEnclosure.AddItem ("4 - Enclosing metal cabinet with open ventilation holes - internal acoustic lining")
        cboxEnclosure.AddItem ("5 - Enclosing metal cabinet with all ventilation holes muffled and internal acoustic lining")
    End If

End Sub


Private Sub cboxEnclosure_Change()

Dim SplitEnclosureString() As String

    If Me.cboxEnclosure.Value = "" Then
    ReDim SplitEnclosureString(1)
    'SplitEnclosureString(0) = "0"
    Else
    SplitEnclosureString = Split(Me.cboxEnclosure.Text)
    End If

    'assign corrections for casing noise reduction, from 31.5Hz
    Select Case SplitEnclosureString(0) 'first elemenet
    Case Is = ""
    EnclosureReduction = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    Case Is = "0"
    EnclosureReduction = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
    Case Is = "1"
    EnclosureReduction = Array(-2, -2, -2, -3, -3, -3, -4, -5, -6)
    Case Is = "2"
    EnclosureReduction = Array(-4, -5, -5, -6, -6, -7, -8, -9, -10)
    Case Is = "3"
    EnclosureReduction = Array(-1, -1, -1, -2, -2, -2, -2, -3, -3)
    Case Is = "4"
    EnclosureReduction = Array(-3, -4, -4, -5, -6, -7, -8, -8, -8)
    Case Is = "5"
    EnclosureReduction = Array(-6, -7, -8, -9, -10, -11, -12, -13, -14)
    End Select

'update text boxes to show casing enclosure reductions
txt31enc.Value = EnclosureReduction(0)
txt63enc.Value = EnclosureReduction(1)
txt125enc.Value = EnclosureReduction(2)
txt250enc.Value = EnclosureReduction(3)
txt500enc.Value = EnclosureReduction(4)
txt1kenc.Value = EnclosureReduction(5)
txt2kenc.Value = EnclosureReduction(6)
txt4kenc.Value = EnclosureReduction(7)
txt8kenc.Value = EnclosureReduction(8)

CalcSpectrum

End Sub

Private Sub txtPower_Change()
CalcSpectrum
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
    
EnclosureTypes

End Sub

Sub CalcSpectrum()
Dim Lw As Single

    If IsNumeric(Me.txtPower.Value) Then
    
    Lw = 93 + 4 * Application.WorksheetFunction.Log(Me.txtPower.Value)
    Me.txtLw.Text = Round(Lw, 1)
    
    Me.txt31.Value = Round(Lw + CSng(Me.txt31_st_cor) + CSng(Me.txt31enc), 1)
    Me.txt63.Value = Round(Lw + CSng(Me.txt63_st_cor) + CSng(Me.txt63enc), 1)
    Me.txt125.Value = Round(Lw + CSng(Me.txt125_st_cor) + CSng(Me.txt125enc), 1)
    Me.txt250.Value = Round(Lw + CSng(Me.txt250_st_cor) + CSng(Me.txt250enc), 1)
    Me.txt500.Value = Round(Lw + CSng(Me.txt500_st_cor) + CSng(Me.txt500enc), 1)
    Me.txt1k.Value = Round(Lw + CSng(Me.txt1k_st_cor) + CSng(Me.txt1kenc), 1)
    Me.txt2k.Value = Round(Lw + CSng(Me.txt2k_st_cor) + CSng(Me.txt2kenc), 1)
    Me.txt4k.Value = Round(Lw + CSng(Me.txt4k_st_cor) + CSng(Me.txt4kenc), 1)
    Me.txt8k.Value = Round(Lw + CSng(Me.txt8k_st_cor) + CSng(Me.txt8kenc), 1)
    
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
    
    End If
End Sub


Sub Enclosures()


End Sub


