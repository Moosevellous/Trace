VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstGasTurbine 
   Caption         =   "SWL Estimator - Gas Turbines"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   OleObjectBlob   =   "frmEstGasTurbine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstGasTurbine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LwTurbine As Single

Private Sub btnCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#steam-turbines")
End Sub

Private Sub btnOK_Click()
'store public variables
TurbineEqn = Me.lblGasEqn.Caption
    If Me.txtPower.Value <> "" Then
    TurbinePower = Me.txtPower.Value
    
    'spectrum corrections
    TurbineCorrection(0) = CLng(Me.txt31adj.Value)
    TurbineCorrection(1) = CLng(Me.txt63adj.Value)
    TurbineCorrection(2) = CLng(Me.txt125adj.Value)
    TurbineCorrection(3) = CLng(Me.txt250adj.Value)
    TurbineCorrection(4) = CLng(Me.txt500adj.Value)
    TurbineCorrection(5) = CLng(Me.txt1kadj.Value)
    TurbineCorrection(6) = CLng(Me.txt2kadj.Value)
    TurbineCorrection(7) = CLng(Me.txt4kadj.Value)
    TurbineCorrection(8) = CLng(Me.txt8kadj.Value)
    
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
    
        If Me.optCasing.Value = True Then
        GasTurbineType = "Casing"
        ElseIf Me.optExhaust.Value = True Then
        GasTurbineType = "Exhaust"
        ElseIf Me.optInlet.Value = True Then
        GasTurbineType = "Inlet"
        End If
        
    btnOkPressed = True
    Else
    btnOkPressed = False
    End If
    
EnclosureDescription = Me.cboxEnclosure.Value
    
Me.Hide
Unload Me
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

Private Sub optCasing_Click()
SelectPath
End Sub

Private Sub optExhaust_Click()
SelectPath
End Sub

Private Sub optInlet_Click()
SelectPath
End Sub

Private Sub txtPower_Change()
SelectPath
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
EnclosureTypes
SelectPath
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

Sub SelectPath()

    If Me.txtPower.Value <> "" And Me.txtPower.Value <> 0 Then
    
        If Me.optCasing.Value = True Then '<--------CASING
        Me.lblGasEqn.Caption = "Lw=120+5*log(MW)"
        LwTurbine = 120 + (5 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.txtLw.Value = Round(LwTurbine, 1)
        Me.txt31adj.Value = -10
        Me.txt63adj.Value = -7
        Me.txt125adj.Value = -5
        Me.txt250adj.Value = -4
        Me.txt500adj.Value = -4
        Me.txt1kadj.Value = -4
        Me.txt2kadj.Value = -4
        Me.txt4kadj.Value = -4
        Me.txt8kadj.Value = -4
        Me.cboxEnclosure.Enabled = True
        CalcSpectrum
        ElseIf Me.optInlet.Value = True Then '<--------INLET
        Me.lblGasEqn.Caption = "Lw=127+15*log(MW)"
        LwTurbine = 127 + (15 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.txtLw.Value = Round(LwTurbine, 1)
        Me.txt31adj.Value = -19
        Me.txt63adj.Value = -18
        Me.txt125adj.Value = -17
        Me.txt250adj.Value = -17
        Me.txt500adj.Value = -14
        Me.txt1kadj.Value = -8
        Me.txt2kadj.Value = -3
        Me.txt4kadj.Value = -3
        Me.txt8kadj.Value = -6
        Me.cboxEnclosure.ListIndex = 0 'no enclosure! no capes!
        Me.cboxEnclosure.Enabled = False
        CalcSpectrum
        ElseIf Me.optExhaust.Value = True Then '<--------EXHAUST
        Me.lblGasEqn.Caption = "Lw=133+10*log(MW)"
        LwTurbine = 133 + (10 * Application.WorksheetFunction.Log(Me.txtPower.Value))
        Me.txtLw.Value = Round(LwTurbine, 1)
        Me.txt31adj.Value = -12
        Me.txt63adj.Value = -8
        Me.txt125adj.Value = -6
        Me.txt250adj.Value = -6
        Me.txt500adj.Value = -7
        Me.txt1kadj.Value = -9
        Me.txt2kadj.Value = -11
        Me.txt4kadj.Value = -15
        Me.txt8kadj.Value = -21
        Me.cboxEnclosure.ListIndex = 0 'no enclosure! no capes!
        Me.cboxEnclosure.Enabled = False
        CalcSpectrum
        End If
        

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
    Me.txt31.Value = Round(LwTurbine + Me.txt31adj.Value + Me.txt31enc.Value, 1)
    Me.txt63.Value = Round(LwTurbine + Me.txt63adj.Value + Me.txt63enc.Value, 1)
    Me.txt125.Value = Round(LwTurbine + Me.txt125adj.Value + Me.txt125enc.Value, 1)
    Me.txt250.Value = Round(LwTurbine + Me.txt250adj.Value + Me.txt250enc.Value, 1)
    Me.txt500.Value = Round(LwTurbine + Me.txt500adj.Value + Me.txt500enc.Value, 1)
    Me.txt1k.Value = Round(LwTurbine + Me.txt1kadj.Value + Me.txt1kenc.Value, 1)
    Me.txt2k.Value = Round(LwTurbine + Me.txt2kadj.Value + Me.txt2kenc.Value, 1)
    Me.txt4k.Value = Round(LwTurbine + Me.txt4kadj.Value + Me.txt4kenc.Value, 1)
    Me.txt8k.Value = Round(LwTurbine + Me.txt8kadj.Value + Me.txt8kenc.Value, 1)
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
