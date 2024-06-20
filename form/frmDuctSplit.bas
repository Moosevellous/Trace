VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuctSplit 
   Caption         =   "Duct Split"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12195
   OleObjectBlob   =   "frmDuctSplit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuctSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn10pc_Click()
Me.optPercentageSplit.Value = True
Me.txtP1.Value = 10
End Sub

Private Sub btn25pc_Click()
Me.optPercentageSplit.Value = True
Me.txtP1.Value = 25
End Sub

Private Sub btn50pc_Click()
Me.optPercentageSplit.Value = True
Me.txtP1.Value = 50
End Sub

Private Sub btn75pc_Click()
Me.optPercentageSplit.Value = True
Me.txtP1.Value = 75
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#duct-split")
End Sub

Private Sub optA1Circular_Click()
Me.txtW1.Enabled = False
CalcDuctAreas
End Sub

Private Sub optA1Rectangular_Click()
Me.txtW1.Enabled = True
CalcDuctAreas
End Sub

Private Sub optA2Circular_Click()
Me.txtW2.Enabled = False
CalcDuctAreas
End Sub

Private Sub optA2Rectangular_Click()
Me.txtW2.Enabled = True
CalcDuctAreas
End Sub

Private Sub optBranch_Click()

End Sub

Private Sub optDimensions_Click()

Me.txtL1.Enabled = True
Me.txtL2.Enabled = True
Me.txtW1.Enabled = True
Me.txtW2.Enabled = True
Me.lblEqn.Enabled = True
Me.lblAtten.Enabled = True
Me.lbldb.Enabled = True
Me.optA1Circular.Enabled = True
Me.optA1Rectangular.Enabled = True
Me.optA2Circular.Enabled = True
Me.optA2Rectangular.Enabled = True

Me.txtP1.Enabled = False
Me.lblEqn2.Enabled = False
Me.lblAttenP.Enabled = False
Me.lbldB2.Enabled = False

Me.txtRatio1.Enabled = False
Me.lblEqn3.Enabled = False
Me.lblAttenR.Enabled = False
Me.lbldB3.Enabled = False

End Sub

Private Sub optPercentageSplit_Click()

Me.txtL1.Enabled = False
Me.txtL2.Enabled = False
Me.txtW1.Enabled = False
Me.txtW2.Enabled = False
Me.lblEqn.Enabled = False
Me.lblAtten.Enabled = False
Me.lbldb.Enabled = False
Me.optA1Circular.Enabled = False
Me.optA1Rectangular.Enabled = False
Me.optA2Circular.Enabled = False
Me.optA2Rectangular.Enabled = False


Me.txtP1.Enabled = True
Me.lblEqn2.Enabled = True
Me.lblAttenP.Enabled = True
Me.lbldB2.Enabled = True

Me.txtRatio1.Enabled = False
Me.lblEqn3.Enabled = False
Me.lblAttenR.Enabled = False
Me.lbldB3.Enabled = False

End Sub

Private Sub optRatio_Click()
Me.txtL1.Enabled = False
Me.txtL2.Enabled = False
Me.txtW1.Enabled = False
Me.txtW2.Enabled = False
Me.lblEqn.Enabled = False
Me.lblAtten.Enabled = False
Me.lbldb.Enabled = False
Me.optA1Circular.Enabled = False
Me.optA1Rectangular.Enabled = False
Me.optA2Circular.Enabled = False
Me.optA2Rectangular.Enabled = False

Me.txtP1.Enabled = False
Me.lblEqn2.Enabled = False
Me.lblAttenP.Enabled = False
Me.lbldB2.Enabled = False

Me.txtRatio1.Enabled = True
Me.lblEqn3.Enabled = True
Me.lblAttenR.Enabled = True
Me.lbldB3.Enabled = True

End Sub

Private Sub optTee_Click()

End Sub

Private Sub txtP1_Change()
    If txtP1.Value <> "" Then
    txtP2.Value = 100 - txtP1.Value
    CalcDuctAreas
    Else
    txtP2.Value = 100
    End If
End Sub

Private Sub txtRatio1_Change()
    If Me.txtRatio1.Value <> "" Then
    CalcDuctAreas
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
Call CalcDuctAreas
    'set public variables
    If Me.optRatio.Value = True Then
    ductA1 = Me.txtRatio1.Value
    ductA2 = Me.txtRatio2.Value
    ductSplitType = "Ratio"
    ElseIf Me.optPercentageSplit.Value = True Then
    ductSplitType = "Percent"
    ductA1 = CDbl(Me.txtP1.Value / 100)
    Else 'area/dimensions whatever you want to call it
    ductSplitType = "Area"
        
        If Me.optA1Rectangular.Value = True Then
        ductA1 = (CDbl(Me.txtL1.Value) / 1000) * (CDbl(Me.txtW1.Value) / 1000)
        Else
        ductA1 = (CDbl(Me.txtL1.Value) / 2000) ^ 2 * Application.WorksheetFunction.Pi
        End If
        
        If Me.optA2Rectangular.Value = True Then
        ductA2 = (CDbl(Me.txtL2.Value) / 1000) * (CDbl(Me.txtW2.Value) / 1000)
        Else 'circular
        ductA2 = (CDbl(Me.txtL2.Value) / 2000) ^ 2 * Application.WorksheetFunction.Pi
        End If
        
    End If
btnOkPressed = True
Me.Hide
End Sub

Private Sub txtL_Change()
Call CalcDuctAreas
End Sub

Private Sub txtL1_Change()
Call CalcDuctAreas
End Sub

Private Sub txtL2_Change()
Call CalcDuctAreas
End Sub

Private Sub txtW1_Change()
Call CalcDuctAreas
End Sub

Private Sub txtW2_Change()
Call CalcDuctAreas
End Sub

Private Sub CalcDuctAreas()
Dim A1 As Single
Dim A2 As Single
Dim p1 As Double
Dim p2 As Double
Dim Ratio As Double
Dim Atten As Double
    
    If optDimensions.Value = True Then
        'check for blank text box
        If Me.txtL1.Value <> "" And Me.txtL2.Value <> "" And Me.txtW1.Value <> "" And Me.txtW2.Value <> "" Then
        
            'A1
            If Me.optA1Rectangular.Value = True Then
            A1 = (CDbl(Me.txtL1.Value) / 1000) * (CDbl(Me.txtW1.Value) / 1000)
            Else 'circular
            A1 = Application.WorksheetFunction.Pi() * ((CDbl(Me.txtL1.Value) / 2000) ^ 2)
            End If
            'A2
            If Me.optA2Rectangular = True Then
            A2 = (CDbl(Me.txtL2.Value) / 1000) * (CDbl(Me.txtW2.Value) / 1000)
            Else 'circular
            A2 = Application.WorksheetFunction.Pi() * ((CDbl(Me.txtL2.Value) / 2000) ^ 2)
            End If
            
        Me.txtA1.Value = CStr(Round(A1, 3))
        Me.txtA2.Value = CStr(Round(A2, 3))
        Atten = 10 * Application.WorksheetFunction.Log10(A2 / (A1 + A2))
        lblAtten.Caption = CStr(Round(Atten, 0))
        End If

    End If
    
    If optPercentageSplit.Value = True Then
    p1 = CDbl(txtP1.Value)
    p2 = CDbl(txtP2.Value)
        If p1 = 0 Then
        Atten = 0
        Else
        Atten = 10 * Application.WorksheetFunction.Log10(p1 / 100)
        End If
    lblAttenP = CStr(Round(Atten, 0))
    End If
        
    If optRatio.Value = True Then
    Ratio = CDbl(Me.txtRatio1.Value)
    Atten = 10 * Application.WorksheetFunction.Log10(1 / Ratio)
    Me.lblAttenR.Caption = CStr(Round(Atten, 0))
    End If
End Sub
