VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuctAreas 
   Caption         =   "Duct Areas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   OleObjectBlob   =   "frmDuctAreas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuctAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptionButton1_Click()

End Sub

Private Sub txtP1_Change()
    If txtP1.Value <> "" Then
    txtP2.Value = 100 - txtP1.Value
    CalcDuctAreas
    Else
    txtP2.Value = 100
    End If
End Sub

Private Sub UserForm_Activate()
    With frmDuctAreas
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub

Private Sub lblDescription_Click()

End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
frmDuctAreas.Hide
End Sub

Private Sub btnOK_Click()
Call CalcDuctAreas
ductA1 = (CDbl(frmDuctAreas.txtL1.Value) / 1000) * (CDbl(frmDuctAreas.txtW1.Value) / 1000)
ductA2 = (CDbl(frmDuctAreas.txtL2.Value) / 1000) * (CDbl(frmDuctAreas.txtW2.Value) / 1000)
btnOkPressed = True
frmDuctAreas.Hide
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
Dim A1 As Double
Dim A2 As Double
Dim P1 As Double
Dim P2 As Double
Dim Atten As Double

    If optDimensions.Value = True Then
        'check for blank text box
        If frmDuctAreas.txtL1.Value <> "" And frmDuctAreas.txtL2.Value <> "" And frmDuctAreas.txtW1.Value <> "" And frmDuctAreas.txtW2.Value <> "" Then
        A1 = (CDbl(frmDuctAreas.txtL1.Value) / 1000) * (CDbl(frmDuctAreas.txtW1.Value) / 1000)
        A2 = (CDbl(frmDuctAreas.txtL2.Value) / 1000) * (CDbl(frmDuctAreas.txtW2.Value) / 1000)
        frmDuctAreas.txtA1.Value = CStr(A1)
        frmDuctAreas.txtA2.Value = CStr(A2)
        Atten = 10 * Application.WorksheetFunction.Log10(A2 / (A1 + A2))
        lblAtten.Caption = CStr(Round(Atten, 0))
        End If
    End If
    
    If optPercentageSplit.Value = True Then
    P1 = CDbl(txtP1.Value)
    P2 = CDbl(txtP2.Value)
        If P1 = 0 Then
        Atten = 0
        Else
        Atten = 10 * Application.WorksheetFunction.Log10(P1 / 100)
        End If
    lblAttenP = CStr(Round(Atten, 0))
    End If
        
    If optRatio.Value = True Then
    End If
End Sub
