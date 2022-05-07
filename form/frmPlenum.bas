VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlenum 
   Caption         =   "Plenum Insertion Loss (ASHRAE)"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12780
   OleObjectBlob   =   "frmPlenum.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlenum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SoS As Single
Dim b As Single
Dim n As Single
Dim Vol As Single
Dim InletArea As Single
Dim OutletArea As Single
Dim R As Single
Dim theta As Single
'Dim PlenumL As Long 'not requried as already public variables
'Dim PlenumW As Long
'Dim PlenumH As Long
Dim alpha2(6) As Variant 'alpha2, plenum lining
Dim alpha1(6) As Variant 'alpha1, bare plenum material
Dim alphaTotal(6) As Variant
Dim f_co As Single
Dim SA As Double

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#plenum")
End Sub

Private Sub btnOK_Click()

Dim SplitStr() As String

    If CheckFormFields = True Then
    
    'set public variables
    PlenumL = Me.txtL.Value
    PlenumW = Me.txtW.Value
    PlenumH = Me.txtH.Value
    DuctInL = Me.txtInL.Value
    DuctInW = Me.txtInW.Value
    DuctOutL = Me.txtOutL.Value
    DuctOutW = Me.txtOutW.Value
    R_H = Me.txtHorizontalOffset.Value
    r_v = Me.txtVerticalOffset.Value
    UnlinedType = Me.cBoxUnlinedMaterial.Value
    SplitStr = Split(Me.cBoxLining.Value, ",", Len(Me.cBoxLining.Value), vbTextCompare)
    PlenumLiningType = SplitStr(0) 'first element, before the comma
    PlenumWallEffectStr = Me.cBoxWallEffect.Value
    PlenumPercentUnlined = Me.sbUnlinedPercent.Value
    
        'Q factor
        If Me.optInletCorner.Value = True Then
        PlenumQ = 4
        Else
        PlenumQ = 2 'default value
        End If
        
        
        'Elbow Effect
        If Me.optEndSide.Value = True Then 'CHECK THIS
        ApplyPlenumElbowEffect = True
        ElseIf Me.optEndEnd.Value = True Then
        ApplyPlenumElbowEffect = False
        Else
        msg = MsgBox("Error: Plenum configuration invalid", vbOKOnly, "Inside Outside Inside On")
        End If
    
    btnOkPressed = True
    Unload Me
    Else
    btnOkPressed = False
    End If
End Sub


Private Sub cBoxConfiguration_Change()
PreviewInsertionLoss
End Sub

Private Sub cBoxLining_Change()

Dim alphaSelected As Variant

    Select Case Me.cBoxLining.Value
    Case Is = "Concrete"
    alphaSelected = Array(0.01, 0.01, 0.01, 0.02, 0.02, 0.02, 0.03)
    Case Is = "Bare Sheet Metal"
    alphaSelected = Array(0.04, 0.04, 0.04, 0.05, 0.05, 0.05, 0.07)
    Case Is = "25mm fibreglass, 48 kg/m" & chr(179)
    alphaSelected = Array(0.05, 0.11, 0.28, 0.68, 0.9, 0.93, 0.96)
    Case Is = "50mm fibreglass, 48 kg/m" & chr(179)
    alphaSelected = Array(0.1, 0.17, 0.86, 1#, 1#, 1#, 1#)
    Case Is = "75mm fibreglass, 48 kg/m" & chr(179)
    alphaSelected = Array(0.3, 0.53, 1#, 1#, 1#, 1#, 1#)
    Case Is = "100mm fibreglass, 48 kg/m" & chr(179)
    alphaSelected = Array(0.5, 0.84, 1#, 1#, 1#, 1#, 0.97)
    End Select
    
    'put in array
    For i = LBound(alpha2) To UBound(alpha2)
    alpha2(i) = alphaSelected(i)
    Next i

CalculateAlphaTotal

PreviewInsertionLoss

End Sub

Private Sub cboxUnlinedMaterial_Change()

Dim alphaSelected As Variant
    
    Select Case Me.cBoxUnlinedMaterial.Value
    Case Is = "Concrete"
    alphaSelected = Array(0.01, 0.01, 0.01, 0.02, 0.02, 0.02, 0.03)
    Case Is = "Bare Sheet Metal"
    alphaSelected = Array(0.04, 0.04, 0.04, 0.05, 0.05, 0.05, 0.07)
    End Select
    
    'put in array
    For i = LBound(alpha1) To UBound(alpha1)
    alpha1(i) = alphaSelected(i)
    Next i
    
CalculateAlphaTotal

PreviewInsertionLoss

End Sub



Private Sub cBoxWallEffect_Change()
CalculateWallEffect
PreviewInsertionLoss
End Sub

Private Sub CommandButton1_Click()
PreviewInsertionLoss
End Sub

Private Sub optEndEnd_Click()
PreviewInsertionLoss
End Sub

Private Sub optEndSide_Click()
PreviewInsertionLoss
End Sub

Private Sub optInletCentre_Click()
PreviewInsertionLoss
End Sub

Private Sub optInletCorner_Click()
PreviewInsertionLoss
End Sub

Private Sub sbUnlinedPercent_Change()
Me.txtUnlinedPercent.Value = Me.sbUnlinedPercent.Value
CalculateAlphaTotal
PreviewInsertionLoss
End Sub

Private Sub txtH_Change()
CalculateVolume
CalculateSurfaceArea
PreviewInsertionLoss
End Sub

Private Sub txtHorizontalOffset_Change()
Calc_R_and_Theta
PreviewInsertionLoss
End Sub

Private Sub txtInL_Change()
CalculateInletArea
CalculateCutoffFrequency
PreviewInsertionLoss
End Sub

Private Sub txtInW_Change()
CalculateInletArea
CalculateCutoffFrequency
PreviewInsertionLoss
End Sub

Private Sub txtL_Change()
CalculateVolume
CalculateSurfaceArea
PreviewInsertionLoss
End Sub

Private Sub txtOutL_Change()
CalculateOutletArea
PreviewInsertionLoss
End Sub

Private Sub txtOutW_Change()
CalculateOutletArea
PreviewInsertionLoss
End Sub

Private Sub txtVerticalOffset_Change()
Calc_R_and_Theta
PreviewInsertionLoss
End Sub

Private Sub txtW_Change()
CalculateVolume
CalculateSurfaceArea
PreviewInsertionLoss
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
Me.lblTheta.Caption = Me.lblTheta.Caption & ChrW(952)
End Sub

Private Sub UserForm_Initialize()
PopulateComboBox
Calculate_EVERYTHING
PreviewInsertionLoss
End Sub

Sub Calculate_EVERYTHING()

CalculateVolume
CalculateInletArea
CalculateOutletArea
Calc_R_and_Theta
CalculateSurfaceArea
CalculateCutoffFrequency
CalculateAlphaTotal
End Sub

Sub PopulateComboBox()

Me.cBoxLining.AddItem ("Concrete")
Me.cBoxLining.AddItem ("Bare Sheet Metal")
Me.cBoxLining.AddItem ("25mm fibreglass, 48 kg/m" & chr(179))
Me.cBoxLining.AddItem ("50mm fibreglass, 48 kg/m" & chr(179))
Me.cBoxLining.AddItem ("75mm fibreglass, 48 kg/m" & chr(179))
Me.cBoxLining.AddItem ("100mm fibreglass, 48 kg/m" & chr(179))

Me.cBoxUnlinedMaterial.AddItem ("Concrete")
Me.cBoxUnlinedMaterial.AddItem ("Bare Sheet Metal")

Me.cBoxWallEffect.AddItem ("0 - None")
Me.cBoxWallEffect.AddItem ("1 - 25mm, 40kg/m" & chr(179) & " (Fabric Facing)")
Me.cBoxWallEffect.AddItem ("2 - 50mm, 40kg/m" & chr(179) & " (Fabric Facing)")
Me.cBoxWallEffect.AddItem ("3 - 100mm, 40kg/m" & chr(179) & " (Perf. Facing)")
Me.cBoxWallEffect.AddItem ("4 - 200mm, 40kg/m" & chr(179) & " (Perf. Facing)")
Me.cBoxWallEffect.AddItem ("5 - 100mm (Tuned, No Media)")
Me.cBoxWallEffect.AddItem ("6 - 100mm, 40kg/m" & chr(179) & " (Double Solid Metal)")

'Default values
Me.cBoxLining.Value = "25mm fibreglass, 48 kg/m" & chr(179)
Me.cBoxUnlinedMaterial.Value = "Bare Sheet Metal"
Me.cBoxWallEffect.Value = ("1 - 25mm, 40kg/m" & chr(179) & " (Fabric Facing)")
    
End Sub

Sub CalculateVolume()
    If Me.txtL.Value <> "" And Me.txtW.Value <> "" And Me.txtH.Value <> "" Then
    Vol = (Me.txtL.Value / 1000) * (Me.txtW.Value / 1000) * (Me.txtH.Value / 1000)
    Me.txtV.Value = Round(Vol, 2)
    End If
End Sub

Sub CalculateInletArea()
    If Me.txtInL.Value <> "" And Me.txtInW.Value <> "" Then
    InletArea = (Me.txtInL.Value / 1000) * (Me.txtInW.Value / 1000)
    Me.txtInletArea.Value = Round(InletArea, 2)
    End If
End Sub

Sub CalculateOutletArea()
    If Me.txtInL.Value <> "" And Me.txtInW.Value <> "" Then
    OutletArea = (Me.txtOutL.Value / 1000) * (Me.txtOutW.Value / 1000)
    Me.txtOutletArea.Value = Round(OutletArea, 2)
    End If
End Sub

Sub Calc_R_and_Theta()

Dim R As Single
Dim theta As Long

    If Me.txtHorizontalOffset.Value <> "" And Me.txtVerticalOffset.Value <> "" Then
    R = PlenumDistanceR(Me.txtHorizontalOffset.Value, Me.txtVerticalOffset.Value, Me.txtL.Value)
    theta = PlenumAngleTheta(Me.txtL.Value, R)
    
        'warning for >45 degrees
        If theta > 45 Then
        msg = MsgBox("Please note that the ASHRAE method only allows offset angles up to 45 degrees. " & chr(10) _
        & "Consider using the End In -> Side out (90 degree) option.", vbOKOnly, "Warning: ASHRAE Plenum Method")
        End If
    
    Me.txtR.Value = Round(R, 1)
    Me.txtTheta.Value = theta
    End If
    
End Sub


Sub CalculateCutoffFrequency()

    If Me.txtInL.Value <> "" And Me.txtInW.Value <> "" Then
    f_co = PlenumCutoffFrequency(CSng(Me.txtInL.Value / 1000), CSng(Me.txtInW.Value / 1000)) 'central function in module NoiseFunctions
    Me.txtCutoffFrequency.Value = Round(f_co, 1)
    End If
    
End Sub

Sub CalculateSurfaceArea()
    If IsNumeric(Me.txtL.Value) And IsNumeric(Me.txtH.Value) And IsNumeric(Me.txtInletArea.Value) And IsNumeric(Me.txtOutletArea.Value) Then
    SA = PlenumSurfaceArea(Me.txtL.Value, Me.txtH.Value, Me.txtW.Value, Me.txtInletArea, Me.txtOutletArea)
    Me.txtSurfaceArea.Value = Round(SA, 2)
    End If
End Sub

Sub CalculateAlphaTotal()
Dim SA_Unlined As Double
Dim SA_Lined As Double
    If Me.txtSurfaceArea.Value <> 0 And Me.txtInletArea.Value <> 0 And Me.txtOutletArea.Value <> 0 And Me.txtSurfaceArea.Value <> "" And Me.txtInletArea.Value <> "" And Me.txtOutletArea.Value <> "" Then
    SA_Unlined = SA * (Me.txtUnlinedPercent.Value / 100)
    SA_Lined = SA - SA_Unlined
        For i = LBound(alphaTotal) To UBound(alphaTotal)
            If checkAlpha = True Then
            alphaTotal(i) = ((((InletArea + OutletArea + SA_Unlined) * alpha1(i))) + ((SA_Lined) * alpha2(i))) / SA 'surface area doesn't include inlet and outlet areas
            End If
        Next i
        
    Me.txt63_alpha.Value = Round(alphaTotal(0), 2)
    Me.txt125_alpha.Value = Round(alphaTotal(1), 2)
    Me.txt250_alpha.Value = Round(alphaTotal(2), 2)
    Me.txt500_alpha.Value = Round(alphaTotal(3), 2)
    Me.txt1k_alpha.Value = Round(alphaTotal(4), 2)
    Me.txt2k_alpha.Value = Round(alphaTotal(5), 2)
    Me.txt4k_alpha.Value = Round(alphaTotal(6), 2)
    End If
    
End Sub

Sub CalculateWallEffect()
Dim i As Integer 'wall effect index
i = CInt(Left(Me.cBoxWallEffect, 1))
Me.txt50We.Value = PlenumWallEffect(50, i) * -1
Me.txt63We.Value = PlenumWallEffect(63, i) * -1
Me.txt80We.Value = PlenumWallEffect(80, i) * -1
Me.txt100We.Value = PlenumWallEffect(100, i) * -1
Me.txt125We.Value = PlenumWallEffect(125, i) * -1
Me.txt160We.Value = PlenumWallEffect(160, i) * -1
Me.txt200We.Value = PlenumWallEffect(200, i) * -1
Me.txt250We.Value = PlenumWallEffect(250, i) * -1
Me.txt315We.Value = PlenumWallEffect(315, i) * -1
Me.txt400We.Value = PlenumWallEffect(400, i) * -1
Me.txt500We.Value = PlenumWallEffect(500, i) * -1
End Sub

Sub PreviewInsertionLoss()
Dim IL63 As Single
Dim IL125 As Single
Dim IL250 As Single
Dim IL500 As Single
Dim IL1k As Single
Dim IL2k As Single
Dim IL4k As Single
Dim SplitStr() As String
Dim LiningType As String
Dim Q As Integer
Dim ApplyElbow As Boolean
    If CheckFormFields = True Then

    SplitStr = Split(Me.cBoxLining.Value, ",", Len(Me.cBoxLining.Value), vbTextCompare)
    LiningType = SplitStr(0) 'first element, before the comma
    
    If Me.optInletCorner.Value = True Then
    Q = 4
    Else
    Q = 2
    End If
    
    'PlenumElbowEffect
    If Me.optEndSide.Value = True Then
    ApplyElbow = True
    Else
    ApplyElbow = False
    End If
    
    IL63 = PlenumLoss_ASHRAE("63", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL125 = PlenumLoss_ASHRAE("125", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL250 = PlenumLoss_ASHRAE("250", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL500 = PlenumLoss_ASHRAE("500", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL1k = PlenumLoss_ASHRAE("1k", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL2k = PlenumLoss_ASHRAE("2k", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    IL4k = PlenumLoss_ASHRAE("4k", Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtInL.Value, Me.txtInW.Value, Me.txtOutL.Value, Me.txtOutW.Value, Q, Me.txtVerticalOffset.Value, Me.txtHorizontalOffset.Value, LiningType, Me.cBoxUnlinedMaterial.Value, Me.cBoxWallEffect.Value, ApplyElbow, Me.sbUnlinedPercent.Value)
    
    Me.txt63.Value = Round(IL63, 1)
    Me.txt125.Value = Round(IL125, 1)
    Me.txt250.Value = Round(IL250, 1)
    Me.txt500.Value = Round(IL500, 1)
    Me.txt1k.Value = Round(IL1k, 1)
    Me.txt2k.Value = Round(IL2k, 1)
    Me.txt4k.Value = Round(IL4k, 1)
    End If
End Sub

Function CheckFormFields() As Boolean
    If Me.txtL.Value = "" Or Me.txtW.Value = "" Or Me.txtH.Value = "" Or _
    Me.txtInL.Value = "" Or Me.txtInW.Value = "" Or _
    Me.txtOutL.Value = "" Or Me.txtOutW.Value = "" Or _
    Me.cBoxLining.Value = "" Or Me.cBoxUnlinedMaterial.Value = "" Or _
    Me.txtVerticalOffset.Value = "" Or Me.txtHorizontalOffset.Value = "" Then
    'Blank values
    CheckFormFields = False
    Else
    CheckFormFields = True
    End If
End Function

Function checkAlpha() As Boolean

On Error GoTo catch

checkAlpha = True

    For i = LBound(alphaTotal) To UBound(alphaTotal)
    'Debug.Print alpha1(i)
    'Debug.Print alpha2(i)
    Next i
    
Exit Function

catch:
    checkAlpha = False
    
End Function
