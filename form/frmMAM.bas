VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMAM 
   Caption         =   "Mass-Air-Mass Resonance Calculator"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "frmMAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Basics#mass-air-mass-calculator")
End Sub

Private Sub btnOK_Click()
Dim Side1 As String
Dim Side2 As String
    
    'set public variables
    If IsNumeric(Me.txtSurfDensity1.Value) Then
    MAM_M1 = Round(Me.txtSurfDensity1.Value, 2)
    End If
    
    If IsNumeric(Me.txtSurfDensity2.Value) Then
    MAM_M2 = Round(Me.txtSurfDensity2.Value, 2)
    End If
    
    If IsNumeric(Me.txtCavityWidth.Value) Then
    MAM_Width = Me.txtCavityWidth.Value 'in mm
    End If

    'determine dsescription strings for commenting
    'side 1
    If Me.optM1Other.Value = True Then
    Side1 = "Side 1: " & Me.txtSurfDensity1.Value & "kg/m3"
    Else
    Side1 = "Side 1: " & Me.txtThickness1.Value & "mm " & GetDescription("m1") & " " _
        & Me.txtSurfDensity1.Value & "kg/m3"
    End If
    'side 2
    If Me.optM2Other.Value = True Then
    Side2 = "Side 2: " & Me.txtSurfDensity2.Value & "kg/m3"
    Else
    Side2 = "Side 2: " & Me.txtThickness2 & "mm " & GetDescription("m2") & " " _
        & Me.txtSurfDensity2.Value & "kg/m3"
    End If
'set the string
MAM_Description = Side1 & chr(10) & Side2
    
btnOkPressed = True
Unload Me
End Sub

Function GetDescription(GroupName As String)
    For i = 0 To Me.Controls.Count - 1
        If TypeName(Me.Controls(i)) = "OptionButton" Then
            If Me.Controls(i).GroupName = GroupName And _
                Me.Controls(i).Value = True Then
            GetDescription = Me.Controls(i).Caption
            Exit Function
            End If
        End If
    Next i
End Function

Private Sub optM1FRPB_Click()
PreviewValues
End Sub

Private Sub optM1Glass_Click()
PreviewValues
End Sub

Private Sub optM1Other_Click()
PreviewValues
End Sub

Private Sub optM1PB_Click()
PreviewValues
End Sub

Private Sub optM2FRPB_Click()
PreviewValues
End Sub

Private Sub optM2Glass_Click()
PreviewValues
End Sub

Private Sub optM2Other_Click()
PreviewValues
End Sub

Private Sub optM2PB_Click()
PreviewValues
End Sub

Private Sub spinLayer1_Change()
Me.txtNumLayers1.Value = "x" & Me.spinLayer1.Value
PreviewValues
End Sub

Private Sub spinLayer2_Change()
Me.txtNumLayers2.Value = "x" & Me.spinLayer2.Value
PreviewValues
End Sub

Private Sub txtAirTemp_Change()
PreviewValues
End Sub

Private Sub txtCavityWidth_Change()
PreviewValues
End Sub

Private Sub txtLayerThickness1_Change()
PreviewValues
End Sub

Private Sub txtLayerThickness2_Change()
PreviewValues
End Sub

Private Sub txtSurfDensity1_Change()
    If Me.optM1Other.Value = True Then
    PreviewValues
    End If
End Sub

Private Sub txtSurfDensity2_Change()
    If Me.optM2Other.Value = True Then
    PreviewValues
    End If
End Sub

Private Sub txtThickness1_Change()
PreviewValues
End Sub

Private Sub txtThickness2_Change()
PreviewValues
End Sub

Private Sub UserForm_Activate()
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewValues
End Sub

Sub PreviewValues()
Dim MAM As Double
Dim m1 As Double
Dim m2 As Double
Dim D As Double
Dim thickness1 As Double
Dim thickness2 As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'enable/disable options - side 1
    If Me.optM1Other.Value = True Then
    Me.txtLayerThickness1.Enabled = False
    Me.txtLayerThickness1.BackColor = &H80000005
    Me.txtSurfDensity1.Enabled = True
    Me.txtSurfDensity1.BackColor = &HC0FFFF
    
    'set variable m1
    m1 = CheckNumericValue(Me.txtSurfDensity1.Value)

    Else 'calculate surface density based on thickness
    Me.txtLayerThickness1.Enabled = True
    Me.txtLayerThickness1.BackColor = &HC0FFFF
    Me.txtSurfDensity1.Enabled = False
    Me.txtSurfDensity1.BackColor = &H80000005
    
        'Calculate Mass 1
        If IsNumeric(Me.txtLayerThickness1.Value) Then
        Me.txtThickness1.Value = Me.txtLayerThickness1.Value * Me.spinLayer1
        thickness1 = Me.txtThickness1.Value / 1000
            If Me.optM1Glass.Value = True Then
            m1 = MaterialDensity("Glass") * thickness1
            ElseIf Me.optM1PB.Value = True Then
            m1 = MaterialDensity("PB") * thickness1
            ElseIf Me.optM1FRPB.Value = True Then
            m1 = MaterialDensity("FRPB") * thickness1
            End If
        Me.txtSurfDensity1.Value = m1
        End If
    
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'enable/disable options - side 2
    If Me.optM2Other.Value = True Then
    Me.txtLayerThickness2.Enabled = False
    Me.txtLayerThickness2.BackColor = &H80000005
    Me.txtSurfDensity2.Enabled = True
    Me.txtSurfDensity2.BackColor = &HC0FFFF
    m2 = CheckNumericValue(Me.txtSurfDensity1.Value)
    
    Else 'calculate surface density based on thickness
    Me.txtLayerThickness2.Enabled = True
    Me.txtLayerThickness2.BackColor = &HC0FFFF
    Me.txtSurfDensity2.Enabled = False
    Me.txtSurfDensity2.BackColor = &H80000005
    
        'Calculate Mass 2
        If IsNumeric(Me.txtLayerThickness2.Value) Then
        Me.txtThickness2.Value = Me.txtLayerThickness2.Value * Me.spinLayer2
        thickness2 = Me.txtThickness2.Value / 1000
            If Me.optM2Glass.Value = True Then
            m2 = MaterialDensity("Glass") * thickness2
            ElseIf Me.optM2PB.Value = True Then
            m2 = MaterialDensity("PB") * thickness2
            ElseIf Me.optM2FRPB.Value = True Then
            m2 = MaterialDensity("FRPB") * thickness2
            End If
        Me.txtSurfDensity2.Value = m2
        End If
    End If


    'calculate speed of sound
    If IsNumeric(Me.txtAirTemp) Then
    Me.txtSOS.Value = Round(SpeedOfSound(Me.txtAirTemp.Value), 1)
    Else
    Me.txtSOS.Value = "-"
    End If
    
    'Calculate Mass-air-mass
    If IsNumeric(Me.txtCavityWidth.Value) And IsNumeric(Me.txtSurfDensity1.Value) And IsNumeric(Me.txtSurfDensity2.Value) Then
    MAM = MassAirMass(m1, m2, Me.txtCavityWidth.Value, Me.txtAirTemp.Value)
    Me.txtMAM.Value = Round(MAM, 1)
    Else
    Me.txtMAM.Value = "-"
    End If

End Sub

Function MaterialDensity(MaterialName As String)
    Select Case MaterialName
    Case Is = "Glass"
    MaterialDensity = 2433.3 'kg/m3
    Case Is = "PB"
    MaterialDensity = 646.2 'kg/m3
    Case Is = "FRPB"
    MaterialDensity = 807.7 'kg/m3
    End Select
End Function
