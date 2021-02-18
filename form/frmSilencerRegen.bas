VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSilencerRegen 
   Caption         =   "Silencer Regenerated Noise"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   OleObjectBlob   =   "frmSilencerRegen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSilencerRegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
PreviewRegen
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
PreviewRegen

'set global variables

    If Me.optFantech.Value = True Then
    RegenMode = "Fantech"
    numModules = Me.txtNumModules.Value
    ElseIf Me.optNAP.Value = True Then
    RegenMode = "NAP"
    Else 'ERROR
    End If

    'flowrate
    If Me.txtFlowRate.Value <> "" Then
        If Me.optLitres.Value = True Then
        FlowUnitsM3ps = False
        Else
        FlowUnitsM3ps = True
        End If
    FlowRate = Me.txtFlowRate.Value
    End If

PFA = Me.txtFA.Value
SilH = ScreenInput(Me.txtH.Value)
SilW = ScreenInput(Me.txtW.Value)
SilencerModel = Me.txtTypeCode
numModules = Me.sbModules.Value
btnOkPressed = True
Me.Hide
Unload Me

End Sub


Private Sub optFantech_Click()
Me.txtNumModules.Enabled = True
Me.sbModules.Enabled = True
UpdatePicture ("img\FantechRegen.JPG")
PreviewRegen
End Sub

Private Sub optNAP_Click()
Me.txtNumModules.Enabled = False
Me.sbModules.Enabled = False
UpdatePicture ("img\NAPregen.JPG")
PreviewRegen
End Sub

Private Sub optLitres_Click()
PreviewRegen
End Sub

Private Sub optMetresCubed_Click()
PreviewRegen
End Sub

Sub UpdatePicture(FilePath As String)
GetSettings
    If TestLocation(ROOTPATH & "\" & FilePath) = True Then
    Me.imgFigure.Picture = LoadPicture(ROOTPATH & "\" & FilePath)
    End If
End Sub

Private Sub sbModules_Change()
Me.txtNumModules.Value = Me.sbModules.Value
PreviewRegen
End Sub

Private Sub txtFlowRate_Change()
PreviewRegen
End Sub

Private Sub txtH_Change()
PreviewRegen
End Sub


Private Sub txtTypeCode_Change()
PreviewRegen
End Sub

Private Sub txtW_Change()
PreviewRegen
End Sub

Sub PreviewRegen()
Dim FlowrateM3ps As Double
Dim DuctAreaMsq As Double
Dim VelocityMps As Double

Dim Model As String
Dim Length As Integer
Dim FreeArea As Integer
Dim SplitModel() As String

Model = Me.txtTypeCode.Value
    
    'fill out model details
    If Me.optFantech.Value = True Then 'FANTECH
        Select Case Mid(Me.txtTypeCode.Value, 3, 2)
        Case Is = "07"
        Me.txtFA.Value = 27
        Case Is = "10"
        Me.txtFA.Value = 33
        Case Is = "12"
        Me.txtFA.Value = 38
        Case Is = "15"
        Me.txtFA.Value = 43
        Case Is = "17"
        Me.txtFA.Value = 47
        Case Is = "20"
        Me.txtFA.Value = 50
        Case Is = "22"
        Me.txtFA.Value = 53
        Case Is = "25"
        Me.txtFA.Value = 56
        Case Else
        Me.txtFA.Value = "-"
        End Select
        
        Select Case UCase(Right(Me.txtTypeCode.Value, 1))
        Case Is = "A"
        Me.txtL.Value = 600
        Case Is = "B"
        Me.txtL.Value = 900
        Case Is = "C"
        Me.txtL.Value = 1200
        Case Is = "D"
        Me.txtL.Value = 1500
        Case Is = "E"
        Me.txtL.Value = 1800
        Case Is = "F"
        Me.txtL.Value = 2100
        Case Is = "G"
        Me.txtL.Value = 2400
        Case Else
        Me.txtL.Value = 0
        End Select
    Else 'NAP
    SplitModel = Split(Model, "/", Len(Model), vbTextCompare)
        'check for array size
        If UBound(SplitModel) >= 1 Then
            'length, in cm, convert to mm
            If IsNumeric(SplitModel(1)) Then
            Me.txtL.Value = CInt(SplitModel(1)) * 10
            End If
        'free area
        Me.txtFA.Value = Right(SplitModel(0), 2)
        Else
        Me.txtL.Value = "-"
        Me.txtFA.Value = "-"
        End If
    End If

    'show the match has been found
    If IsNumeric(Me.txtL.Value) And Me.txtL.Value <> 0 Then
    Me.lblMatchL.Visible = True
    Else
    Me.lblMatchL.Visible = False
    End If
        'show the match has been found
    If IsNumeric(Me.txtFA.Value) And Me.txtFA.Value <> 0 Then
    Me.lblMatchFA.Visible = True
    Else
    Me.lblMatchFA.Visible = False
    End If

'Calculate Regen Sound Power

    'calculate area and velocity
    If IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
    DuctAreaMsq = (Me.txtW.Value * Me.txtH.Value) / 1000000 'area in m^2
    Me.txtDuctArea.Value = Round(DuctAreaMsq, 3)
        If IsNumeric(Me.txtFlowRate.Value) And IsNumeric(Me.txtFA.Value) Then
            'RegenNoise
            If Me.optLitres.Value = True Then
            FlowrateM3ps = CDbl(Me.txtFlowRate.Value) / 1000
            Me.txtVelocity.Value = Round((Me.txtFlowRate.Value / 1000) / DuctAreaMsq, 1)
            Else 'metres cubed per second
            FlowrateM3ps = Me.txtFlowRate.Value
            Me.txtVelocity.Value = Round(Me.txtFlowRate.Value / DuctAreaMsq, 2)
            End If
        
            'preview sound power values
            If Me.optFantech.Value = True Then
            Me.txt63.Value = Round(ScreenInput(FantechAttenRegen("63", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt125.Value = Round(ScreenInput(FantechAttenRegen("125", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt250.Value = Round(ScreenInput(FantechAttenRegen("250", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt500.Value = Round(ScreenInput(FantechAttenRegen("500", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt1k.Value = Round(ScreenInput(FantechAttenRegen("1k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt2k.Value = Round(ScreenInput(FantechAttenRegen("2k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt4k.Value = Round(ScreenInput(FantechAttenRegen("4k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Me.txt8k.Value = Round(ScreenInput(FantechAttenRegen("8k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtNumModules.Value)), 1)
            Else 'nap
            Me.txt63.Value = Round(ScreenInput(NAPAttenRegen("63", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt125.Value = Round(ScreenInput(NAPAttenRegen("125", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt250.Value = Round(ScreenInput(NAPAttenRegen("250", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt500.Value = Round(ScreenInput(NAPAttenRegen("500", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt1k.Value = Round(ScreenInput(NAPAttenRegen("1k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt2k.Value = Round(ScreenInput(NAPAttenRegen("2k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt4k.Value = Round(ScreenInput(NAPAttenRegen("4k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            Me.txt8k.Value = Round(ScreenInput(NAPAttenRegen("8k", FlowrateM3ps, Me.txtFA.Value, Me.txtW, Me.txtH.Value, Me.txtTypeCode.Value)), 1)
            End If
        Else 'problem! show nothing
        Me.txt63.Value = "-"
        Me.txt125.Value = "-"
        Me.txt250.Value = "-"
        Me.txt500.Value = "-"
        Me.txt1k.Value = "-"
        Me.txt2k.Value = "-"
        Me.txt4k.Value = "-"
        Me.txt8k.Value = "-"
        End If
    Else
    Me.txtDuctArea.Value = "-"
    End If
End Sub

