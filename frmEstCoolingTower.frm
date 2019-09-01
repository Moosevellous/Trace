VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstCoolingTower 
   Caption         =   "SWL Estimator - Cooling Tower"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12450
   OleObjectBlob   =   "frmEstCoolingTower.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstCoolingTower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
    
    'check for radio boxes
    If Me.optCentrifugalType.Value = False And Me.optPropellerType.Value = False Then
    msg = MsgBox("Cooling Tower type not selected", vbOKOnly, "You must chooooooose")
    Else
            If IsNumeric(Me.txtPower.Value) = False Then
            btnOkPressed = False
            Else
            btnOkPressed = True
            CTPower = Me.txtPower.Value
            CTEqn = Me.lblEqn.Caption
            
                If Me.optCentrifugalType.Value = True Then
                CT_Type = "Centrifugal" 'global variable
                ElseIf Me.optPropellerType.Value = True Then
                CT_Type = "Propelller" 'global variable
                End If
                
                'set directional effects array (global variable)
                If Me.chkDirectionalEffects.Value = True Then
                CT_Dir_checked = True
                CT_Direction(0) = Me.txt31dir.Value
                CT_Direction(1) = Me.txt63dir.Value
                CT_Direction(2) = Me.txt125dir.Value
                CT_Direction(3) = Me.txt250dir.Value
                CT_Direction(4) = Me.txt500dir.Value
                CT_Direction(5) = Me.txt1kdir.Value
                CT_Direction(6) = Me.txt2kdir.Value
                CT_Direction(7) = Me.txt4kdir.Value
                CT_Direction(8) = Me.txt8kdir.Value
                CT_Direction(9) = Me.cboxDir.Value & "; Face: " & Me.cBoxSide.Value 'description
                End If
                
                
            End If
        Me.Hide
    End If


End Sub

Private Sub cboxDir_Change()
    If cboxDir.Value = "" Then
    Me.cBoxSide.Enabled = False
    'Me.chkDirectionalEffects.Value = False
    Else
    Me.cBoxSide.Enabled = True
    End If
End Sub

Private Sub cBoxSide_Change()
Dim CTcode As String
Dim DirArray() As Variant

CTcode = Left(Me.cboxDir.Value, 1) 'a/b/c/d options from cboxDir

ReDim Preserve DirArray(0 To 8)

    Select Case CTcode
    
    Case Is = "a"
    
        Select Case cBoxSide.Value
        Case Is = "Front"
        DirArray = Array(3, 3, 2, 3, 4, 3, 3, 4, 4)
        Case Is = "Side"
        DirArray = Array(0, 0, 0, -2, -3, -4, -5, -5, -5)
        Case Is = "Rear"
        DirArray = Array(0, 0, -1, -2, -3, -4, -5, -6, -6)
        Case Is = "Top"
        DirArray = Array(-3, -3, -2, 0, 1, 2, 3, 4, 5)
        End Select
    
    Case Is = "b"
    
        Select Case cBoxSide.Value
        Case Is = "Front"
        DirArray = Array(2, 2, 4, 6, 6, 5, 5, 5, 5)
        Case Is = "Side"
        DirArray = Array(1, 1, 1, -2, -5, -5, -5, -5, -4)
        Case Is = "Rear"
        DirArray = Array(-3, -3, -4, -7, -7, -7, -8, -11, -3)
        Case Is = "Top"
        DirArray = Array(-5, -5, -5, -5, -2, 0, 0, 2, 4)
        End Select
        
    Case Is = "c"
    
        Select Case cBoxSide.Value
        Case Is = "Front"
        DirArray = Array(0, 0, 0, 1, 2, 2, 2, 3, 3)
        Case Is = "Side"
        DirArray = Array(-2, -2, -2, -3, -4, -4, -5, -6, -6)
        Case Is = "Top"
        DirArray = Array(3, 3, 3, 3, 2, 2, 2, 1, 1)
        Case Is = "Rear"
        msg = MsgBox("This option does not exist in the reference source. Refer to the wiki.", vbOKOnly, "NOPE!")
        DirArray = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
        End Select

    
    Case Is = "d"
        If cBoxSide.Value = "Top" Then
        DirArray = Array(2, 2, 2, 3, 3, 4, 4, 5, 5)
        ElseIf cBoxSide.Value = "" Then
        Erase DirArray
        Else 'any other side
        DirArray = Array(-1, -1, -1, -2, -2, -3, -3, -4, -4)
        End If
        
    End Select
   
'populate fields
txt31dir.Value = DirArray(0)
txt63dir.Value = DirArray(1)
txt125dir.Value = DirArray(2)
txt250dir.Value = DirArray(3)
txt500dir.Value = DirArray(4)
txt1kdir.Value = DirArray(5)
txt2kdir.Value = DirArray(6)
txt4kdir.Value = DirArray(7)
txt8kdir.Value = DirArray(8)

'calculate sound power spectrum
SetEqn
    
End Sub

Private Sub chkDirectionalEffects_Click()
    
    If Me.chkDirectionalEffects.Value = True Then
    CT_Dir_checked = True 'public variable
    cboxDir.Enabled = True
    cBoxSide.Enabled = True
    Else
    CT_Dir_checked = False 'public variable
    cboxDir.Enabled = False
    cBoxSide.Enabled = False
    End If

SetEqn

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub optCentrifugalType_Click()
    If optCentrifugalType.Value = True Then
    SelectCentrifugalType
    Else
    SelectPropellerType
    End If
    
    SetEqn

End Sub

Private Sub optPropellerType_Click()
    If optPropellerType.Value = True Then
    SelectPropellerType
    Else
    SelectCentrifugalType
    End If
    
    SetEqn

End Sub

Private Sub txtPower_Change()
SetEqn
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False 'default
SetEqn
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
Me.cboxDir.Clear
Me.cboxDir.AddItem ("")
Me.cboxDir.AddItem ("a - Centrifugal fan blow through")
Me.cboxDir.AddItem ("b - Axial flow, blow through type")
Me.cboxDir.AddItem ("c - Induced draft, propeller type")
Me.cboxDir.AddItem ("d - Underflow forved draft propeller type")
Me.cBoxSide.Clear
Me.cBoxSide.AddItem ("")
Me.cBoxSide.AddItem ("Front")
Me.cBoxSide.AddItem ("Side")
Me.cBoxSide.AddItem ("Rear")
Me.cBoxSide.AddItem ("Top")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Centralised macros start here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SelectPropellerType()
Me.optAbove60kW.Value = False
Me.optUnder60kW.Value = False
CT_Type = "Propeller" 'set CT type (public variable)
SetCorrection
Me.txt31adj.Value = CT_Correction(0)
Me.txt63adj.Value = CT_Correction(1)
Me.txt125adj.Value = CT_Correction(2)
Me.txt250adj.Value = CT_Correction(3)
Me.txt500adj.Value = CT_Correction(4)
Me.txt1kadj.Value = CT_Correction(5)
Me.txt2kadj.Value = CT_Correction(6)
Me.txt4kadj.Value = CT_Correction(7)
Me.txt8kadj.Value = CT_Correction(8)
End Sub

Sub SelectCentrifugalType()
Me.optAbove75kW.Value = False
Me.optUnder75kW.Value = False
CT_Type = "Centrifugal" 'set CT type (public variable)
SetCorrection
Me.txt31adj.Value = CT_Correction(0)
Me.txt63adj.Value = CT_Correction(1)
Me.txt125adj.Value = CT_Correction(2)
Me.txt250adj.Value = CT_Correction(3)
Me.txt500adj.Value = CT_Correction(4)
Me.txt1kadj.Value = CT_Correction(5)
Me.txt2kadj.Value = CT_Correction(6)
Me.txt4kadj.Value = CT_Correction(7)
Me.txt8kadj.Value = CT_Correction(8)
End Sub

Sub SetCorrection()
'CT_Type set from global
    If CT_Type = "Centrifugal" Then
    CT_Correction(0) = -6
    CT_Correction(1) = -6
    CT_Correction(2) = -8
    CT_Correction(3) = -10
    CT_Correction(4) = -11
    CT_Correction(5) = -13
    CT_Correction(6) = -12
    CT_Correction(7) = -18
    CT_Correction(8) = -25
    ElseIf CT_Type = "Propeller" Then
    CT_Correction(0) = -8
    CT_Correction(1) = -5
    CT_Correction(2) = -5
    CT_Correction(3) = -8
    CT_Correction(4) = -11
    CT_Correction(5) = -15
    CT_Correction(6) = -18
    CT_Correction(7) = -21
    CT_Correction(8) = -29
    Else
    'do nothing
    End If
End Sub

Sub SetEqn()
    If optPropellerType.Value = True Then
    
        If txtPower.Value > 75 Then
        lblEqn.Caption = "Lw=96+10log(kW)"
        optAbove75kW.Value = True
        optUnder75kW.Value = False
        Else
        lblEqn.Caption = "Lw=100+8log(kW)"
        optAbove75kW.Value = False
        optUnder75kW.Value = True
        End If
        
        Call CalcLw(Me.txtPower)
        
    ElseIf optCentrifugalType.Value = True Then
    
        If txtPower.Value <= 60 Then 'up to 60kW
        lblEqn.Caption = "Lw=85+11log(kW)"
        optAbove60kW.Value = True
        Else
        lblEqn.Caption = "Lw=93+7log(kW)"
        optUnder60kW.Value = True
        End If
        
        Call CalcLw(Me.txtPower)
        
    End If
End Sub

Sub CalcLw(kW As Variant)
Dim LwOverall As Single
Dim CheckBlankVal As Boolean


    If IsNumeric(kW) Then
        If CT_Type = "Centrifugal" Then
            If kW > 75 Then
            LwOverall = 96 + 10 * Application.WorksheetFunction.Log10(txtPower.Value)
            Else
            LwOverall = 100 + 8 * Application.WorksheetFunction.Log10(txtPower.Value)
            End If
        ElseIf CT_Type = "Propeller" Then
            If kW > 60 Then
            LwOverall = 85 + 11 * Application.WorksheetFunction.Log10(txtPower.Value)
            Else
            LwOverall = 93 + 7 * Application.WorksheetFunction.Log10(txtPower.Value)
            End If
        Else 'error!
        txtLw.Value = ""
        End If
    txtLw.Value = Round(LwOverall, 1)
    Else
    txtLw.Value = ""
    End If
    
    SetCorrection
    
    'Directional effects
    
    CheckBlankVal = checkTXT 'check for blank directional array
    
    If LwOverall = Empty Then LwOverall = 0 'check for blank
    
    'Spectrum
    If Me.chkDirectionalEffects.Value = True And Me.cboxDir.Value <> "" And Me.cBoxSide.Value <> "" And CheckBlankVal <> False Then
    
    txt31.Value = Round(LwOverall + CT_Correction(0) + txt31dir.Value, 0)
    txt63.Value = Round(LwOverall + CT_Correction(1) + txt63dir.Value, 0)
    txt125.Value = Round(LwOverall + CT_Correction(2) + txt125dir.Value, 0)
    txt250.Value = Round(LwOverall + CT_Correction(3) + txt250dir.Value, 0)
    txt500.Value = Round(LwOverall + CT_Correction(4) + txt500dir.Value, 0)
    txt1k.Value = Round(LwOverall + CT_Correction(5) + txt1kdir.Value, 0)
    txt2k.Value = Round(LwOverall + CT_Correction(6) + txt2kdir.Value, 0)
    txt4k.Value = Round(LwOverall + CT_Correction(7) + txt4kdir.Value, 0)
    txt8k.Value = Round(LwOverall + CT_Correction(8) + txt8kdir.Value, 0)
    Else
    txt31.Value = Round(LwOverall + CT_Correction(0), 0)
    txt63.Value = Round(LwOverall + CT_Correction(1), 0)
    txt125.Value = Round(LwOverall + CT_Correction(2), 0)
    txt250.Value = Round(LwOverall + CT_Correction(3), 0)
    txt500.Value = Round(LwOverall + CT_Correction(4), 0)
    txt1k.Value = Round(LwOverall + CT_Correction(5), 0)
    txt2k.Value = Round(LwOverall + CT_Correction(6), 0)
    txt4k.Value = Round(LwOverall + CT_Correction(7), 0)
    txt8k.Value = Round(LwOverall + CT_Correction(8), 0)
    End If
    
End Sub

Function checkTXT() As Boolean

    If txt31dir.Value = "" Then
    checkTXT = False
    ElseIf txt63dir.Value = "" Then
    checkTXT = False
    ElseIf txt125dir.Value = "" Then
    checkTXT = False
    ElseIf txt250dir.Value = "" Then
    checkTXT = False
    ElseIf txt500dir.Value = "" Then
    checkTXT = False
    ElseIf txt1kdir.Value = "" Then
    checkTXT = False
    ElseIf txt2kdir.Value = "" Then
    checkTXT = False
    ElseIf txt4kdir.Value = "" Then
    checkTXT = False
    ElseIf txt8kdir.Value = "" Then
    checkTXT = False
    Else
    checkTXT = True
    End If
End Function

