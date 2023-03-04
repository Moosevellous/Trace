VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuctAtten 
   Caption         =   "Duct Attenuation"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6180
   OleObjectBlob   =   "frmDuctAtten.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuctAtten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function InputsOK() As Boolean
InputsOK = True
If Me.txtW.Value = "" Then InputsOK = False
If Me.txtH.Value = "" Then InputsOK = False
If Me.txtThickness.Value = "" Then InputsOK = False
If Me.txtL.Value = "" Then InputsOK = False
End Function

Private Function getDuctShape()

    If Me.opt25mm.Value Then
    W = 25
    ElseIf Me.opt50mm.Value Then
    W = 50
    ElseIf Me.optUnlined.Value Then
    W = 0
    End If
    
    If Me.optCir.Value Then
    s = "C"
    ElseIf Me.optRect.Value Then
    s = "R"
    End If

getDuctShape = CStr(W) & " " & s

End Function

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#solid-duct")
End Sub

Private Sub lblUnlinedOnly_Click()
Me.optSRL.Value = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub opt25mm_Click()
Me.txtThickness.Enabled = False
Me.txtThickness.Value = 25
PreviewInsertionLoss
End Sub

Private Sub opt50mm_Click()
Me.txtThickness.Enabled = False
Me.txtThickness.Value = 50
PreviewInsertionLoss
End Sub

Private Sub optASHRAE_Click()
EnableButtons
PreviewInsertionLoss
End Sub

Private Sub optCir_Click()
EnableButtons
PreviewInsertionLoss
End Sub

Private Sub optCustom_Click()
Me.txtThickness.Enabled = True
End Sub

Private Sub optRect_Click()
EnableButtons
PreviewInsertionLoss
End Sub

Private Sub optReynolds_Click()
EnableButtons
PreviewInsertionLoss
End Sub

Private Sub optSRL_Click()
EnableButtons
PreviewInsertionLoss
End Sub

Private Sub optUnlined_Click()
Me.txtThickness.Enabled = False
Me.txtThickness.Value = 0
PreviewInsertionLoss
End Sub

Private Sub txtH_Change()
CheckDuctSize
PreviewInsertionLoss
End Sub

Private Sub txtL_Change()
CheckDuctSize
PreviewInsertionLoss
End Sub

Private Sub txtThickness_Change()
PreviewInsertionLoss
End Sub

Private Sub txtW_Change()
PreviewInsertionLoss
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewInsertionLoss
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()

    If Me.optASHRAE.Value Then
    ductMethod = "ASHRAE"
    ElseIf Me.optReynolds Then
    ductMethod = "Reynolds"
    ElseIf Me.optSRL.Value = True Then
    ductMethod = "SRL"
    Else
    ductMethod = ""
    End If
    
'<---TODO check isnumeric on controls
ductH = CSng(Me.txtH.Value)
ductW = CSng(Me.txtW.Value)
ductL = CSng(Me.txtL.Value)
ductShape = getDuctShape
ductLiningThickness = CSng(Me.txtThickness.Value)
btnOkPressed = True
Me.Hide
End Sub

Sub PreviewInsertionLoss()
Dim ductParam As String
Dim thicknessParam As Double
    
    'lining thickness
    If Me.opt25mm.Value = True Then
    ductParam = "25"
    ElseIf Me.opt50mm.Value = True Then
    ductParam = "50"
    Else
    ductParam = "0"
    End If
    
    'duct dimension
    If Me.optRect.Value = True Then
    ductParam = ductParam & " R"
    Else
    ductParam = ductParam & " C"
    End If
    
    'lining thickness (for reynolds only)
    If Me.opt25mm.Value = True Then
    thicknessParam = 25
    ElseIf Me.opt50mm.Value = True Then
    thicknessParam = 50
    Else
        If IsNumeric(Me.txtThickness.Value) Then
        thicknessParam = CDbl(Me.txtThickness.Value)
        Else
        thicknessParam = 0
        End If
    End If
    
    'calculation type
    If IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) And IsNumeric(Me.txtL.Value) Then 'all values ok!
    
        If Me.optASHRAE.Value = True Then
        Me.txt63.Value = DuctAtten_ASHRAE(63, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt125.Value = DuctAtten_ASHRAE(125, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt250.Value = DuctAtten_ASHRAE(250, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt500.Value = DuctAtten_ASHRAE(500, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt1k.Value = DuctAtten_ASHRAE(1000, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt2k.Value = DuctAtten_ASHRAE(2000, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt4k.Value = DuctAtten_ASHRAE(4000, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        Me.txt8k.Value = DuctAtten_ASHRAE(8000, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
        
        Me.txtRawVal.Value = TXT_RAW 'set from public variable
        Me.txtTableHead.Value = TXT_HEAD 'set from public variable
        
        ElseIf Me.optReynolds.Value = True Then
        
            If Me.optRect.Value = True Then 'RECTANGULAR METHOD - REYNOLDS
            Me.txt63.Value = DuctAtten_Reynolds(63, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt125.Value = DuctAtten_Reynolds(125, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt250.Value = DuctAtten_Reynolds(250, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt500.Value = DuctAtten_Reynolds(500, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt1k.Value = DuctAtten_Reynolds(1000, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt2k.Value = DuctAtten_Reynolds(2000, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt4k.Value = DuctAtten_Reynolds(4000, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            Me.txt8k.Value = DuctAtten_Reynolds(8000, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
            ElseIf Me.optCir.Value = True Then 'CIRCULAR METHOD - REYNOLDS
            Me.txt63.Value = CheckNumericValue(DuctAttenCircular_Reynolds(63, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt125.Value = CheckNumericValue(DuctAttenCircular_Reynolds(125, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt250.Value = CheckNumericValue(DuctAttenCircular_Reynolds(250, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt500.Value = CheckNumericValue(DuctAttenCircular_Reynolds(500, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt1k.Value = CheckNumericValue(DuctAttenCircular_Reynolds(1000, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt2k.Value = CheckNumericValue(DuctAttenCircular_Reynolds(2000, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt4k.Value = CheckNumericValue(DuctAttenCircular_Reynolds(4000, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            Me.txt8k.Value = CheckNumericValue(DuctAttenCircular_Reynolds(8000, CDbl(Me.txtH.Value), thicknessParam, CDbl(Me.txtL.Value)), 1)
            End If
            
        ElseIf Me.optSRL.Value = True Then
        Me.txt63.Value = DuctBendAtten_SRL(63, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt125.Value = DuctBendAtten_SRL(125, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt250.Value = DuctBendAtten_SRL(250, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt500.Value = DuctBendAtten_SRL(500, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt1k.Value = DuctBendAtten_SRL(1000, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt2k.Value = DuctBendAtten_SRL(2000, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt4k.Value = DuctBendAtten_SRL(4000, CLng(Me.txtW.Value), Right(ductParam, 1), CLng(Me.txtL.Value))
        Me.txt8k.Value = "-"
        
        Me.txtRawVal.Value = TXT_RAW 'set from public variable
        Me.txtTableHead.Value = TXT_HEAD 'set from public variable
        
        Else
        'no method? do nothing
        End If
    Else
    Me.txt63.Value = "-"
    Me.txt125.Value = "-"
    Me.txt250.Value = "-"
    Me.txt500.Value = "-"
    Me.txt1k.Value = "-"
    Me.txt2k.Value = "-"
    Me.txt4k.Value = "-"
    Me.txt8k.Value = "-"
    Me.txtRawVal.Value = "" 'nothing!
    Me.txtTableHead.Value = ""
    End If 'close check for input values

End Sub

Sub EnableButtons()
    'options for rectangular/circular
    If Me.optRect.Value = True Then
    Me.txtW.Enabled = True
    Me.txtH.Enabled = True
    Me.lblDimensions.Caption = "Dimensions (H x W)"
    Else 'circular
    Me.txtW.Enabled = False
    Me.txtH.Enabled = True
    Me.lblDimensions.Caption = "Dimensions (diameter)"
    End If

    'reynolds
    If Me.optReynolds = True Then
    Me.optCustom.Enabled = True
    Me.txtThickness.Enabled = True
    Me.opt25mm.Enabled = True
    Me.opt50mm.Enabled = True
    Me.txtRawVal.Visible = False
    Me.txtTableHead.Visible = False
    Me.lblTableValues.Visible = False
    'SRL
    ElseIf Me.optSRL.Value = True Then
    Me.optCustom.Enabled = False
    Me.txtThickness.Enabled = False
    Me.optUnlined.Value = True
    Me.opt25mm.Enabled = False
    Me.opt50mm.Enabled = False
    Me.txtRawVal.Visible = True
    Me.txtTableHead.Visible = True
    Me.lblTableValues.Visible = True
    'ASHRAE
    Else
        If Me.optCustom.Value = True Then
        Me.opt25mm.Value = True
        End If
    Me.optCustom.Enabled = False
    Me.txtThickness.Enabled = False
    Me.opt25mm.Enabled = True
    Me.opt50mm.Enabled = True
    Me.txtRawVal.Visible = True
    Me.txtTableHead.Visible = True
    Me.lblTableValues.Visible = True
    End If
End Sub

Sub CheckDuctSize()
Dim MaxArea As Double
Dim DuctArea As Double
MaxArea = 3.66 * 1.02   'max dimension in metres from ASHRAE Tables
    If Me.optASHRAE.Value = True And IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
    DuctArea = Me.txtW.Value * Me.txtH.Value / 100000
        If DuctArea > MaxArea Then
        msg = MsgBox("Warning, ASHRAE tables only go up to a total duct area of 3.7332m" & chr(178), vbOKOnly, "Error - ASHRAE duct size")
        End If
    End If
End Sub
