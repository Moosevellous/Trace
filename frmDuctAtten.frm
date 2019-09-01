VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuctAtten 
   Caption         =   "Duct Attenuation"
   ClientHeight    =   6840
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

Private Function getDuctShape()

    If Me.opt25mm.Value Then
    w = 25
    ElseIf Me.opt50mm.Value Then
    w = 50
    ElseIf Me.optUnlined.Value Then
    w = 0
    End If
    
    If Me.optCir.Value Then
    s = "C"
    ElseIf Me.optRect.Value Then
    s = "R"
    End If

getDuctShape = CStr(w) & " " & s

End Function

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
SwitchCustomThickness
PreviewInsertionLoss
End Sub

Private Sub optCir_Click()
Me.txtW.Enabled = False
End Sub

Private Sub optCustom_Click()
Me.txtThickness.Enabled = True
End Sub

Private Sub optRect_Click()
Me.txtW.Enabled = True
End Sub

Private Sub optReynolds_Click()
SwitchCustomThickness
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

Private Sub txtW_Change()
PreviewInsertionLoss
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewInsertionLoss
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
frmDuctAtten.Hide
End Sub

Private Sub btnOK_Click()

    If Me.optASHRAE.Value Then
    ductMethod = "ASHRAE"
    ElseIf Me.optReynolds Then
    ductMethod = "Reynolds"
    Else
    ductMethod = ""
    End If

ductH = CSng(Me.txtH.Value)
ductW = CSng(Me.txtW.Value)
ductL = CSng(Me.txtL.Value)
ductShape = getDuctShape
ductLiningThickness = CSng(Me.txtThickness.Value)
btnOkPressed = True
frmDuctAtten.Hide
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
    thicknessParam = CDbl(Me.txtThickness.Value)
    End If
    
    'calculation type
    If Me.optASHRAE.Value = True Then
    Me.txt63.Value = GetASHRAE(63, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt125.Value = GetASHRAE(125, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt250.Value = GetASHRAE(250, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt500.Value = GetASHRAE(500, CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt1k.Value = GetASHRAE("1k", CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt2k.Value = GetASHRAE("2k", CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt4k.Value = GetASHRAE("4k", CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    Me.txt8k.Value = GetASHRAE("8k", CLng(Me.txtH.Value), CLng(Me.txtW.Value), ductParam, CLng(Me.txtL.Value))
    ElseIf Me.optReynolds.Value = True Then
    Me.txt63.Value = GetReynoldsDuct(63, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt125.Value = GetReynoldsDuct(125, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt250.Value = GetReynoldsDuct(250, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt500.Value = GetReynoldsDuct(500, CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt1k.Value = GetReynoldsDuct("1k", CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt2k.Value = GetReynoldsDuct("2k", CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt4k.Value = GetReynoldsDuct("4k", CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Me.txt8k.Value = GetReynoldsDuct("8k", CDbl(Me.txtH.Value), CDbl(Me.txtW.Value), thicknessParam, CDbl(Me.txtL.Value))
    Else
    'nothing
    End If
    

End Sub

Sub SwitchCustomThickness()
    If Me.optReynolds = True Then
    Me.optCustom.Enabled = True
    Me.txtThickness.Enabled = True
    Else 'ASHRAE enabled
        'default to nearest thickness
        If Me.optCustom.Value = True Then
        Me.opt25mm.Value = True
        End If
    Me.optCustom.Enabled = False
    Me.txtThickness.Enabled = False
    End If
End Sub

Sub CheckDuctSize()
Dim MaxArea As Double
Dim DuctArea As Double
MaxArea = 3.66 * 1.02   'max dimension in metres from ASHRAE Tables
    If Me.optASHRAE.Value = True Then
    DuctArea = Me.txtW.Value * Me.txtL.Value / 100000
        If DuctArea > MaxArea Then
        msg = MsgBox("Warning, ASHRAE tables only go up to a total duct area of 3.7332m" & chr(178), vbOKOnly, "Error - ASHRAE duct size")
        End If
    End If
End Sub
