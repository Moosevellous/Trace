VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBasicFunctions 
   Caption         =   "Basic Functions"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   OleObjectBlob   =   "frmBasicFunctions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBasicFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Basics")
End Sub

Private Sub btnOK_Click()
    If Me.optSPLSUM.Value = True Then
    BasicFunctionType = "SPLSUM"
    ElseIf Me.optSPLAV.Value = True Then
    BasicFunctionType = "SPLAV"
    ElseIf Me.optSPLMINUS.Value = True Then
    BasicFunctionType = "SPLMINUS"
    ElseIf Me.optSPLSUMIF.Value = True Then
    BasicFunctionType = "SPLSUMIF"
    ElseIf Me.optSPLAVIF.Value = True Then
    BasicFunctionType = "SPLAVIF"
    ElseIf Me.optSum.Value = True Then
    BasicFunctionType = "SUM"
    ElseIf Me.optAverage.Value = True Then
    BasicFunctionType = "AVERAGE"
    End If

RangeSelection = Me.refRangeSelector.Value

    If Me.refRange2Selector.Enabled = True Then
    Range2Selection = Me.refRange2Selector.Value
    Else
    Range2Selection = Null
    End If
    
    If Me.chkApplyToSheetType = True Then
    ApplyToSheetType = True
    Else
    ApplyToSheetType = False
    End If
    
BasicsApplyStyle = Me.cBoxApplyStyle.Value 'style-ish

btnOkPressed = True
Me.Hide
End Sub

Private Sub optAverage_Click()
Me.lblExampleFormula.Caption = "=AVERAGE(Range)"
hideRange2
End Sub

Private Sub optSPLAV_Click()
Me.lblExampleFormula.Caption = "=SPLAV(Range)"
hideRange2
End Sub

Private Sub optSPLAVIF_Click()
Me.lblExampleFormula.Caption = "=SPLAVIF(Range,Condition)"
showRange2
End Sub

Private Sub optSPLMINUS_Click()
Me.lblExampleFormula.Caption = "=SPLMINUS(SPLtotal,SPL2)"
showRange2
End Sub

Private Sub optSPLSUM_Click()
Me.lblExampleFormula.Caption = "=SPLSUM(Range)"
hideRange2
End Sub

Private Sub optSPLSUMIF_Click()
Me.lblExampleFormula.Caption = "=SPLSUMIF(Range,Condition)"
showRange2
End Sub

Sub showRange2()
Me.refRange2Selector.Visible = True
Me.lblSelectRange2.Visible = True
End Sub
Sub hideRange2()
Me.refRange2Selector.Visible = False
Me.lblSelectRange2.Visible = False
End Sub

Private Sub optSum_Click()
Me.lblExampleFormula.Caption = "=SUM(Range)"
hideRange2
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
Me.cBoxApplyStyle.AddItem ("None")
Me.cBoxApplyStyle.AddItem ("Normal")
Me.cBoxApplyStyle.AddItem ("Lw Source")
Me.cBoxApplyStyle.AddItem ("Subtotal")
Me.cBoxApplyStyle.AddItem ("Total")
End Sub
