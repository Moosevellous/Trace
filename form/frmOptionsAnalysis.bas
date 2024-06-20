VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsAnalysis 
   Caption         =   "Options Analysis"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   OleObjectBlob   =   "frmOptionsAnalysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptionsAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnCheckOptions_Click()
CountOptions
End Sub

Private Sub btnHelp_Click()
GotoWikiPage "Sheet-Functions#options-analysis"
End Sub

Private Sub btnOK_Click()
'set public variables
CountOptions
RngVar1 = Me.RefVar1Rng.Value
RngVar2 = Me.RefVar2Rng.Value
TargetRng = Me.RefTargetRng.Value
ResultRng = Me.RefResult.Value
ApplyHeatMap = Me.chkApplyHeatmap.Value
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub lblSourceRange_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.lblHint.Caption = Me.lblSourceRange.ControlTipText
End Sub

'Private Sub RefVar1Rng_Change()
'CountOptions
'End Sub
'
'Private Sub RefVar2Rng_Change()
'CountOptions
'End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Sub CountOptions()
Dim Var1Row As Integer
Dim Var2Row As Integer
Dim Var1Selector As String
Dim Var2Selector As String
Dim nVar1 As Integer
Dim nVar2 As Integer

Var1Row = ExtractAddressElement(Me.RefVar1Rng.Value, 2)
Var2Row = ExtractAddressElement(Me.RefVar2Rng.Value, 2)
Var1Selector = Cells(CInt(Var1Row), T_Description).Address
Var2Selector = Cells(CInt(Var2Row), T_Description).Address

    'count first range
    If HasDataValidation(Range(Var1Selector)) Then
    RngVar1 = Range(Var1Selector).Validation.Formula1
    nVar1 = Range(RngVar1).Count
    Else 'just the one option
    nVar1 = 1
    End If
    
    'count second range
    If HasDataValidation(Range(Var2Selector)) Then
    RngVar2 = Range(Var2Selector).Validation.Formula1
    nVar2 = Range(RngVar2).Count
    Else 'just the one option
    nVar2 = 1
    End If

Me.lblCount.Visible = True
Me.lblVar1Count.Caption = nVar1
Me.lblVar2Count.Caption = nVar2
Me.lblVar1Count.Visible = True
Me.lblVar2Count.Visible = True

'multiply!
Me.lblNumResults.Caption = nVar1 * nVar2
Me.lblNumResults.Visible = True
End Sub

