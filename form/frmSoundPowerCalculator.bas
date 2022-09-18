VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSoundPowerCalculator 
   Caption         =   "Sound Power Calculator"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "frmSoundPowerCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSoundPowerCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage
End Sub

Private Sub btnOK_Click()
'Set public variables
    If IsNumeric(Me.txtL.Value) And IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) And IsNumeric(Me.txtOffset.Value) Then
    roomL = Me.txtL.Value
    roomW = Me.txtW.Value
    roomH = Me.txtH.Value
    OffsetDistance = Me.txtOffset.Value
    btnOkPressed = True
    End If
Unload Me
End Sub


Private Sub optConformal_Click()
lblStotal.Caption = "Conformal surface area="
PreviewValues
End Sub

Private Sub optParallel_Click()
lblStotal.Caption = "Parallel box surface area="
PreviewValues
End Sub

Private Sub txtH_Change()
PreviewValues
End Sub

Private Sub txtL_Change()
PreviewValues
End Sub

Private Sub txtOffset_Change()
    If IsNumeric(Me.txtOffset.Value) Then
        If Me.txtOffset.Value < 1 Then
        Me.lblWarning.Visible = True
        Else
        Me.lblWarning.Visible = False
        End If
    End If

PreviewValues
End Sub

Private Sub txtW_Change()
PreviewValues
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
PreviewValues
End Sub

Sub PreviewValues()
    If IsNumeric(Me.txtL.Value) And IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) And IsNumeric(Me.txtOffset.Value) Then
        If Me.optConformal.Value = True Then
        Me.txtStotal = Round(ConformalSurfaceArea(Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtOffset.Value), 1)
        Else
        Me.txtStotal = Round(ParallelipipedSurfaceArea(Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtOffset.Value), 1)
        End If
    End If
End Sub
