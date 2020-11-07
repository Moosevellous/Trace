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
roomL = Me.txtL.Value
roomW = Me.txtW.Value
roomH = Me.txtH.Value
OffsetDistance = Me.txtOffset.Value
btnOkPressed = True
Unload Me
End Sub



Private Sub txtH_Change()
PreviewValues
End Sub

Private Sub txtL_Change()
PreviewValues
End Sub

Private Sub txtOffset_Change()
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
    If IsNumeric(Me.txtL.Value) And IsNumeric(Me.txtW.Value) And IsNumeric(Me.txtH.Value) Then
    Me.txtStotal = ParallelipipedSurfaceArea(Me.txtL.Value, Me.txtW.Value, Me.txtH.Value, Me.txtOffset.Value)
    End If
End Sub
