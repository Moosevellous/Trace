VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegenNoise 
   Caption         =   "Regenerated Noise"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmRegenNoise.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegenNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    With frmRegenNoise
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
frmRegenNoise.Hide
End Sub

Private Sub btnOK_Click()

    If Me.optElbow.Value Then
    regenNoiseElement = "Elbow"
    ElseIf Me.optDamper.Value Then
    regenNoiseElement = "Damper"
    ElseIf Me.optTransition.Value Then
    regenNoiseElement = "Transition"
    End If

btnOkPressed = True
frmRegenNoise.Hide
End Sub
