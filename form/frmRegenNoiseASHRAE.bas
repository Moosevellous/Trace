VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegenNoiseASHRAE 
   Caption         =   "Regenerated Noise (ASHRAE)"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmRegenNoiseASHRAE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegenNoiseASHRAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
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
Me.Hide
End Sub
