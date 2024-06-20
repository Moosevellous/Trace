VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFrequencyBandCutoff 
   Caption         =   "Frequency Band Cutoff - ANSI S1.11"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "frmFrequencyBandCutoff.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFrequencyBandCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Basics#frequency-band-cutoff")
End Sub

Private Sub btnOK_Click()
    If Me.optUpper.Value = True Then
    FBC_mode = "upper"
    ElseIf Me.optLower.Value = True Then
    FBC_mode = "lower"
    Else
    ErrorWithInputs
    End If

    If Me.optBand1.Value = True Then
    FBC_bandwidth = 1
    ElseIf Me.optBand3.Value = True Then
    FBC_bandwidth = 3
    Else
    ErrorWithInputs
    End If
    
FBC_baseTen = Me.opttBaseTen.Value 'boolean value, oh yeah!
    
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Sub ErrorWithInputs()

End Sub
