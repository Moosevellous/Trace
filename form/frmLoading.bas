VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoading 
   Caption         =   "Loading..."
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "frmLoading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
End
End Sub

Private Sub UserForm_Initialize()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
