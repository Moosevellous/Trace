VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Export Source Code"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7995
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
End
End Sub


Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.9 * Application.Width) - (0.9 * .Width)
    .Top = Application.Top + 25 '+ (0.9 * Application.Height) - (0.9 * .Height)
    End With
End Sub
