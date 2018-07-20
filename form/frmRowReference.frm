VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRowReference 
   Caption         =   "Row Reference"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "frmRowReference.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRowReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
UserSelectedAddress = refRangeSelector.Value
btnOkPressed = True
Me.Hide
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Me.refRangeSelector.Value = ""
End Sub
