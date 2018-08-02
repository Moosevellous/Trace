VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClickAddress 
   Caption         =   "Add / Average"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4290
   OleObjectBlob   =   "frmClickAddress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClickAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    With frmClickAddress
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
frmClickAddress.Hide
End Sub

Private Sub btnOK_Click()
UserSelectedAddress = refRangeSelector.Value

    If frmClickAddress.optSum = True Then
    SumOrAverage = "Sum"
    ElseIf frmClickAddress.optAverage = True Then
    SumOrAverage = "Average"
    Else
    msg = MsgBox("What did you press? Option does not exist.", vbOKOnly, "ERROR!")
    End If

btnOkPressed = True
frmClickAddress.Hide
End Sub
