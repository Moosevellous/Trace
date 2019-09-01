VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVibUnits 
   Caption         =   "Vibration - Convert Units"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "frmVibUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVibUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
Dim PowerOf As Integer

btnOkPressed = True

    If Me.optAccel.Value Then
    PowerOf = -6
    ElseIf Me.optVelocity.Value Then
    PowerOf = -9
    ElseIf Me.optDisplacement.Value Then
    PowerOf = -12
    Else
    msg = MsgBox("Error, no value selected", vbOKOnly, "You must choooooooose")
    End
    End If

    If Me.optMetres.Value Then
    VibRef = "1e" & CStr(PowerOf)
    ElseIf Me.optMillimetres.Value Then
    VibRef = "1e" & CStr(PowerOf + 3)
    Else
    msg = MsgBox("Error, no value selected", vbOKOnly, "You must choooooooose")
    End
    End If


Me.Hide
Unload Me
End Sub

Private Sub optAccel_Click()
Me.optMetres.Caption = "m/s" & chr(178)
Me.optMillimetres.Caption = "mm/s" & chr(178)
End Sub

Private Sub optDisplacement_Click()
Me.optMetres.Caption = "m"
Me.optMillimetres.Caption = "mm"
End Sub

Private Sub optVelocity_Click()
Me.optMetres.Caption = "m/s"
Me.optMillimetres.Caption = "mm/s"
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

