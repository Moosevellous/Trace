VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbows 
   Caption         =   "Elbow Loss"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   OleObjectBlob   =   "frmElbows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElbows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optRadius_Click()
Me.optLined.Enabled = False
Me.optUnlined.Enabled = False
Me.optVanes.Enabled = False
Me.optNoVanes.Enabled = False
End Sub

Private Sub optSquare_Click()
Me.optLined.Enabled = True
Me.optUnlined.Enabled = True
Me.optVanes.Enabled = True
Me.optNoVanes.Enabled = True
End Sub

Private Sub UserForm_Activate()
    With frmElbows
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
frmElbows.Hide
End Sub

Private Sub btnOK_Click()

ductW = txtW.Value

    If Me.optLined.Value Then
    elbowLining = "Lined"
    Else
    elbowLining = "Unlined"
    End If

    If Me.optSquare.Value Then
    elbowShape = "Square"
    Else
    elbowShape = "Radius"
    End If
    
    If Me.optVanes.Value Then
    elbowVanes = "Vanes"
    Else
    elbowVanes = "No Vanes"
    End If
    
btnOkPressed = True
frmElbows.Hide
End Sub

