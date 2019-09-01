VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbows 
   Caption         =   "Elbow Loss"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   OleObjectBlob   =   "frmElbows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElbows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub optLined_Click()
PreviewElbowLoss
End Sub

Private Sub optNoVanes_Click()
PreviewElbowLoss
End Sub

Private Sub optRadius_Click()
Me.optLined.Enabled = False
Me.optUnlined.Enabled = False
Me.optVanes.Enabled = False
Me.optNoVanes.Enabled = False
PreviewElbowLoss
End Sub

Private Sub optSquare_Click()
Me.optLined.Enabled = True
Me.optUnlined.Enabled = True
Me.optVanes.Enabled = True
Me.optNoVanes.Enabled = True
PreviewElbowLoss
End Sub

Private Sub optUnlined_Click()
PreviewElbowLoss
End Sub

Private Sub optVanes_Click()
PreviewElbowLoss
End Sub

Private Sub txtW_Change()
PreviewElbowLoss
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewElbowLoss
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

Sub PreviewElbowLoss()
Dim Lining As String
Dim Vanes As String
Dim Shape As String

    If Me.optLined.Value Then
    Lining = "Lined"
    Else
    Lining = "Unlined"
    End If

    If Me.optSquare.Value Then
    Shape = "Square"
    Else
    Shape = "Radius"
    End If
    
    If Me.optVanes.Value Then
    Vanes = "Vanes"
    Else
    Vanes = "No Vanes"
    End If

    If Me.txtW.Value <> "" And IsNumeric(Me.txtW.Value) Then
    Me.txt31.Value = GetElbowLoss("31.5", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt63.Value = GetElbowLoss("63", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt125.Value = GetElbowLoss("125", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt250.Value = GetElbowLoss("250", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt500.Value = GetElbowLoss("500", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt1k.Value = GetElbowLoss("1k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt2k.Value = GetElbowLoss("2k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt4k.Value = GetElbowLoss("4k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt8k.Value = GetElbowLoss("8k", Me.txtW.Value, Shape, Lining, Vanes)
    End If
    
End Sub

