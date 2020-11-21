VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbowBend 
   Caption         =   "Elbow/Bend Loss"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   OleObjectBlob   =   "frmElbowBend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElbowBend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#elbow--bend")
End Sub

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
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewElbowLoss
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
'set public variables
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
    ElbowVanes = "Vanes"
    Else
    ElbowVanes = "No Vanes"
    End If
    
    If Me.chkCalcRegen.Value = True Then
    CalcRegen = True
    Else
    CalcRegen = False
    End If
    
btnOkPressed = True
Me.Hide
Unload Me
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
    Me.txt31.Value = ElbowLoss_ASHRAE("31.5", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt63.Value = ElbowLoss_ASHRAE("63", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt125.Value = ElbowLoss_ASHRAE("125", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt250.Value = ElbowLoss_ASHRAE("250", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt500.Value = ElbowLoss_ASHRAE("500", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt1k.Value = ElbowLoss_ASHRAE("1k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt2k.Value = ElbowLoss_ASHRAE("2k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt4k.Value = ElbowLoss_ASHRAE("4k", Me.txtW.Value, Shape, Lining, Vanes)
    Me.txt8k.Value = ElbowLoss_ASHRAE("8k", Me.txtW.Value, Shape, Lining, Vanes)
    End If
    
End Sub

