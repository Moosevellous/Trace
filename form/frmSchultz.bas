VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSchultz 
   Caption         =   "Room Correction - Schultz method"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   OleObjectBlob   =   "frmSchultz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSchultz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
'set public variables
roomL = Me.txtL.Value
roomW = Me.txtW.Value
roomH = Me.txtH.Value
RoomVolume = Me.txtV.Value
DistanceFromSource = Me.txtD.Value

btnOkPressed = True
Me.Hide
End Sub

Private Sub txtD_Change()
PreviewValues
End Sub

Private Sub txtH_Change()
PreviewValues
End Sub

Private Sub txtL_Change()
PreviewValues
End Sub

Private Sub txtW_Change()
PreviewValues
End Sub

Sub PreviewValues()

Dim R_L#, R_W#, R_H#, D#, Volume# '# means double
Dim R_Vol As Double

R_L = CheckNumericValue(Me.txtL.Value)
R_W = CheckNumericValue(Me.txtW.Value)
R_H = CheckNumericValue(Me.txtH.Value)
D = CheckNumericValue(Me.txtD.Value)

'volume & log volume
R_Vol = R_L * R_W * R_H
Me.txtV.Value = Round(R_Vol, 1)

    If R_Vol > 0 Then
    Me.txtLogV = CheckNumericValue(5 * Application.WorksheetFunction.Log10(R_Vol), 1)
    'values
    Me.txtSA31.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "31.5"), 1)
    Me.txtSA63.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "63"), 1)
    Me.txtSA125.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "125"), 1)
    Me.txtSA250.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "250"), 1)
    Me.txtSA500.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "500"), 1)
    Me.txtSA1k.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "1k"), 1)
    Me.txtSA2k.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "2k"), 1)
    Me.txtSA4k.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "4k"), 1)
    Me.txtSA8k.Value = Round(RoomCorrection_Schultz(R_L, R_W, R_H, D, "8k"), 1)
    Else
    Me.txtLogV.Value = "-"
    Me.txtSA31.Value = "-"
    Me.txtSA63.Value = "-"
    Me.txtSA125.Value = "-"
    Me.txtSA250.Value = "-"
    Me.txtSA500.Value = "-"
    Me.txtSA1k.Value = "-"
    Me.txtSA2k.Value = "-"
    Me.txtSA4k.Value = "-"
    Me.txtSA8k.Value = "-"
    End If
    
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewValues
End Sub
