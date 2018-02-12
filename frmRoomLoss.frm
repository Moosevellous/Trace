VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoomLoss 
   Caption         =   "Room Loss"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4635
   OleObjectBlob   =   "frmRoomLoss.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRoomLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public V As Double

'Private Sub optClassic_Click()
'Me.cBoxRT.Enabled = False
'Me.lblReverbTime.Enabled = False
'Me.cBoxRoomType.Enabled = True
'Me.lblRoomAbsorption.Enabled = True
'End Sub

'Private Sub optRoomConstant_Click()
'Me.cBoxRT.Enabled = False
'Me.lblReverbTime.Enabled = False
'Me.cBoxRoomType.Enabled = False
'Me.lblRoomAbsorption.Enabled = False
'End Sub

Private Sub optRT_Click()
Me.cBoxRT.Enabled = True
Me.lblReverbTime.Enabled = True
Me.cBoxRoomType.Enabled = False
Me.lblRoomAbsorption.Enabled = False
End Sub

Private Sub UserForm_Activate()
    With Me
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
build_cBox
CalcVolume
End Sub

Sub CalcVolume()
On Error GoTo errorcatch
V = CDbl(Me.txtL.Value) * CDbl(Me.txtW.Value) * CDbl(Me.txtH.Value)
Me.txtV.Value = V
errorcatch:
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
btnOkPressed = True

'    If Me.optClassic.Value = True Then
'    roomLossType = "Classic" 'global variable
'    ElseIf Me.optRT.Value = True Then
'    roomLossType = "RT"
'    End If

CalcVolume
roomType = Me.cBoxRoomType.Text
roomL = CDbl(Me.txtL.Value)
roomW = CDbl(Me.txtW.Value)
roomH = CDbl(Me.txtH.Value)
Me.Hide
End Sub

Private Sub txtH_Change()
CalcVolume
End Sub

Private Sub txtL_Change()
CalcVolume
End Sub

Private Sub txtW_Change()
CalcVolume
End Sub

Private Sub UserForm_Click()

End Sub

Sub build_cBox()
    If Me.cBoxRoomType.ListCount = 0 Then
    Me.cBoxRoomType.AddItem ("Dead")
    Me.cBoxRoomType.AddItem ("Av. Dead")
    Me.cBoxRoomType.AddItem ("Average")
    Me.cBoxRoomType.AddItem ("Av. Live")
    Me.cBoxRoomType.AddItem ("Live")
    End If
    
'    If Me.cBoxRT.ListCount = 0 Then
'    Me.cBoxRT.AddItem ("<0.2 sec")
'    Me.cBoxRT.AddItem ("0.2 to 0.5 sec")
'    Me.cBoxRT.AddItem ("0.5 to 1 sec")
'    Me.cBoxRT.AddItem ("1 to 1.5 sec")
'    Me.cBoxRT.AddItem ("1.5 to 2 sec")
'    Me.cBoxRT.AddItem (">2 sec")
'    End If
End Sub


Sub Populate_frmRoomLoss() 'InputVol As Long, InputRoomType As String)
Me.txtL.Text = roomL
Me.txtW.Text = roomW
Me.txtH.Text = roomH
Me.cBoxRoomType.Text = roomType
End Sub
