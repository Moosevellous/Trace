VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoomLossClassic 
   Caption         =   "Room Loss (Classic)"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7260
   OleObjectBlob   =   "frmRoomLossClassic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRoomLossClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public V As Double

Private Sub btnHelp_Click()
GotoWikiPage ("Noise-Functions#classic")
End Sub

Private Sub cBoxRoomType_Change()
PreviewValues
End Sub

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
'
'Private Sub optRT_Click()
'Me.cBoxRT.Enabled = True
'Me.lblReverbTime.Enabled = True
'Me.cBoxRoomType.Enabled = False
'Me.lblRoomAbsorption.Enabled = False
'End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
build_cBox
PreviewValues
End Sub

Sub PreviewValues()
Dim V As Double
Dim l As Double
Dim W As Double
Dim H As Double
Dim S_total As Double
Dim alphaValues() As Variant

On Error GoTo errorCatch

'alphas
alphaValues = RoomAlphaDefault(Me.cBoxRoomType.Value)
Me.txt31.Value = alphaValues(0)
Me.txt63.Value = alphaValues(1)
Me.txt125.Value = alphaValues(2)
Me.txt250.Value = alphaValues(3)
Me.txt500.Value = alphaValues(4)
Me.txt1k.Value = alphaValues(5)
Me.txt2k.Value = alphaValues(6)
Me.txt4k.Value = alphaValues(7)
Me.txt8k.Value = alphaValues(8)

'Calc volume & S_total
l = CDbl(Me.txtL.Value)
W = CDbl(Me.txtW.Value)
H = CDbl(Me.txtH.Value)
V = l * W * H
Me.txtV.Value = V
S_total = (l * W * 2) + (l * H * 2) + (W * H * 2)
Me.txtStotal.Value = S_total

'Room Loss
Me.txtSA31.Value = Round(RoomLossTypical("31.5", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA63.Value = Round(RoomLossTypical("63", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA125.Value = Round(RoomLossTypical("125", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA250.Value = Round(RoomLossTypical("250", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA500.Value = Round(RoomLossTypical("500", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA1k.Value = Round(RoomLossTypical("1k", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA2k.Value = Round(RoomLossTypical("2k", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA4k.Value = Round(RoomLossTypical("4k", l, W, H, Me.cBoxRoomType.Value), 1)
Me.txtSA8k.Value = Round(RoomLossTypical("8k", l, W, H, Me.cBoxRoomType.Value), 1)

errorCatch:
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
btnOkPressed = True

'    If Me.optClassic.Value = True Then
'    roomLossType = "Classic" 'public variable
'    ElseIf Me.optRT.Value = True Then
'    roomLossType = "RT"
'    End If

PreviewValues
roomType = Me.cBoxRoomType.Text
roomL = CDbl(Me.txtL.Value)
roomW = CDbl(Me.txtW.Value)
roomH = CDbl(Me.txtH.Value)
Me.Hide
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


Sub PrePopulateForm() 'InputVol As Long, InputRoomType As String)
Me.txtL.Text = roomL
Me.txtW.Text = roomW
Me.txtH.Text = roomH
Me.cBoxRoomType.Text = roomType
End Sub
