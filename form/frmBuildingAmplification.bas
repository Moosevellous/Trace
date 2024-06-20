VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBuildingAmplification 
   Caption         =   "Building Amplification"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11115
   OleObjectBlob   =   "frmBuildingAmplification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBuildingAmplification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Vibration#building-vibration")
End Sub

Private Sub btnOK_Click()
AmplificationType = Me.cBoxType.Value
btnOkPressed = True
Unload Me
End Sub


Private Sub cBoxType_Change()
FloorVibration = Array(10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 10, 9, 9, 0, 0, 0, 0, 0, 0)
GBN = Array(0, 0, 0, 0, 0, 0, 6, 7, 7, 7, 6, 6, 5, 5, 4, 3, 2, 1, 1)
    
    Select Case Me.cBoxType.text
    Case Is = "Ground-borne Noise"
    SelectedLoss = GBN
    Case Is = "Floor Vibration"
    SelectedLoss = FloorVibration
    Case Is = ""
    SelectedLoss = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    End Select
    
Me.txt5.Value = SelectedLoss(0)
Me.txt6.Value = SelectedLoss(1)
Me.txt8.Value = SelectedLoss(2)
Me.txt10.Value = SelectedLoss(3)
Me.txt12.Value = SelectedLoss(4)
Me.txt16.Value = SelectedLoss(5)
Me.txt20.Value = SelectedLoss(6)
Me.txt25.Value = SelectedLoss(7)
Me.txt31.Value = SelectedLoss(8)
Me.txt40.Value = SelectedLoss(9)
Me.txt50.Value = SelectedLoss(10)
Me.txt63.Value = SelectedLoss(11)
Me.txt80.Value = SelectedLoss(12)
Me.txt100.Value = SelectedLoss(13)
Me.txt125.Value = SelectedLoss(14)
Me.txt160.Value = SelectedLoss(15)
Me.txt200.Value = SelectedLoss(16)
Me.txt250.Value = SelectedLoss(17)
Me.txt315.Value = SelectedLoss(18)

End Sub

Private Sub UserForm_Initialize()
    With Me
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

PopulateComboBox

End Sub

Sub PopulateComboBox()
    With Me.cBoxType
    .AddItem ("Ground-borne Noise")
    .AddItem ("Floor Vibration")
    End With
End Sub

