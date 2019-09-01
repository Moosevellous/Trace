VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCouplingLoss 
   Caption         =   "Coupling Loss"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   OleObjectBlob   =   "frmCouplingLoss.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCouplingLoss"
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
BuildingType = Me.cBoxType.Value
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub cBoxType_Change()
CRL = Array(2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2, 2, 2)
LargeMasonryOnPiles = Array(6, 6, 6, 6, 7, 7, 7, 8, 9, 10, 11, 12, 13, 13, 14, 14, 15, 15, 15)
LargeMasonryOnSpreadFootings = Array(11, 11, 11, 11, 12, 13, 14, 14, 15, 15, 15, 15, 14, 14, 14, 14, 13, 12, 11)
TwoToFourStoreyMasonryOnSpreadFootings = Array(5, 6, 6, 7, 9, 11, 11, 12, 13, 13, 13, 13, 13, 12, 12, 11, 10, 9, 8)
OneToTwoStoreyCommercial = Array(4, 5, 5, 6, 7, 8, 8, 9, 9, 9, 9, 9, 9, 8, 8, 8, 7, 6, 5)
SingleResidential = Array(3, 3, 4, 4, 5, 5, 6, 6, 6, 6, 6, 6, 6, 5, 5, 5, 4, 4, 4)
    
    Select Case Me.cBoxType.Text
    Case Is = "CRL"
    SelectedLoss = CRL
    Case Is = "Large Masonry On Piles"
    SelectedLoss = LargeMasonryOnPiles
    Case Is = "Large Masonry on Spread Footings"
    SelectedLoss = LargeMasonryOnSpreadFootings
    Case Is = "2-4 Storey Masonry on Spread Footings"
    SelectedLoss = TwoToFourStoreyMasonryOnSpreadFootings
    Case Is = "1-2 Storey Commercial"
    SelectedLoss = OneToTwoStoreyCommercial
    Case Is = "Single Residential"
    SelectedLoss = SingleResidential
    Case Is = ""
    SelectedLoss = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    End Select
    
Me.txt5.Value = -1 * SelectedLoss(0)
Me.txt6.Value = -1 * SelectedLoss(1)
Me.txt8.Value = -1 * SelectedLoss(2)
Me.txt10.Value = -1 * SelectedLoss(3)
Me.txt12.Value = -1 * SelectedLoss(4)
Me.txt16.Value = -1 * SelectedLoss(5)
Me.txt20.Value = -1 * SelectedLoss(6)
Me.txt25.Value = -1 * SelectedLoss(7)
Me.txt31.Value = -1 * SelectedLoss(8)
Me.txt40.Value = -1 * SelectedLoss(9)
Me.txt50.Value = -1 * SelectedLoss(10)
Me.txt63.Value = -1 * SelectedLoss(11)
Me.txt80.Value = -1 * SelectedLoss(12)
Me.txt100.Value = -1 * SelectedLoss(13)
Me.txt125.Value = -1 * SelectedLoss(14)
Me.txt160.Value = -1 * SelectedLoss(15)
Me.txt200.Value = -1 * SelectedLoss(16)
Me.txt250.Value = -1 * SelectedLoss(17)
Me.txt315.Value = -1 * SelectedLoss(18)

End Sub

Private Sub UserForm_Activate()
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
PopulateComboBox
End Sub

Sub PopulateComboBox()
    With Me.cBoxType
    .AddItem ("CRL")
    .AddItem ("Large Masonry On Piles")
    .AddItem ("Large Masonry on Spread Footings")
    .AddItem ("2-4 Storey Masonry on Spread Footings")
    .AddItem ("1-2 Storey Commercial")
    .AddItem ("Single Residential")
    End With
End Sub


