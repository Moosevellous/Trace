VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStandardCalc 
   Caption         =   "Insert Standard Calculation Sheet"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7740
   OleObjectBlob   =   "frmStandardCalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStandardCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
IMPORTSHEETNAME = ""
frmStandardCalc.Hide
btnOkPressed = False
End Sub

Private Sub btnLoadStandardCalc_Click()
IMPORTSHEETNAME = cBoxSelectTemplate.Text
frmStandardCalc.Hide
btnOkPressed = True
End Sub

Private Sub UserForm_Activate()
    With frmStandardCalc
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub
