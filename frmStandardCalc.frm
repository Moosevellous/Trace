VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStandardCalc 
   Caption         =   "Load/Import Sheets"
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
ImportSheetName = ""
frmStandardCalc.Hide
btnOkPressed = False
Unload Me
End Sub

Private Sub btnLoadStandardCalc_Click()
ImportSheetName = cBoxSelectTemplate.Text
frmStandardCalc.Hide
btnOkPressed = True
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
