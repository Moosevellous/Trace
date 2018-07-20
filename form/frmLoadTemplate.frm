VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoadTemplate 
   Caption         =   "Insert New Template Sheet"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   OleObjectBlob   =   "frmLoadTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoadTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
IMPORTSHEETNAME = ""
frmLoadTemplate.Hide
End Sub

Private Sub btnLoadTemplate_Click()
IMPORTSHEETNAME = cBoxSelectTemplate.Text
frmLoadTemplate.Hide
End Sub

Private Sub cBoxSelectTemplate_Change()
tBoxDescription.Text = DESCRIPTION(cBoxSelectTemplate.ListIndex + 1)
End Sub

Private Sub UserForm_Activate()
    With frmLoadTemplate
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
'select first item
Me.cBoxSelectTemplate.ListIndex = 0
End Sub
