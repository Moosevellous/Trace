VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRowReference 
   Caption         =   "Row Reference"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   OleObjectBlob   =   "frmRowReference.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRowReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
UserSelectedAddress = refRangeSelector.Value
    If Me.optMultiRow.Value = True Then
    LookupMultiRow = True
    Else 'default to single row
    LookupMultiRow = False
    End If
btnOkPressed = True
Me.Hide
End Sub


Private Sub refRangeSelector_Change()
    If InStr(1, Me.refRangeSelector.Value, ":", vbTextCompare) > 0 Then
    Me.optMultiRow.Value = True
    Else
    Me.optSingleRow.Value = True
    End If
End Sub

Private Sub UserForm_Activate()
    With frmRowReference
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
Me.refRangeSelector.Value = ""
End Sub


