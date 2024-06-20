VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRowReference 
   Caption         =   "Row Reference"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
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
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Row-Functions#row-reference")
End Sub

Private Sub btnOK_Click()
'set public variables
UserSelectedAddress = Me.refRangeSelector.Value
UserDestinationAddress = Me.refDestinationSelector.Value
LookupMultiRow = Me.optMultiRow.Value = True
DynamicReferencing = Me.chkDynamicRef.Value
AddSchedMarker = Me.chkSchedMarker.Value
RegenDestinationRange = Me.optRegenSWL.Value 'true if true!
btnOkPressed = True
Me.Hide
Unload Me
End Sub


Private Sub optMultiRow_Click()
Me.chkDynamicRef.Value = False
Me.chkDynamicRef.Enabled = False
End Sub

Private Sub optSingleRow_Click()
Me.chkDynamicRef.Enabled = True
End Sub

Private Sub refRangeSelector_Change()
    If InStr(1, Me.refRangeSelector.Value, ":", vbTextCompare) > 0 Then
    Me.optMultiRow.Value = True
    Else
    Me.optSingleRow.Value = True
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With frmRowReference
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_Initialize()
Me.refRangeSelector.Value = ""
End Sub


