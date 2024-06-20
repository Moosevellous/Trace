VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRowSelector 
   Caption         =   "Row Selector"
   ClientHeight    =   1965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "frmRowSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRowSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selStartRw As Integer

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage
End Sub

Private Sub btnOK_Click()

    If CInt(Me.txtRowsAbove.Value) = 0 And CInt(Me.txtRowsBelow.Value) = 0 Then End

If (T_FirstSelectedRow - CInt(Me.txtRowsAbove.Value)) < T_FirstRow Then
    T_FirstSelectedRow = T_FirstRow
Else
    T_FirstSelectedRow = T_FirstSelectedRow - CInt(Me.txtRowsAbove.Value)
End If

T_LastSelectedRow = T_FirstSelectedRow + CInt(Me.txtRowsBelow.Value) + CInt(Me.txtRowsAbove.Value) - 1

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub sbRowsAbove_Change()
Me.txtRowsAbove.Value = Me.sbRowsAbove.Value
UpdatePreview
End Sub

Private Sub sbRowsBelow_Change()
Me.txtRowsBelow.Value = Me.sbRowsBelow.Value
UpdatePreview
End Sub



Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
SetSheetTypeControls
btnOkPressed = False
'T_FirstSelectedRow = Selection.Row
End Sub

Sub UpdatePreview()

selStartRw = T_FirstSelectedRow - CInt(Me.txtRowsAbove.Value)
    
    If selStartRw < T_FirstRow Then
    selStartRw = T_FirstRow
    Me.sbRowsAbove.Value = Me.sbRowsAbove.Value - 1
    End If

Range(Cells(selStartRw, T_LossGainStart), _
    Cells(T_FirstSelectedRow + CInt(Me.txtRowsBelow.Value) - 1, T_LossGainEnd)).Select

End Sub
