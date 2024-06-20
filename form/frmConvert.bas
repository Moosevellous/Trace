VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvert 
   Caption         =   "Convert"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   OleObjectBlob   =   "frmConvert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHelp_Click()
GotoWikiPage ("Row-Functions#convert")
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
UserSelectedAddress = refRangeSelector.Value

    If Me.optSum = True Then
    SumAverageMode = "Sum"
    ElseIf Me.optAverage = True Then
    SumAverageMode = "Average"
    ElseIf Me.optLogAv.Value = True Then
    SumAverageMode = "Log Av"
    ElseIf Me.optTL.Value = True Then
    SumAverageMode = "TL"
    Else
    msg = MsgBox("What did you press? Option does not exist.", vbOKOnly, "ERROR!")
    End If

btnOkPressed = True
Me.Hide
End Sub
