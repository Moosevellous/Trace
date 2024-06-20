VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecorder 
   Caption         =   "Calc Recorder"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   OleObjectBlob   =   "frmRecorder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecordingInProgress As Boolean

Private Sub btnDone_Click()
'Me.Hide
Unload Me
End Sub


Private Sub btnHelp_Click()
GotoWikiPage ("Sheet-Functions#plot")
End Sub

Private Sub btnPlayCalcBlock_Click()
If Len(Me.lstBlocks.Value) = 0 Then
    msg = MsgBox("Booooo!", vbOKOnly)
End If
End Sub

Private Sub btnRecStop_Click()
If RecordingInProgress = False Then 'start recording
    If Me.txtBlockName.Value = "" Then
    Me.txtBlockName.Value = "Block_1"
    End If
    Me.btnRecStop.Caption = "Stop Recording"
    Me.txtBlockName.Enabled = False
    RecordingInProgress = True
Else 'stop
    Me.btnRecStop.Caption = "Record Calculation Block"
    Me.txtBlockName.Enabled = True
    RecordingInProgress = False
    Me.lstBlocks.AddItem Me.txtBlockName.Value
    Me.txtBlockName.Value = ""
End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
RecordingInProgress = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

