VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstCompressorSmall 
   Caption         =   "SWL Estimator - Compressor (Small)"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   OleObjectBlob   =   "frmEstCompressorSmall.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstCompressorSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Estimator-Functions#compressor")
End Sub

Private Sub btnOK_Click()

CalcLw

''''''''''''''''''''''''
'STORE VALUES
''''''''''''''''''''''''
CompressorSPL(0) = CLng(Me.txt31.Value)
CompressorSPL(1) = CLng(Me.txt63.Value)
CompressorSPL(2) = CLng(Me.txt125.Value)
CompressorSPL(3) = CLng(Me.txt250.Value)
CompressorSPL(4) = CLng(Me.txt500.Value)
CompressorSPL(5) = CLng(Me.txt1k.Value)
CompressorSPL(6) = CLng(Me.txt2k.Value)
CompressorSPL(7) = CLng(Me.txt4k.Value)
CompressorSPL(8) = CLng(Me.txt8k.Value)

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Sub SelectPower()
'static values
    If Me.optUpTo1_5kW.Value = True Then
    Me.txt31.Value = 82
    Me.txt63.Value = 81
    Me.txt125.Value = 81
    Me.txt250.Value = 80
    Me.txt500.Value = 83
    Me.txt1k.Value = 86
    Me.txt2k.Value = 86
    Me.txt4k.Value = 84
    Me.txt8k.Value = 81
    ElseIf Me.opt2to6kW.Value = True Then
    Me.txt31.Value = 87
    Me.txt63.Value = 84
    Me.txt125.Value = 84
    Me.txt250.Value = 83
    Me.txt500.Value = 86
    Me.txt1k.Value = 89
    Me.txt2k.Value = 89
    Me.txt4k.Value = 87
    Me.txt8k.Value = 84
    ElseIf Me.opt7to75kW.Value = True Then
    Me.txt31.Value = 92
    Me.txt63.Value = 87
    Me.txt125.Value = 87
    Me.txt250.Value = 86
    Me.txt500.Value = 89
    Me.txt1k.Value = 92
    Me.txt2k.Value = 92
    Me.txt4k.Value = 90
    Me.txt8k.Value = 87
    Else
    'error, do nothing????
    End If
End Sub

Sub CalcLw()

SelectPower

Me.txt31Lw.Value = CInt(Me.txt31.Value) + CInt(Me.txt31adj.Value)
Me.txt63Lw.Value = CInt(Me.txt63.Value) + CInt(Me.txt63adj.Value)
Me.txt125Lw.Value = CInt(Me.txt125.Value) + CInt(Me.txt125adj.Value)
Me.txt250Lw.Value = CInt(Me.txt250.Value) + CInt(Me.txt250adj.Value)
Me.txt500Lw.Value = CInt(Me.txt500.Value) + CInt(Me.txt500adj.Value)
Me.txt1kLw.Value = CInt(Me.txt1k.Value) + CInt(Me.txt1kadj.Value)
Me.txt2kLw.Value = CInt(Me.txt2k.Value) + CInt(Me.txt2kadj.Value)
Me.txt4kLw.Value = CInt(Me.txt4k.Value) + CInt(Me.txt4kadj.Value)
Me.txt8kLw.Value = CInt(Me.txt8k.Value) + CInt(Me.txt8kadj.Value)

End Sub

Private Sub opt2to6kW_Click()
CalcLw
End Sub

Private Sub opt7to75kW_Click()
CalcLw
End Sub

Private Sub optUpTo1_5kW_Click()
CalcLw
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

