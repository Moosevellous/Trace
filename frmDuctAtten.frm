VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuctAtten 
   Caption         =   "Duct Attenuation"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   OleObjectBlob   =   "frmDuctAtten.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuctAtten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optCir_Click()
Me.txtW.Enabled = False
End Sub

Private Sub optRect_Click()
Me.txtW.Enabled = True
End Sub

Private Sub UserForm_Activate()
    With frmDuctAtten
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
frmDuctAtten.Hide
End Sub

Private Sub btnOK_Click()
ductL = CInt(Me.txtL)
ductW = CInt(Me.txtW)
ductShape = getDuctShape
btnOkPressed = True
frmDuctAtten.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lvlX_Click()

End Sub

Private Sub txtL_Change()

End Sub



Private Function getDuctShape()

    If Me.opt25mm.Value Then
    W = 25
    ElseIf Me.opt50mm.Value Then
    W = 50
    ElseIf Me.optUnlined.Value Then
    W = 0
    End If
    
    If Me.optCir.Value Then
    s = "C"
    ElseIf Me.optRect.Value Then
    s = "R"
    End If

getDuctShape = CStr(W) & " " & s

End Function
