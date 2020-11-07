VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVibUnits 
   Caption         =   "Vibration - Convert Units"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "frmVibUnits.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVibUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PowerOf As Integer

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Vibration#convert-units")
End Sub

Private Sub btnOK_Click()
UpdateConversionFactor
btnOkPressed = True

Me.Hide
Unload Me
End Sub

Private Sub optAccel_Click()
Me.optMetres.Caption = "m/s" & chr(178)
Me.optMillimetres.Caption = "mm/s" & chr(178)
UpdateConversionFactor
End Sub

Private Sub optDisplacement_Click()
Me.optMetres.Caption = "m"
Me.optMillimetres.Caption = "mm"
UpdateConversionFactor
End Sub

Private Sub optMetres_Click()
UpdateConversionFactor
End Sub

Private Sub optMillimetres_Click()
UpdateConversionFactor
End Sub

Private Sub optVelocity_Click()
Me.optMetres.Caption = "m/s"
Me.optMillimetres.Caption = "mm/s"
UpdateConversionFactor
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
UpdateConversionFactor
End Sub

Sub UpdateConversionFactor()

    'select x in 10^x
    If Me.optAccel.Value Then
    PowerOf = -6
    ElseIf Me.optVelocity.Value Then
    PowerOf = -9
    ElseIf Me.optDisplacement.Value Then
    PowerOf = -12
    Else
    msg = MsgBox("Error, no value selected", vbOKOnly, "You must choooooooose")
    End
    End If
    
    'set public variable
    If Me.optMetres.Value Then
    VibRef = "1e" & CStr(PowerOf)
    ElseIf Me.optMillimetres.Value Then
    VibRef = "1e" & CStr(PowerOf + 3)
    Else
    msg = MsgBox("Error, no value selected", vbOKOnly, "You must choooooooose")
    End
    End If

'show in text box
Me.txtConversionFactor.Value = VibRef
    
End Sub
