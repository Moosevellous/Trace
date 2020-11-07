VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTarget 
   Caption         =   "Set Target / Limit"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "frmTarget.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim defaultCompliant As Long
Dim defaultMargin As Long
Dim defaultLimit As Long

Private Sub UserForm_Initialize()
'set default colours
defaultCompliant = RGB(146, 208, 80)
defaultMargin = RGB(255, 235, 156)
defaultLimit = RGB(224, 68, 68)
DefaultColours
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnCompliantColour_Click()
Application.Dialogs(xlDialogEditColor).Show 1, defaultCompliant Mod 256, (defaultCompliant \ 256) Mod 256, (defaultCompliant \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.CompliantColourBox.BackColor = FullColourCode
End Sub

Private Sub btnMarginColour_Click()
Application.Dialogs(xlDialogEditColor).Show 1, defaultMargin Mod 256, (defaultMargin \ 256) Mod 256, (defaultMargin \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.MarginColourBox.BackColor = FullColourCode
End Sub

Private Sub btnLimitColours_Click()
Application.Dialogs(xlDialogEditColor).Show 1, defaultLimit Mod 256, (defaultLimit \ 256) Mod 256, (defaultLimit \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.LimitColourBox.BackColor = FullColourCode
End Sub

Private Sub btnDefaultColours_Click()
DefaultColours
End Sub

Private Sub btnOK_Click()
'check target type
targetType = "" 'public variable
    If Me.optdB.Value = True Then targetType = "dB"
    If Me.optdBA.Value = True Then targetType = "dBA"
    If Me.optdBC.Value = True Then targetType = "dBC"
    If Me.optNR.Value = True Then targetType = "NR"
    If Me.optBand.Value = True Then targetType = "Band"
    If targetType = "" Then ErrorIncompleteForm 'no target selected

    If Me.optWholeValue.Value = True Then
    targetRoundingWholeNumber = True
    Else
    targetRoundingWholeNumber = False
    End If

    If Len(Me.txtLimitVal) > 0 Then
    targetLimitValue = Me.txtLimitVal.Value
    End If

'    If Me.chkEnableMargin.Value = True And Len(Me.txtMarginVal.Value) > 0 Then
'    targetMarginValue = Me.txtMarginVal.Value
'    End If

    If Len(Me.txtCompliantVal.Value) > 0 Then
    targetCompliantValue = Me.txtCompliantVal.Value
    End If
    
    'colurs
    targetLimitColour = Me.LimitColourBox.BackColor
    targetMarginColour = Me.MarginColourBox.BackColor
    targetCompliantColour = Me.CompliantColourBox.BackColor
    
btnOkPressed = True
Me.Hide
End Sub

'Private Sub chkEnableCompliant_Click()
'
'    If Me.chkEnableCompliant.Value = True Then
'    Me.txtCompliantVal.Enabled = True
'    Me.lblCompliant.Enabled = True
'    Else
'    Me.txtCompliantVal.Enabled = False
'    Me.lblCompliant.Enabled = False
'    End If
'
'End Sub
'
'Private Sub chkEnableLimit_Click()
'
'    If Me.chkEnableLimit.Value = True Then
'    Me.txtLimitVal.Enabled = True
'    Me.lblLimit.Enabled = True
'    Else
'    Me.txtLimitVal.Enabled = False
'    Me.lblLimit.Enabled = False
'    End If
'
'End Sub
'
'Private Sub chkEnableMargin_Click()
'
'    If Me.chkEnableMargin.Value = True Then
'    Me.txtMarginVal.Enabled = True
'    Me.lblMarginal.Enabled = True
'    Else
'    Me.txtMarginVal.Enabled = False
'    Me.lblMarginal.Enabled = False
'    End If
'
'End Sub

Private Sub txtTarget_Change()
UpdateValues
End Sub


Sub UpdateValues()

End Sub


Private Sub CompliantColourBox_Change()
Me.CompliantColourBox.Text = ""
End Sub

Private Sub CompliantColourBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Dialogs(xlDialogEditColor).Show 1, defaultCompliant Mod 256, (defaultCompliant \ 256) Mod 256, (defaultCompliant \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.CompliantColourBox.BackColor = FullColourCode
End Sub

Private Sub LimitColourBox_Change()
Me.LimitColourBox.Text = ""
End Sub

Private Sub LimitColourBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Dialogs(xlDialogEditColor).Show 1, defaultLimit Mod 256, (defaultLimit \ 256) Mod 256, (defaultLimit \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.LimitColourBox.BackColor = FullColourCode
End Sub


Private Sub MarginColourBox_Change()
Me.MarginColourBox.Text = ""
End Sub

Private Sub MarginColourBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Application.Dialogs(xlDialogEditColor).Show 1, defaultMargin Mod 256, (defaultMargin \ 256) Mod 256, (defaultMargin \ 256 \ 256) Mod 256
FullColourCode = ActiveWorkbook.Colors(1)
Me.MarginColourBox.BackColor = FullColourCode
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
End Sub

Sub DefaultColours()
Me.CompliantColourBox.BackColor = defaultCompliant
Me.MarginColourBox.BackColor = defaultMargin
Me.LimitColourBox.BackColor = defaultLimit
End Sub

Sub ErrorIncompleteForm()
msg = MsgBox("Error - Please select your target options.", vbOKOnly, "Form incomplete")
End Sub
