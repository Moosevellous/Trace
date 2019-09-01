VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmISO9613 
   Caption         =   "ISO9613-1:1996 Complete Calculation"
   ClientHeight    =   10605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "frmISO9613.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmISO9613"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()

    If Me.chkAdiv.Value = True Then
    ISOFullElements(0) = True
    iso9613_d = Me.txtDistance.Value
    iso9613_d_ref = Me.txtDistanceRef.Value
    Else
    ISOFullElements(0) = False
    End If
    
    If Me.chkAatm.Value = True Then
    ISOFullElements(1) = True
    Else
    ISOFullElements(1) = False
    End If
    
    If Me.chkAgr.Value = True Then
    ISOFullElements(2) = True
    Else
    ISOFullElements(2) = False
    End If
    
    If Me.chkAbar.Value = True Then
    ISOFullElements(3) = True
    Else
    ISOFullElements(3) = False
    End If

'If Me.chkAdiv.Value = True Then ISOFullElements(0) = True 'Amisc

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub chkAatm_Click()
    If Me.chkAatm.Value = True Then
    Me.frameAatm.Enabled = True
        For i = 0 To Me.frameAatm.Controls.Count - 1
        Me.frameAatm.Controls(i).Enabled = True
        Next i
    PreviewAatm
    Else
    Me.frameAatm.Enabled = False
        For i = 0 To Me.frameAatm.Controls.Count - 1
        Me.frameAatm.Controls(i).Enabled = False
        Next i
    End If
End Sub

Private Sub chkAdiv_Click()
    If Me.chkAdiv.Value = True Then
    Me.frameAdiv.Enabled = True
        For i = 0 To Me.frameAdiv.Controls.Count - 1
        Me.frameAdiv.Controls(i).Enabled = True
        Next i
    Else
    Me.frameAdiv.Enabled = False
        For i = 0 To Me.frameAdiv.Controls.Count - 1
        Me.frameAdiv.Controls(i).Enabled = False
        Next i
    End If
End Sub

Private Sub chkAgr_Click()
    If Me.chkAgr.Value = True Then
    Me.frameAgr.Enabled = True
        For i = 0 To Me.frameAgr.Controls.Count - 1
        Me.frameAgr.Controls(i).Enabled = True
        Next i
    Else
    Me.frameAgr.Enabled = False
        For i = 0 To Me.frameAgr.Controls.Count - 1
        Me.frameAgr.Controls(i).Enabled = False
        Next i
    End If
End Sub

Private Sub imgAgr_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub


Private Sub lblAgrRec_Click()

End Sub

Private Sub lblAgrRec_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Me.imgAgr.Visible = True
End Sub

Private Sub opt10degrees_Click()
PreviewAatm
End Sub

Private Sub opt15degrees_Click()
PreviewAatm
End Sub

Private Sub opt20degrees_Click()
PreviewAatm
End Sub

Private Sub opt20percentRH_Click()
PreviewAatm
End Sub

Private Sub opt30degrees_Click()
PreviewAatm
End Sub

Private Sub opt50percentRH_Click()
PreviewAatm
End Sub

Private Sub opt70percentRH_Click()
PreviewAatm
End Sub

Private Sub opt80percentRH_Click()
PreviewAatm
End Sub

'ground
Private Sub optGMid0_Click()
PreviewAgr
End Sub

Private Sub optGMid100_Click()
PreviewAgr
End Sub

Private Sub optGMid50_Click()
PreviewAgr
End Sub

Private Sub optGMidCustom_Click()
PreviewAgr
End Sub

Private Sub optGRec0_Click()
PreviewAgr
End Sub

Private Sub optGRec100_Click()
PreviewAgr
End Sub

Private Sub optGRec50_Click()
PreviewAgr
End Sub

Private Sub optGRecCustom_Click()
PreviewAgr
End Sub

Private Sub optGSrc0_Click()
PreviewAgr
End Sub

Private Sub optGSrc100_Click()
PreviewAgr
End Sub

Private Sub optGSrc50_Click()
PreviewAgr
End Sub

Private Sub optGsrcCustom_Click()
PreviewAgr
End Sub

Private Sub txtDistance_Change()
PreviewAdiv
PreviewAatm
PreviewAgr
End Sub

Private Sub txtGmidCustom_Change()
PreviewAgr
End Sub

Private Sub txtGrecCustom_Change()
PreviewAgr
End Sub

Private Sub txtGsrcCustom_Change()
PreviewAgr
End Sub

Private Sub UserForm_Activate()
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
'preview stuff
PreviewAdiv
PreviewAatm
PreviewAgr
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub PreviewAdiv()
    If IsNumeric(Me.txtDistance.Value) And Me.txtDistance.Value <> "" Then
    Me.txtAdiv.Value = ISO9613_Adiv(Me.txtDistance.Value, Me.txtDistanceRef.Value)
    End If
End Sub


Sub PreviewAatm()
Dim PreviewTemperature As Integer
Dim PreviewRH As Integer

    'Italisize checkboxes
    If Me.opt15degrees.Value = True Then
    Me.opt10degrees.Font.Italic = True
    Me.opt15degrees.Font.Italic = False
    Me.opt20degrees.Font.Italic = True
    Me.opt30degrees.Font.Italic = True
    Me.opt20percentRH.Font.Italic = False
    Me.opt50percentRH.Font.Italic = False
    Me.opt70percentRH.Font.Italic = True
    Me.opt80percentRH.Font.Italic = False
    Else
    Me.opt10degrees.Font.Italic = False
    Me.opt15degrees.Font.Italic = True
    Me.opt20degrees.Font.Italic = False
    Me.opt30degrees.Font.Italic = False
    Me.opt20percentRH.Font.Italic = True
    Me.opt50percentRH.Font.Italic = True
    Me.opt70percentRH.Font.Italic = False
    Me.opt80percentRH.Font.Italic = True
    End If

    'set temperature and relative humidity values
    If Me.opt10degrees.Value = True Then
    PreviewTemperature = 10
    ElseIf Me.opt15degrees.Value = True Then
    PreviewTemperature = 15
    ElseIf Me.opt20degrees.Value = True Then
    PreviewTemperature = 20
    ElseIf Me.opt30degrees.Value = True Then
    PreviewTemperature = 30
    Else
    PreviewTemperature = 0
    End If
    
    If Me.opt20percentRH.Value = True Then
    PreviewRH = 20
    ElseIf Me.opt50percentRH.Value = True Then
    PreviewRH = 50
    ElseIf Me.opt70percentRH.Value = True Then
    PreviewRH = 70
    ElseIf Me.opt80percentRH.Value = True Then
    PreviewRH = 80
    Else
    PreviewRH = 0
    End If
    
    'preview values
    If Me.txtDistance.Value <> "" Then
    Me.txtAatm63.Value = ISO9613_Aatm("63", Me.txtDistance.Value, PreviewTemperature, PreviewRH)
    Me.txtAatm125.Value = ISO9613_Aatm("125", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm250.Value = ISO9613_Aatm("250", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm500.Value = ISO9613_Aatm("500", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm1k.Value = ISO9613_Aatm("1k", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm2k.Value = ISO9613_Aatm("2k", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm4k.Value = ISO9613_Aatm("4k", Me.txtDistance, PreviewTemperature, PreviewRH)
    Me.txtAatm8k.Value = ISO9613_Aatm("8k", Me.txtDistance, PreviewTemperature, PreviewRH)
    End If
    
End Sub

Sub PreviewAgr()
Dim G_source As Double
Dim G_receiver As Double
Dim G_middle As Double

    'source
    If Me.optGSrc0.Value = True Then
    G_source = 0
    ElseIf Me.optGSrc50.Value = True Then
    G_source = 0.5
    ElseIf Me.optGSrc100.Value = True Then
    G_source = 1
    ElseIf Me.optGsrcCustom.Value = True Then
        If IsNumeric(Me.txtGsrcCustom.Value) Then
        G_source = Me.txtGsrcCustom.Value
        Else
        G_source = 0
        End If
    End If

    'middle
    If Me.optGMid0.Value = True Then
    G_middle = 0
    ElseIf Me.optGMid50.Value = True Then
    G_middle = 0.5
    ElseIf Me.optGMid100.Value = True Then
    G_middle = 1
    ElseIf Me.optGMidCustom.Value = True Then
        If IsNumeric(Me.txtGmidCustom.Value) Then
        G_middle = Me.txtGmidCustom.Value
        Else
        G_middle = 0
        End If
    End If
    
    'receiver
    If Me.optGRec0.Value = True Then
    G_receiver = 0
    ElseIf Me.optGRec50.Value = True Then
    G_receiver = 0.5
    ElseIf Me.optGRec100.Value = True Then
    G_receiver = 1
    ElseIf Me.optGRecCustom.Value = True Then
        If IsNumeric(Me.txtGrecCustom.Value) Then
        G_receiver = Me.txtGrecCustom.Value
        Else
        G_receiver = 0
        End If
    End If
    
'calc Values
Me.txtAgr63.Value = Round(ISO9613_Agr("63", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr125.Value = Round(ISO9613_Agr("125", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr250.Value = Round(ISO9613_Agr("250", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr500.Value = Round(ISO9613_Agr("500", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr1k.Value = Round(ISO9613_Agr("1k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr2k.Value = Round(ISO9613_Agr("2k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr4k.Value = Round(ISO9613_Agr("4k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
Me.txtAgr8k.Value = Round(ISO9613_Agr("8k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
End Sub
