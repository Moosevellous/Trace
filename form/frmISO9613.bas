VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmISO9613 
   Caption         =   "ISO9613-1:1996 Complete Calculation"
   ClientHeight    =   12540
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

Private Sub btnHelp_Click()
GotoWikiPage ("Noise-Functions#iso9613-2")
End Sub

Private Sub btnOK_Click()

ISOFullElements(0) = Me.chkAdiv.Value
ISOFullElements(1) = Me.chkAatm.Value
ISOFullElements(2) = Me.chkAgr.Value
ISOFullElements(3) = Me.chkAbar.Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Common Elements
    iso9613_d = Me.txtDistance.Value
    iso9613_d_ref = Me.txtDistanceRef.Value
    iso9613_SourceHeight = Me.txtSrcHeight.Value
    iso9613_ReceiverHeight = Me.txtRecHeight.Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Divergence
    If Me.chkAdiv.Value = True Then
    ISOFullElements(0) = True
    iso9613_d = Me.txtDistance.Value
    iso9613_d_ref = Me.txtDistanceRef.Value
    Else
    ISOFullElements(0) = False
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Air Absorption
    If Me.chkAatm.Value = True Then
    ISOFullElements(1) = True
    Else
    ISOFullElements(1) = False
    End If
    
    If Me.opt10degrees.Value = True Then
    iso9613_Temperature = 10
    ElseIf Me.opt15degrees = True Then
    iso9613_Temperature = 15
    ElseIf Me.opt20degrees.Value = True Then
    iso9613_Temperature = 20
    ElseIf Me.opt30degrees.Value = True Then
    iso9613_Temperature = 30
    Else
    'catch error
    End If
    
    If Me.opt20percentRH.Value = True Then
    iso9613_RelHumidity = 20
    ElseIf Me.opt50percentRH.Value = True Then
    iso9613_RelHumidity = 50
    ElseIf Me.opt70percentRH.Value = True Then
    iso9613_RelHumidity = 70
    ElseIf Me.opt80percentRH.Value = True Then
    iso9613_RelHumidity = 80
    Else
    'catch error
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Ground Absorption
    If Me.chkAgr.Value = True Then
    ISOFullElements(2) = True
    Else
    ISOFullElements(2) = False
    End If
    
    If Me.optGSrc0.Value = True Then
    iso9613_G_source = 0
    ElseIf Me.optGSrc50.Value = True Then
    iso9613_G_source = 0.5
    ElseIf Me.optGSrc100.Value = True Then
    iso9613_G_source = 1
    ElseIf Me.optGsrcCustom.Value = True Then
        If IsNumeric(Me.txtGsrcCustom.Value) Then
        iso9613_G_source = Me.txtGsrcCustom.Value
        Else
        iso9613_G_source = 0
        End If
    Else
    'catch error
    End If
    
    If Me.optGMid0.Value = True Then
    iso9613_G_middle = 0
    ElseIf Me.optGMid50.Value = True Then
    iso9613_G_middle = 0.5
    ElseIf Me.optGMid100.Value = True Then
    iso9613_G_middle = 1
    ElseIf Me.optGMidCustom.Value = True Then
        If IsNumeric(Me.txtGmidCustom.Value) Then
        iso9613_G_middle = Me.txtGmidCustom.Value
        Else
        iso9613_G_middle = 0
        End If
    Else
    'catch error
    End If
    
    If Me.optGRec0.Value = True Then
    iso9613_G_receiver = 0
    ElseIf Me.optGRec50.Value = True Then
    iso9613_G_receiver = 0.5
    ElseIf Me.optGRec100.Value = True Then
    iso9613_G_receiver = 1
    ElseIf Me.optGRecCustom.Value = True Then
        If IsNumeric(Me.txtGrecCustom.Value) Then
        iso9613_G_receiver = Me.txtGrecCustom.Value
        Else
        iso9613_G_receiver = 0
        End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Barrier attenuation
    If Me.chkAbar.Value = True Then
    ISOFullElements(3) = True
    iso9613_ReceiverHeight = Me.txtRecHeight.Value
    iso9613_SourceToBarrier = Me.txtSourceToBarrier.Value
    iso9613_SrcToBarrierEdge = Me.txtSrcToBarrierEdge.Value
    iso9613_RecToBarrierEdge = Me.txtRecToBarrierEdge.Value
    iso9613_BarrierHeight = Me.txtBarrierHeight.Value
    iso9613_BarrierHeightReceiverSide = Me.txtBarrierHeightRec.Value
    iso9613_DoubleDiffraction = Me.chkDoubleDiffraction.Value
    iso9613_BarrierThickness = Me.txtBarrierThickness.Value
    iso9613_MultiSource = Me.chkMultiSource.Value
    Else
    ISOFullElements(3) = False
    End If
    

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

Private Sub chkAbar_Click()

Abar_Agr_Check

    If Me.chkAbar.Value = True Then
    Me.frameAbar.Enabled = True
        For i = 0 To Me.frameAbar.Controls.Count - 1
        Me.frameAbar.Controls(i).Enabled = True
        Next i
    Else
    Me.frameAbar.Enabled = False
        For i = 0 To Me.frameAbar.Controls.Count - 1
        Me.frameAbar.Controls(i).Enabled = False
        Next i
    End If

PreviewAbar

End Sub

Sub Abar_Agr_Check()
    If Me.chkAbar.Value = True And Me.chkAgr.Value = False Then
    msg = MsgBox("Warning: Barrier effect depends on the ground effect. " & chr(10) & "Refer to the standard for more.", vbOKOnly, "Trickyyyyyyyy")
    Me.chkAgr.Value = True
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

Abar_Agr_Check

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

Private Sub chkDoubleDiffraction_Click()
    If Me.chkDoubleDiffraction = True Then
         If Me.txtBarrierThickness.Value = 0 Then Me.txtBarrierThickness.Value = 0.5
    Me.txtBarrierThickness.Enabled = True
    Me.txtBarrierHeightRec.Enabled = True
    Else
    Me.txtBarrierThickness.Enabled = False
    Me.txtBarrierHeightRec.Enabled = False
    Me.txtBarrierHeightRec.Value = Me.txtBarrierHeight.Value
    End If
PreviewAbar
End Sub

Private Sub chkMultiSource_Click()
PreviewAbar
End Sub

Private Sub lblAgrRec_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub txtBarrierHeight_Change()
    If Me.chkDoubleDiffraction.Value = False Then
    Me.txtBarrierHeightRec.Value = Me.txtBarrierHeight.Value
    End If
PreviewAbar
End Sub

Private Sub txtBarrierHeightRec_Change()
PreviewAbar
End Sub

Private Sub txtBarrierThickness_Change()
PreviewAbar
End Sub

Private Sub txtDistance_Change()
PreviewAdiv
PreviewAatm
PreviewAgr
PreviewAbar
End Sub

Private Sub txtGmidCustom_Change()
    If Me.txtGmidCustom.Value <= 1 Then
    PreviewAgr
    Else
    msg = MsgBox("Value must be no more than 1 (soft ground)", vbOKOnly, "Error - Ground Absorption")
    End If
End Sub

Private Sub txtGrecCustom_Change()
    If Me.txtGrecCustom.Value <= 1 Then
    PreviewAgr
    Else
    msg = MsgBox("Value must be no more than 1 (soft ground)", vbOKOnly, "Error - Ground Absorption")
    End If
End Sub

Private Sub txtGsrcCustom_Change()
    If Me.txtGsrcCustom.Value <= 1 Then
    PreviewAgr
    Else
    msg = MsgBox("Value must be no more than 1 (soft ground)", vbOKOnly, "Error - Ground Absorption")
    End If
End Sub

Private Sub txtRecHeight_Change()
PreviewAgr
PreviewAbar
End Sub

Private Sub txtSourceToBarrier_Change()
PreviewAbar
End Sub

Private Sub txtSrcHeight_Change()
PreviewAgr
PreviewAbar
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
'preview stuff
PreviewAdiv
PreviewAatm
PreviewAgr
PreviewAbar
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CalcDistances()
    If IsNumeric(Me.txtDistance) And IsNumeric(Me.txtRecToBarrier) And IsNumeric(Me.txtBarrierThickness.Value) Then
        If Me.chkDoubleDiffraction = True Then
        Me.txtRecToBarrier.Value = Me.txtDistance.Value - Me.txtSourceToBarrier.Value - Me.txtBarrierThickness
        Else
        Me.txtRecToBarrier.Value = Me.txtDistance.Value - Me.txtSourceToBarrier.Value
        End If
    End If
End Sub

Sub PreviewAdiv()
    If IsNumeric(Me.txtDistance.Value) And Me.txtDistance.Value <> "" Then
    Me.txtAdiv.Value = ISO9613_Adiv(Me.txtDistance.Value, Me.txtDistanceRef.Value)
    End If
End Sub


Sub PreviewAatm()
Dim PreviewTemperature As Integer
Dim PreviewRH As Integer
Dim PreviewSpectrum() As String

    'Italisize checkboxes
    If Me.opt15degrees.Value = True Then
'    Me.opt10degrees.Enabled = False
'    Me.opt15degrees.Enabled = True
'    Me.opt20degrees.Enabled = False
'    Me.opt30degrees.Enabled = False
    Me.opt20percentRH.Enabled = True
    Me.opt50percentRH.Enabled = True
    Me.opt70percentRH.Enabled = False
    Me.opt80percentRH.Enabled = True
    Else '10, 20 or 30 degrees
'    Me.opt10degrees.Enabled = True
'    Me.opt15degrees.Enabled = False
'    Me.opt20degrees.Enabled = True
'    Me.opt30degrees.Enabled = True
    Me.opt20percentRH.Enabled = False
    Me.opt50percentRH.Enabled = False
    Me.opt80percentRH.Enabled = False
    Me.opt70percentRH.Enabled = True
    Me.opt70percentRH.Value = True
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
    Me.txtAatm63.Value = Left(CStr(ISO9613_Aatm("63", Me.txtDistance.Value, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm125.Value = Left(CStr(ISO9613_Aatm("125", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm250.Value = Left(CStr(ISO9613_Aatm("250", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm500.Value = Left(CStr(ISO9613_Aatm("500", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm1k.Value = Left(CStr(ISO9613_Aatm("1k", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm2k.Value = Left(CStr(ISO9613_Aatm("2k", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm4k.Value = Left(CStr(ISO9613_Aatm("4k", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
    Me.txtAatm8k.Value = Left(CStr(ISO9613_Aatm("8k", Me.txtDistance, PreviewTemperature, PreviewRH)), 4)
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
    
    'calc values - round to 2 decimal places, as a string
    If IsNumeric(Me.txtSrcHeight.Value) And IsNumeric(Me.txtRecHeight.Value) And (Me.txtDistance.Value) Then
    Me.txtAgr63.Value = Round(ISO9613_Agr("63", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr125.Value = Round(ISO9613_Agr("125", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr250.Value = Round(ISO9613_Agr("250", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr500.Value = Round(ISO9613_Agr("500", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr1k.Value = Round(ISO9613_Agr("1k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr2k.Value = Round(ISO9613_Agr("2k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr4k.Value = Round(ISO9613_Agr("4k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    Me.txtAgr8k.Value = Round(ISO9613_Agr("8k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, G_source, G_receiver, G_middle), 1)
    End If
End Sub

Sub PreviewAbar()
CalcDistances
    'preview values
    If CheckAbarInputs = True Then
    Me.txtAbar63.Value = Round(ISO9613_Abar("63", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, _
    Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, _
    Me.txtAgr63.Value), 1)
    Me.txtAbar125.Value = Round(ISO9613_Abar("125", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr125.Value), 1)
    Me.txtAbar250.Value = Round(ISO9613_Abar("250", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr250.Value), 1)
    Me.txtAbar500.Value = Round(ISO9613_Abar("500", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr500.Value), 1)
    Me.txtAbar1k.Value = Round(ISO9613_Abar("1k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr1k.Value), 1)
    Me.txtAbar2k.Value = Round(ISO9613_Abar("2k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr2k.Value), 1)
    Me.txtAbar4k.Value = Round(ISO9613_Abar("4k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr4k.Value), 1)
    Me.txtAbar8k.Value = Round(ISO9613_Abar("8k", Me.txtSrcHeight.Value, Me.txtRecHeight.Value, Me.txtDistance.Value, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, Me.txtAgr8k.Value), 1)
    End If
    
End Sub

Function CheckAbarInputs()
    If IsNumeric(Me.txtSrcHeight.Value) And _
        IsNumeric(Me.txtRecHeight.Value) And _
        IsNumeric(Me.txtDistance.Value) And _
        IsNumeric(Me.txtSourceToBarrier.Value) And _
        IsNumeric(Me.txtSrcToBarrierEdge.Value) And _
        IsNumeric(Me.txtRecToBarrierEdge.Value) And _
        IsNumeric(Me.txtBarrierHeight.Value) And _
        IsNumeric(Me.txtBarrierThickness.Value) And _
        IsNumeric(Me.txtBarrierHeightRec.Value) Then
    CheckAbarInputs = True
    Else
    CheckAbarInputs = False
    End If
End Function
