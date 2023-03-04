VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBarrierAtten 
   Caption         =   "Barrier Attenuation"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   OleObjectBlob   =   "frmBarrierAtten.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBarrierAtten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------
'If all input is numeric and line of sight condition
'-----------------------------------------------------------------------
Function InputsOkay() As Boolean

    If IsNumeric(Me.txtBarrierHeight.Value) And _
        IsNumeric(Me.txtSourceToBarrier.Value) And _
        IsNumeric(Me.txtSrcHeight.Value) And _
        IsNumeric(Me.txtSrcGroundHeight.Value) And _
        IsNumeric(Me.txtRectoBarrier.Value) And _
        IsNumeric(Me.txtRecHeight.Value) And _
        IsNumeric(Me.txtRecGroundHeight.Value) And _
        IsNumeric(Me.txtSrcToBarrierEdge.Value) And _
        IsNumeric(Me.txtRecToBarrierEdge.Value) And _
        IsNumeric(Me.txtBarrierThickness.Value) And _
        IsNumeric(Me.txtBarrierHeightRec.Value) Then
        
        'Line of sight condition
        If BarrierCutsLineofSight(Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, _
            Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, _
            Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value) Then
            
            InputsOkay = True
            lblNotif.Visible = False
        Else
            InputsOkay = False
            lblNotif.Visible = True
        End If
        
    Else
        InputsOkay = False
    End If
    
End Function

Function SpreadingType() As String
    
    If Me.optPlane = True Then
    SpreadingType = "Plane"
    
    ElseIf Me.optCylindrical.Value = True Then
    SpreadingType = "Cylindrical"
    
    ElseIf Me.optSpherical.Value = True Then
    SpreadingType = "Spherical"
    
    Else 'default to nothing
    SpreadingType = "-"
    End If
        
End Function


Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Noise-Functions#Barrier")
End Sub

Private Sub btnOK_Click()

btnOkPressed = True

'set public variables
    
    'Method Identifier
    If Me.optKA.Value = True Then
    Barrier_Method = "KurzeAnderson"
    ElseIf Me.optISO9613.Value = True Then
    Barrier_Method = "ISO9613_Abar"
    ElseIf Me.optMenounou.Value = True Then
    Barrier_Method = "Menounou"
    End If


    'ISO method
    If Me.optISO9613.Value = True Then
    
    TotalRecHeight = CDbl(Me.txtRecGroundHeight.Value) + CDbl(Me.txtRecHeight.Value)
    TotalSrcHeight = CDbl(Me.txtSrcGroundHeight.Value) + CDbl(Me.txtSrcHeight.Value)
    SrcRecDistance = CDbl(Me.txtSourceToBarrier.Value) + CDbl(Me.txtRectoBarrier.Value)
    
    iso9613_d = SrcRecDistance
    iso9613_SourceHeight = TotalSrcHeight
    iso9613_ReceiverHeight = TotalRecHeight
    
    iso9613_SourceToBarrier = Me.txtSourceToBarrier.Value
    iso9613_SrcToBarrierEdge = Me.txtSrcToBarrierEdge.Value
    iso9613_RecToBarrierEdge = Me.txtRecToBarrierEdge.Value
    iso9613_BarrierHeight = Me.txtBarrierHeight.Value
    iso9613_BarrierHeightReceiverSide = Me.txtBarrierHeightRec.Value
    iso9613_DoubleDiffraction = Me.chkDoubleDiffraction.Value
    iso9613_BarrierThickness = Me.txtBarrierThickness.Value
    iso9613_MultiSource = Me.chkMultiSource.Value
    
    Else 'other methods
    Barrier_SourceToBarrier = Me.txtSourceToBarrier.Value
    Barrier_SourceHeight = Me.txtSrcHeight.Value
    Barrier_GroundUnderSrc = Me.txtSrcGroundHeight.Value
    Barrier_RecToBarrier = Me.txtRectoBarrier.Value
    Barrier_ReceiverHeight = Me.txtRecHeight.Value
    Barrier_GroundUnderRec = Me.txtRecGroundHeight.Value
    Barrier_BarrierHeight = Me.txtBarrierHeight.Value
    Barrier_SpreadingType = SpreadingType
    Barrier_SrcToBarrierEdge = Me.txtSrcToBarrierEdge.Value
    Barrier_RecToBarrierEdge = Me.txtRecToBarrierEdge.Value
    Barrier_BarrierHeightReceiverSide = Me.txtBarrierHeightRec.Value
    Barrier_DoubleDiffraction = Me.chkDoubleDiffraction.Value
    Barrier_BarrierThickness = Me.txtBarrierThickness.Value
    Barrier_MultiSource = Me.chkMultiSource.Value
    Barrier_SrcRecDistance = CDbl(Me.txtSourceToBarrier.Value) + CDbl(Me.txtRectoBarrier.Value)
    Barrier_GtoRecheight = CDbl(Me.txtRecHeight) + CDbl(Me.txtRecGroundHeight.Value)
    Barrier_GtoSrcHeight = CDbl(Me.txtSrcHeight.Value) + CDbl(Me.txtSrcGroundHeight.Value)
    End If

Me.Hide
Unload Me
    
End Sub

Private Sub optCylindrical_Click()
SpreadingType
UpdatePreview
End Sub

Private Sub chkDoubleDiffraction_Click()
SelectControls
AdjustForThickBarrier
UpdatePreview
End Sub

Private Sub chkGroundReflections_Click()
UpdatePreview
End Sub

Private Sub chkMultiSource_Click()
UpdatePreview
End Sub

Private Sub optPlane_Click()
SpreadingType
UpdatePreview
End Sub

Private Sub optSpherical_Click()
SpreadingType
UpdatePreview
End Sub

Private Sub optISO9613_Click()
SelectControls
UpdatePreview
End Sub

Private Sub optKA_Click()
SelectControls
UpdatePreview
End Sub

Private Sub optMenounou_Click()
SelectControls
UpdatePreview
End Sub

Private Sub txtBarrierHeight_Change()
UpdatePreview
End Sub

Private Sub txtBarrierHeightRec_Change()
UpdatePreview
End Sub

Private Sub txtBarrierThickness_Change()
AdjustForThickBarrier
UpdatePreview
End Sub

Private Sub txtRecGroundHeight_Change()
UpdatePreview
End Sub

Private Sub txtRecHeight_Change()
UpdatePreview
End Sub

Private Sub txtRecToBar_Change()
UpdatePreview
End Sub

Private Sub txtRectoBarrier_Change()
UpdatePreview
End Sub

Private Sub txtRecToBarrierEdge_Change()
UpdatePreview
End Sub

Private Sub txtSourceToBar_Change()
UpdatePreview
End Sub

Private Sub txtSourceToBarrier_Change()
UpdatePreview
End Sub

Private Sub txtSrcGroundHeight_Change()
UpdatePreview
End Sub

Private Sub txtSrcHeight_Change()
UpdatePreview
End Sub


Private Sub txtSrcToBarrierEdge_Change()
UpdatePreview
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
UpdatePreview
    
End Sub




'-----------------------------------------------------------------------
'Switch buttons on and off
'-----------------------------------------------------------------------
Sub SelectControls()
    'enable/disable buttons for ISO9613 frame
    If Me.optISO9613.Value = True Then
    
        For i = 0 To Me.fraISOOptions.Controls.Count - 1
        Me.fraISOOptions.Controls(i).Enabled = True
        Next
        
        'For Double Diffraction Selection
        If Me.chkDoubleDiffraction = True Then
        Me.txtBarrierThickness.Enabled = True
        Me.txtBarrierHeightRec.Enabled = True
        Else
        Me.txtBarrierThickness.Enabled = False
        Me.txtBarrierHeightRec.Enabled = False
        End If

'    Me.txtRecToBar.Enabled = False
'    Me.txtSourceToBar.Enabled = False

    Else
        'turn off all controls in the ISO frame
        For i = 0 To Me.fraISOOptions.Controls.Count - 1
        Me.fraISOOptions.Controls(i).Enabled = False
        Next
        
    End If
    
    'enable/disable buttons for Menounou frame
    If Me.optMenounou.Value = True Then
    
        For i = 0 To Me.fraMenounou.Controls.Count - 1
        Me.fraMenounou.Controls(i).Enabled = True
        Next
        
    Else
        
        For i = 0 To Me.fraMenounou.Controls.Count - 1
        Me.fraMenounou.Controls(i).Enabled = False
        Next
    End If
    
End Sub


'PS COMMENT
'It's a nice idea but I think it causes more problems than it solves
Sub AdjustForThickBarrier()
    If Me.chkDoubleDiffraction = True Then
        If IsNumeric(Me.txtBarrierThickness.Value) Then
        Me.txtSourceToBarrier.Value = Me.txtSourceToBarrier.Value - (Me.txtBarrierThickness.Value / 2)
        Me.txtRectoBarrier.Value = Me.txtRectoBarrier.Value - (Me.txtBarrierThickness.Value / 2)
        End If
    End If
End Sub

Sub UpdatePreview()

Dim TotalRecHeight As Double
Dim TotalSrcHeight As Double
Dim SrcRecDistance As Double

    If InputsOkay Then
    
        If Me.optKA.Value = True Then
            
        Me.txt31.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("31.5", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt63.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("63", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt125.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("125", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt250.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("250", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt500.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("500", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt1k.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("1k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt2k.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("2k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt4k.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("4k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        Me.txt8k.Value = CheckNumericValue(BarrierAtten_KurzeAnderson("8k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value), 1)
        
        ElseIf Me.optISO9613.Value = True Then
        
        TotalRecHeight = CDbl(Me.txtRecGroundHeight.Value) + CDbl(Me.txtRecHeight.Value)
        TotalSrcHeight = CDbl(Me.txtSrcGroundHeight.Value) + CDbl(Me.txtSrcHeight.Value)
        SrcRecDistance = CDbl(Me.txtSourceToBarrier.Value) + CDbl(Me.txtRectoBarrier.Value)
            
        Me.txt31.Value = CheckNumericValue(ISO9613_Abar("31.5", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt63.Value = CheckNumericValue(ISO9613_Abar("63", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt125.Value = CheckNumericValue(ISO9613_Abar("125", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt250.Value = CheckNumericValue(ISO9613_Abar("250", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt500.Value = CheckNumericValue(ISO9613_Abar("500", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt1k.Value = CheckNumericValue(ISO9613_Abar("1k", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt2k.Value = CheckNumericValue(ISO9613_Abar("2k", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt4k.Value = CheckNumericValue(ISO9613_Abar("4k", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
        Me.txt8k.Value = CheckNumericValue(ISO9613_Abar("8k", TotalSrcHeight, TotalRecHeight, SrcRecDistance, Me.txtSourceToBarrier.Value, Me.txtSrcToBarrierEdge.Value, Me.txtRecToBarrierEdge.Value, Me.txtBarrierHeight.Value, Me.chkDoubleDiffraction.Value, Me.txtBarrierThickness.Value, Me.txtBarrierHeightRec.Value, Me.chkMultiSource.Value, 3), 1)
            
        ElseIf Me.optMenounou.Value = True Then

        Me.txt31.Value = CheckNumericValue(BarrierAtten_Menounou("31.5", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt63.Value = CheckNumericValue(BarrierAtten_Menounou("63", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt125.Value = CheckNumericValue(BarrierAtten_Menounou("125", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt250.Value = CheckNumericValue(BarrierAtten_Menounou("250", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt500.Value = CheckNumericValue(BarrierAtten_Menounou("500", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt1k.Value = CheckNumericValue(BarrierAtten_Menounou("1k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt2k.Value = CheckNumericValue(BarrierAtten_Menounou("2k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt4k.Value = CheckNumericValue(BarrierAtten_Menounou("4k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
        Me.txt8k.Value = CheckNumericValue(BarrierAtten_Menounou("8k", Me.txtSourceToBarrier.Value, Me.txtSrcHeight.Value, Me.txtSrcGroundHeight.Value, Me.txtRectoBarrier.Value, Me.txtRecHeight.Value, Me.txtRecGroundHeight.Value, Me.txtBarrierHeight.Value, SpreadingType), 1)
            
        End If
        
    Else 'something's wrong, return nothing!
        Me.txt31.Value = "-"
        Me.txt63.Value = "-"
        Me.txt125.Value = "-"
        Me.txt250.Value = "-"
        Me.txt500.Value = "-"
        Me.txt1k.Value = "-"
        Me.txt2k.Value = "-"
        Me.txt4k.Value = "-"
        Me.txt8k.Value = "-"
    End If
    
End Sub
