VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlaneSource 
   Caption         =   "Plane source propagation"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   OleObjectBlob   =   "frmPlaneSource.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlaneSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Noise-Functions#plane")
End Sub

Private Sub btnOK_Click()
btnOkPressed = True
PlaneH = CDbl(Me.txtHeight.Value)
PlaneL = CDbl(Me.txtWidth.Value)
PlaneDist = CDbl(Me.txtDistance.Value)
Me.Hide
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub txtDistance_Change()
Call ValueChange
End Sub

Private Sub txtHeight_Change()
Call ValueChange
End Sub

Private Sub txtW_Change()
Call ValueChange
End Sub

Private Sub ValueChange()
Dim H As Double
Dim L As Double
Dim R As Double
Dim Atten As Double
    If IsNumeric(Me.txtHeight.Value) And IsNumeric(Me.txtWidth.Value) And IsNumeric(Me.txtDistance.Value) Then
    H = CDbl(Me.txtHeight.Value)
    L = CDbl(Me.txtWidth.Value)
    R = CDbl(Me.txtDistance.Value)
    
        If H = 0 Or L = 0 Or R = 0 Then
        Atten = 0
        Else
        Atten = -10 * Application.WorksheetFunction.Log10(H * L) + 10 * Application.WorksheetFunction.Log10(Atn((H * L) / (2 * R * Sqr((H ^ 2) + (L ^ 2) + (4 * R ^ 2))))) - 2
        Atten = Round(Atten, 1) 'eqn 5.105 of Beiss and Hansen
        End If
        
    Me.txtAtten.Value = Atten
    Else
    Me.txtAtten.Value = ""
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
