VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmERL 
   Caption         =   "End Reflection Loss"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   OleObjectBlob   =   "frmERL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmERL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Unload Me
End Sub

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#end-reflection-loss-erl")
End Sub

Private Sub btnOK_Click()
    
    If Me.optASHRAE.Value = True Then
    ERL_Mode = "ASHRAE"
    Else 'NEBB
    ERL_Mode = "NEBB"
    End If

    If Me.optFlush.Value = True Then
    ERL_Termination = "Flush"
    Else
    ERL_Termination = "Free"
    End If

    If IsNumeric(Me.txtArea.Value) Then
    ERL_Area = Me.txtArea.Value
    End If
    
    If Me.optCircular.Value = True Then
    ERL_Circular = True
    Else
    ERL_Circular = False
    End If
    
btnOkPressed = True
Unload Me
End Sub

Private Sub btnStandard600_Click()
Me.optRectangular.Value = True
Me.txtW.Value = 600
Me.txtL.Value = 600
End Sub

Private Sub CommandButton1_Click()
Me.optRectangular.Value = True
Me.txtW.Value = 300
Me.txtL.Value = 300
End Sub

Private Sub optASHRAE_Click()
PreviewERL
End Sub

Private Sub optCircular_Click()
Me.txtL.Enabled = False
Me.txtW.Enabled = False
Me.txtDia.Enabled = True
PreviewERL
End Sub

Private Sub optFlush_Click()
PreviewERL
End Sub

Private Sub optFree_Click()
PreviewERL
End Sub

Private Sub optNEBB_Click()
PreviewERL
End Sub

Sub PreviewERL()
Dim TerminationType As String
    
    If Me.optRectangular.Value = True Then
        If IsNumeric(Me.txtL.Value) And IsNumeric(Me.txtW.Value) Then
        Me.txtArea.Value = Round((Me.txtL.Value / 1000) * (Me.txtW.Value / 1000), 3)
        Else
        Me.txtArea.Value = ""
        End If
    ElseIf Me.optCircular.Value = True Then
        If IsNumeric(Me.txtDia.Value) = True Then
        Me.txtArea.Value = Round(Application.WorksheetFunction.Pi * (Me.txtDia.Value / 2000) ^ 2, 3)
        Else
        Me.txtArea.Value = ""
        End If
    End If
    
    If Me.txtArea <> "" Then
        If Me.optFree.Value = True Then
        TerminationType = "Free"
        ElseIf Me.optFlush.Value = True Then
        TerminationType = "Flush"
        End If
    
        If Me.optASHRAE.Value = True Then
        Me.txt31.Value = Round(ERL_ASHRAE(TerminationType, "31.5", Me.txtArea.Value), 1)
        Me.txt63.Value = Round(ERL_ASHRAE(TerminationType, "63", Me.txtArea.Value), 1)
        Me.txt125.Value = Round(ERL_ASHRAE(TerminationType, "125", Me.txtArea.Value), 1)
        Me.txt250.Value = Round(ERL_ASHRAE(TerminationType, "250", Me.txtArea.Value), 1)
        Me.txt500.Value = Round(ERL_ASHRAE(TerminationType, "500", Me.txtArea.Value), 1)
        Me.txt1k.Value = Round(ERL_ASHRAE(TerminationType, "1k", Me.txtArea.Value), 1)
        Me.txt2k.Value = Round(ERL_ASHRAE(TerminationType, "2k", Me.txtArea.Value), 1)
        Me.txt4k.Value = Round(ERL_ASHRAE(TerminationType, "4k", Me.txtArea.Value), 1)
        Me.txt8k.Value = Round(ERL_ASHRAE(TerminationType, "8k", Me.txtArea.Value), 1)
        ElseIf Me.optNEBB.Value = True Then
        Me.txt31.Value = Round(ERL_NEBB(TerminationType, "31.5", Me.txtArea.Value), 1)
        Me.txt63.Value = Round(ERL_NEBB(TerminationType, "63", Me.txtArea.Value), 1)
        Me.txt125.Value = Round(ERL_NEBB(TerminationType, "125", Me.txtArea.Value), 1)
        Me.txt250.Value = Round(ERL_NEBB(TerminationType, "250", Me.txtArea.Value), 1)
        Me.txt500.Value = Round(ERL_NEBB(TerminationType, "500", Me.txtArea.Value), 1)
        Me.txt1k.Value = Round(ERL_NEBB(TerminationType, "1k", Me.txtArea.Value), 1)
        Me.txt2k.Value = Round(ERL_NEBB(TerminationType, "2k", Me.txtArea.Value), 1)
        Me.txt4k.Value = Round(ERL_NEBB(TerminationType, "4k", Me.txtArea.Value), 1)
        Me.txt8k.Value = Round(ERL_NEBB(TerminationType, "8k", Me.txtArea.Value), 1)
        Else
        msg = MsgBox("Error - No type selected", vbOKOnly, "What did you do?")
        End If
    Else
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

Private Sub optRectangular_Click()
Me.txtL.Enabled = True
Me.txtW.Enabled = True
Me.txtDia.Enabled = False
PreviewERL
End Sub

Private Sub txtDia_Change()
PreviewERL
End Sub

Private Sub txtL_Change()
PreviewERL
End Sub

Private Sub txtW_Change()
PreviewERL
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewERL
End Sub

