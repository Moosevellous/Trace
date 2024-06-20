VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmElbowBend 
   Caption         =   "Elbow/Bend Loss"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   OleObjectBlob   =   "frmElbowBend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmElbowBend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#elbow--bend")
End Sub

Private Sub optASHRAE_Click()
EnableButtons
PreviewElbowLoss
End Sub

Private Sub optLined_Click()
PreviewElbowLoss
End Sub

Private Sub optNoVanes_Click()
PreviewElbowLoss
End Sub

Private Sub optRadius_Click()
Me.optLined.Enabled = False
Me.optUnlined.Enabled = False
Me.optVanes.Enabled = False
Me.optNoVanes.Enabled = False
PreviewElbowLoss
End Sub

Private Sub optSquare_Click()
Me.optLined.Enabled = True
Me.optUnlined.Enabled = True
Me.optVanes.Enabled = True
Me.optNoVanes.Enabled = True
PreviewElbowLoss
End Sub

Private Sub optSRL_Click()
EnableButtons
PreviewElbowLoss
End Sub

Private Sub optUnlined_Click()
PreviewElbowLoss
End Sub

Private Sub optVanes_Click()
PreviewElbowLoss
End Sub

Private Sub txtW_Change()
PreviewElbowLoss
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
PreviewElbowLoss
End Sub


Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnOK_Click()
'set public variables
ductW = txtW.Value
    
    If Me.optLined.Value Then
    elbowLining = "Lined"
    Else
    elbowLining = "Unlined"
    End If

    If Me.optSquare.Value Then
    elbowShape = "Square"
    Else
    elbowShape = "Radius"
    End If
    
    If Me.optVanes.Value Then
    elbowVanes = "Vanes"
    Else
    elbowVanes = "No Vanes"
    End If
    
    If Me.chkCalcRegen.Value = True Then
    CalcRegen = True
    Else
    CalcRegen = False
    End If
    
    If Me.optSRL.Value = True Then
    ductMethod = "SRL"
    Else 'default to ASHRAE
    ductMethod = "ASHRAE"
    End If
    
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Sub PreviewElbowLoss()
Dim Lining As String
Dim Vanes As String
Dim Shape As String

    If Me.optSquare.Value Then
    Shape = "Square"
    Else
    Shape = "Radius"
    End If
    
    If Me.optVanes.Value Then
    Vanes = "Vanes"
    Else
    Vanes = "No Vanes"
    End If

    If Me.optLined.Value Then
    Lining = "Lined"
    Else
    Lining = "Unlined"
    End If
    
    If Me.txtW.Value <> "" And IsNumeric(Me.txtW.Value) Then
    

        
        'PREVIEW VALUES
        If Me.optASHRAE.Value = True Then
            
        Me.txt31.Value = ElbowLoss_ASHRAE("31.5", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt63.Value = ElbowLoss_ASHRAE("63", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt125.Value = ElbowLoss_ASHRAE("125", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt250.Value = ElbowLoss_ASHRAE("250", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt500.Value = ElbowLoss_ASHRAE("500", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt1k.Value = ElbowLoss_ASHRAE("1k", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt2k.Value = ElbowLoss_ASHRAE("2k", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt4k.Value = ElbowLoss_ASHRAE("4k", Me.txtW.Value, Shape, Lining, Vanes)
        Me.txt8k.Value = ElbowLoss_ASHRAE("8k", Me.txtW.Value, Shape, Lining, Vanes)
        
        ElseIf Me.optSRL.Value = True Then
            
        Me.txt31.Value = "-"
        Me.txt63.Value = DuctBendAtten_SRL("63", Me.txtW.Value, Lining)
        Me.txt125.Value = DuctBendAtten_SRL("125", Me.txtW.Value, Lining)
        Me.txt250.Value = DuctBendAtten_SRL("250", Me.txtW.Value, Lining)
        Me.txt500.Value = DuctBendAtten_SRL("500", Me.txtW.Value, Lining)
        Me.txt1k.Value = DuctBendAtten_SRL("1k", Me.txtW.Value, Lining)
        Me.txt2k.Value = DuctBendAtten_SRL("2k", Me.txtW.Value, Lining)
        Me.txt4k.Value = DuctBendAtten_SRL("4k", Me.txtW.Value, Lining)
        Me.txt8k.Value = "-"
        Me.txtTableHead.Value = TXT_HEAD
        Me.txtRawVal.Value = TXT_RAW
        
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
        Me.txtTableHead.Value = "-"
        Me.txtRawVal.Value = "-"
        End If
    
    End If
    
End Sub

Sub EnableButtons()

    If Me.optSRL.Value = True Then
    Me.optVanes.Enabled = False
    Me.optNoVanes.Enabled = False
    Me.optRadius.Enabled = False
    Me.lblTableValues.Visible = True
    Me.txtTableHead.Visible = True
    Me.txtRawVal.Visible = True
    Else 'default to ASHRAE
    Me.optVanes.Enabled = True
    Me.optNoVanes.Enabled = True
    Me.optRadius.Enabled = True
    Me.lblTableValues.Visible = False
    Me.txtTableHead.Visible = False
    Me.txtRawVal.Visible = False
    End If

End Sub
