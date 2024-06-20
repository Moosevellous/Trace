VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLouvreDirectivity 
   Caption         =   "Louvre Directivity (SRL Method)"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12675
   OleObjectBlob   =   "frmLouvreDirectivity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLouvreDirectivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Private Sub btnOK_Click()
btnOkPressed = True
Me.Hide
Unload Me
End Sub


Private Sub optWidth05m_Click()
SelectNewPicture
End Sub

Private Sub optWidth10m_Click()
SelectNewPicture
End Sub

Private Sub optWidth15m_Click()
SelectNewPicture
End Sub

Private Sub optWidth20m_Click()
SelectNewPicture
End Sub

Private Sub optWidth25m_Click()
SelectNewPicture
End Sub

Private Sub optWidth35m_Click()
SelectNewPicture
End Sub

Private Sub optWidth45m_Click()
SelectNewPicture
End Sub

Private Sub optWidth55m_Click()
SelectNewPicture
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub


Sub SelectNewPicture()

Dim ImagePath As String
Dim PathStr As String

GetSettings

ImagePath = "img\Louvre_" & ImagePath & SelectedWidth & ".jpg"

PathStr = ROOTPATH & "\" & ImagePath

    If Dir(PathStr, vbNormal) <> "" Then
    Me.imgPolarPlot.Picture = LoadPicture(ROOTPATH & "\" & filePath)
    End If

End Sub


Function SelectedWidth()
Dim i As Integer
    For i = 0 To Me.fraLouvreWidth.Controls.Count - 1
    If Me.fraLouvreWidth.Controls(i).Value = True Then
    SelectedWidth = Me.fraLouvreWidth.Controls(i).Caption
    Exit Function
    End If
    Next
End Function

