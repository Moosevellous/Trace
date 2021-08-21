VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbsDB 
   Caption         =   "Absorption Database"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16350
   OleObjectBlob   =   "frmAbsDB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbsDB"
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
GotoWikiPage
End Sub

Private Sub btnInsert_Click()
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub imgStar1_Enter()
UpdatePicture ("img\Star_filled.gif")
End Sub

'''''''''''''''''''''
Sub UpdatePicture(FilePath As String)
Dim objPic As Image
GetSettings
'set objpic =me.Controls(
    If TestLocation(ROOTPATH & "\" & FilePath) = True Then
    Me.imgStar1.Picture = LoadPicture(ROOTPATH & "\" & FilePath)
    End If
End Sub


Private Sub frameSearchProduct_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'Debug.Print "x"; X; "y"; Y
'check if is within one of the stars, then fill
    If x >= Me.imgStar1.Left And x <= Me.imgStar1.Left + Me.imgStar1.Width And _
    y >= Me.imgStar1.Top And x <= Me.imgStar1.Top + Me.imgStar1.Height Then
    UpdatePicture ("img/star_filled.gif")
    Else
    UpdatePicture ("img/star_hollow.gif")
    End If
End Sub

Private Sub imgStar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'click event
End Sub

Private Sub imgStar1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'Debug.Print "x"; X; "y"; Y

End Sub

'Private Sub UserForm_Click()
'Me.ListBox1.AddItem ("Autex")
'Me.ListBox1.AddItem ("IAC")
'Me.ListBox1.AddItem ("Joe abs")
'Me.ListBox1.AddItem ("something")
'End Sub
