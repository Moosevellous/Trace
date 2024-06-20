VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Trace Settings"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOpenCode_Click()
Application.VBE.MainWindow.Visible = True
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
        .Left = Application.Left + (0.5 * Application.width) - (0.5 * .width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
btnGetSettings_Click
End Sub

Private Sub btnDone_Click()
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub btnGetInfo_Click()

Call GetSettings

'ClearText
PrintText "***"
PrintText "Central folders"
PrintText "***"
PrintText ROOTPATH
PrintText TRACELOGFOLDER
PrintText TEMPLATELOCATION
PrintText STANDARDCALCLOCATION
PrintText FIELDSHEETLOCATION
PrintText EQUIPMENTSHEETLOCATION
PrintText ""

PrintText "***"
PrintText "Central text files...."
PrintText "***"
PrintText TRACELOGFILE
PrintText ASHRAE_DUCT
PrintText ASHRAE_FLEX
PrintText ASHRAE_REGEN
PrintText FANTECH_SILENCERS
PrintText FANTECH_DUCTS
PrintText ACOUSTIC_LOUVRES
PrintText DUCT_DIRLOSS
PrintText ""

End Sub

Private Sub btnGetSettings_Click()

Call GetSettings

'ClearText
PrintText ("***")
PrintText ("Version Info")
PrintText ("***")

    'print all properties of the add-in
    With Application.AddIns("Trace")
    PrintText "Application:", .Application
    PrintText "CLSID:", .CLSID
    PrintText "Creator:", .Creator
    PrintText "FullName", .FullName
    PrintText "Installed:", .Installed
    PrintText "Open:", .IsOpen
    PrintText "Name:", .Name
    PrintText "Parent:", .Parent
    PrintText "Path:", .Path
    PrintText "ID:", .progID
    End With
    
PrintText ""

End Sub

Sub PrintText(inputStr As String, Optional InputStr2 As String)

    'catch asterisks
    If inputStr = "***" Then inputStr = "*******************************"
    
        
    With Me.txtSettings
    .text = Me.txtSettings.text & chr(10) & inputStr & " " & InputStr2
    .SelStart = Len(Me.txtSettings.text) - 1
    .SetFocus
    End With
End Sub

Sub ClearText()
Me.txtSettings.Value = ""
End Sub

Private Sub btnHelp_Click()
GotoWikiPage
End Sub
