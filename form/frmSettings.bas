VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Trace Settings"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17835
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDone_Click()
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub btnGetInfo_Click()
Call GetSettings
'ClearText
PrintText "***************"
PrintText "Central folders"
PrintText "***************"
PrintText ROOTPATH
PrintText TEMPLATELOCATION
PrintText STANDARDCALCLOCATION
PrintText FIELDSHEETLOCATION
PrintText EQUIPMENTSHEETLOCATION
PrintText ""
PrintText "***************"
PrintText "Central text files...."
PrintText "***************"
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
PrintText ("***************")
PrintText ("Version Info")
PrintText ("***************")
PrintText "Application:", Application.AddIns("Trace").Application
PrintText "CLSID:", Application.AddIns("Trace").CLSID
PrintText "Creator:", Application.AddIns("Trace").Creator
PrintText "FullName", Application.AddIns("Trace").FullName
PrintText "Installed:", Application.AddIns("Trace").Installed
PrintText "Open:", Application.AddIns("Trace").IsOpen
PrintText "Name:", Application.AddIns("Trace").Name
PrintText "Parent:", Application.AddIns("Trace").Parent
PrintText "Path:", Application.AddIns("Trace").Path
PrintText "ID:", Application.AddIns("Trace").progID
PrintText ""
End Sub

Sub PrintText(InputStr As String, Optional InputStr2 As String)
Me.txtSettings.Text = Me.txtSettings.Text & chr(10) & InputStr & " " & InputStr2
Me.txtSettings.SelStart = Len(Me.txtSettings.Text) - 1
Me.txtSettings.SetFocus
End Sub

Sub ClearText()
Me.txtSettings.Value = ""
End Sub

Private Sub btnHelp_Click()
GotoWikiPage
End Sub
