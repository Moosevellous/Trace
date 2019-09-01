VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLouvres 
   Caption         =   "Louvre Insertion Loss"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "frmLouvres.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLouvres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IL63() As Double
Dim IL125() As Double
Dim IL250() As Double
Dim IL500() As Double
Dim IL1k() As Double
Dim IL2k() As Double
Dim IL4k() As Double
Dim IL8k() As Double
Dim FA() As String 'unlike silencer code, this is a string
Dim Length() As Double
Dim L_Series() As String
Dim LouvreArray(0 To 280, 0 To 11) As Double
Dim LouvreNameArray(0 To 280) As String
Dim TextFileScanned As Boolean


Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub btnInsert_Click()
LouvreModel = Me.lstOptions.Value
ReDim LouvreIL(8) 'do not preserve
LouvreIL(0) = Me.txt63.Value
LouvreIL(1) = Me.txt125.Value
LouvreIL(2) = Me.txt250.Value
LouvreIL(3) = Me.txt500.Value
LouvreIL(4) = Me.txt1k.Value
LouvreIL(5) = Me.txt2k.Value
LouvreIL(6) = Me.txt4k.Value
LouvreIL(7) = Me.txt8k.Value
LouvreLength = Me.txtLength.Value
LouvreFA = Me.txtFA.Value
LouvreSeries = Me.txtSeries.Value 'this is the public value in Noise module
btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub lstOptions_Click()

Dim i As Integer
i = Me.lstOptions.ListIndex 'zero index, just like arrays
Me.txt63.Value = IL63(i)
Me.txt125.Value = IL125(i)
Me.txt250.Value = IL250(i)
Me.txt500.Value = IL500(i)
Me.txt1k.Value = IL1k(i)
Me.txt2k.Value = IL2k(i)
Me.txt4k.Value = IL4k(i)
Me.txt8k.Value = IL8k(i)
Me.txtFA.Value = FA(i)
Me.txtLength.Value = Length(i)
Me.txtSeries.Value = L_Series(i)

End Sub

Private Sub UserForm_Activate()
GetSettings
Me.lstOptions.Clear
ScanLouvreList

    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

End Sub


Sub ScanLouvreList()
Dim i As Integer
Dim j As Integer
Dim ReadStr() As String
    Open ACOUSTIC_LOUVRES For Input As #1
        i = 0 '<-line number
        found = False
        Do Until EOF(1) Or found = True
        ReDim Preserve ReadStr(i)
        Line Input #1, ReadStr(i)
        'Debug.Print ReadStr(i)
        Application.StatusBar = "Importing: " & ReadStr(i)
        SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
            If Left(SplitStr(0), 1) <> "*" Then
                For j = 2 To 10 'hard coded columns for FantechSilencers
                'Debug.Print splitStr(j)
                    If SplitStr(j) <> "" And IsNumeric(SplitStr(j)) Then
                    LouvreArray(i, j) = CDbl(SplitStr(j))
                    End If
                Next j
                LouvreNameArray(i) = SplitStr(0) 'first column has name of silencer
                Me.lstOptions.AddItem (SplitStr(0))
                LouvreArray(i, 1) = SplitStr(1) 'length of silencer
                
                'make IL arrays the size of the list
                ResizeArray (Me.lstOptions.ListCount)
                
                'Length
                Length(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(1))
                
                'Insertion Losses
                IL63(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(2))
                IL125(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(3))
                IL250(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(4))
                IL500(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(5))
                IL1k(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(6))
                IL2k(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(7))
                IL4k(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(8))
                IL8k(Me.lstOptions.ListCount - 1) = ScreenInput(SplitStr(9))
                
                    'Free area
                    If SplitStr(10) <> "" And SplitStr(10) <> "-" Then
                    FA(Me.lstOptions.ListCount - 1) = SplitStr(10)
                    Else
                    FA(Me.lstOptions.ListCount - 1) = ""
                    End If
                    
                    'Series
                    If SplitStr(11) <> "" And SplitStr(11) <> "-" Then
                    L_Series(Me.lstOptions.ListCount - 1) = SplitStr(12) & " " & SplitStr(11)
                    Else
                    L_Series(Me.lstOptions.ListCount - 1) = ""
                    End If
                    
            End If
        i = i + 1
        Loop
    Close #1
Application.StatusBar = False
End Sub

Private Sub ResizeArray(size As Integer)
ReDim Preserve IL63(size)
ReDim Preserve IL125(size)
ReDim Preserve IL250(size)
ReDim Preserve IL500(size)
ReDim Preserve IL1k(size)
ReDim Preserve IL2k(size)
ReDim Preserve IL4k(size)
ReDim Preserve IL8k(size)
ReDim Preserve FA(size)
ReDim Preserve Length(size)
ReDim Preserve L_Series(size)
End Sub

