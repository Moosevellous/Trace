VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSilencer 
   Caption         =   "Select Fantech Silencer"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14775
   OleObjectBlob   =   "frmSilencer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSilencer"
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
Dim FA() As Double
Dim Length() As Double
Dim SilencerArray(0 To 280, 0 To 11) As Double
Dim SilNameArray(0 To 280) As String
Dim TextFileScanned As Boolean

Private Sub btnInsert_Click()
'SilencerModel is a public variable
SilencerModel = Me.lstOptions.Value
ReDim SilencerIL(8) 'do not preserve
SilencerIL(0) = Me.txt63.Value
SilencerIL(1) = Me.txt125.Value
SilencerIL(2) = Me.txt250.Value
SilencerIL(3) = Me.txt500.Value
SilencerIL(4) = Me.txt1k.Value
SilencerIL(5) = Me.txt2k.Value
SilencerIL(6) = Me.txt4k.Value
SilencerIL(7) = Me.txt8k.Value

SilLength = CDbl(Me.txtLength.Value)
SilFA = CDbl(Me.txtFA.Value)
btnOkPressed = True
Me.Hide
End Sub

Private Sub btnSearch_Click()
GetSettings
    If Me.optSearch.Value = True Then
    Me.lstOptions.Clear
    SearchSilencerName (Me.txtSearchName)
    ElseIf Me.optSolver = True Then
    Me.lstOptions.Clear
    'Call SolveForSilencer(Me.optRectangular, Me.optStraight, Me.optPod, Me.chkQseal.Value, Me.RefSilencerRange.Value, Me.optNR, CDbl(Me.txtNoiseGoal.Value))
    s = SolveForSilencer(Me.RefSilencerRange.Value, Me.RefTargetRange.Value, Me.optNR.Value, CDbl(Me.txtNoiseGoal.Value))
    Else
    msg = MsgBox("Run Error!", vbOKOnly, "HOW?!")
    End If
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
End Sub

Private Sub optCircular_Click()
EnableCircularOptions
End Sub

Private Sub optRectangular_Click()
EnableRectangularOptions
End Sub

Private Sub optDBA_Click()
Me.lblUnits = "dBA"
End Sub

Private Sub optNR_Click()
Me.lblUnits = "NR"
End Sub

Private Sub optSearch_Click()
Me.txtSearchName.Enabled = True
EnableFrame Me.FrameSolver, False
End Sub

Private Sub optSolver_Click()
Me.txtSearchName.Enabled = False
EnableFrame Me.FrameSolver, True
End Sub

Private Sub UserForm_Activate()
Me.lstOptions.Clear
Me.RefSilencerRange.Value = ""
Me.RefTargetRange.Value = ""
TextFileScanned = False
    With Me
        .Top = (Application.Height - Me.Height) / 2
        .Left = (Application.Width - Me.Width) / 2
    End With
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Sub EnableCircularOptions()
Me.optOpen.Enabled = True
Me.optPod.Enabled = True
Me.optStraight.Enabled = False
Me.optTapered.Enabled = False
End Sub

Sub EnableRectangularOptions()
Me.optOpen.Enabled = False
Me.optPod.Enabled = False
Me.optStraight.Enabled = True
Me.optTapered.Enabled = True
End Sub



Public Sub EnableFrame(InFrame As Frame, ByVal Flag As Boolean)
Dim Contrl As control
On Error Resume Next

InFrame.Enabled = Flag 'enable or disable the frame that passed as parameter.
'passing over all controls
    For Each Contrl In InFrame.Controls
        If (Contrl.Container.Name = InFrame.Name) Then
        Contrl.Enabled = Flag
        End If
        
        If Flag = True Then 'some radio buttons are not enabled
'            If Me.optCircular.Value = True Then
'            EnableCircularOptions
'            Else
'            EnableRectangularOptions
'            End If
        End If
        
    Next
End Sub

Function SearchSilencerName(SearchStr As String) As String()
Dim i As Integer
Dim found As Boolean
Dim ReadStr() As String

Open FANTECH_SILENCERS For Input As #1

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    splitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
        If InStr(1, splitStr(11), SearchStr, vbTextCompare) > 0 Then '11th column is silencer name
        
        Me.lstOptions.AddItem (splitStr(11))
        
        'make IL arrays the size of the list
        ResizeArray (Me.lstOptions.ListCount)
        
        'Debug.Print (splitStr(2))
        IL63(Me.lstOptions.ListCount - 1) = CDbl(splitStr(2))
        IL125(Me.lstOptions.ListCount - 1) = CDbl(splitStr(3))
        IL250(Me.lstOptions.ListCount - 1) = CDbl(splitStr(4))
        IL500(Me.lstOptions.ListCount - 1) = CDbl(splitStr(5))
        IL1k(Me.lstOptions.ListCount - 1) = CDbl(splitStr(6))
        IL2k(Me.lstOptions.ListCount - 1) = CDbl(splitStr(7))
        IL4k(Me.lstOptions.ListCount - 1) = CDbl(splitStr(8))
        IL8k(Me.lstOptions.ListCount - 1) = CDbl(splitStr(9))
        
            'Free area
            If splitStr(10) <> "" Then
            FA(Me.lstOptions.ListCount - 1) = CDbl(splitStr(10))
            Else
            FA(Me.lstOptions.ListCount - 1) = 0
            End If
            
            'Length
            If splitStr(2) <> "" Then
            Length(Me.lstOptions.ListCount - 1) = CDbl(splitStr(1))
            Else
            Length(Me.lstOptions.ListCount - 1) = 0
            End If
            
        'Debug.Print IL8k(Me.lstOptions.ListCount - 1)
        Else
        End If
    Loop

    If Me.lstOptions.ListCount > 0 Then
    Me.btnInsert.Enabled = True
    Else
    Me.btnInsert.Enabled = False
    End If

catcherror:
Close #1
End Function

Sub ResizeArray(size As Integer)
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
End Sub

Function SolveForSilencer(SilRng As String, targetRng As String, NRGoal As Boolean, NoiseGoal As Double) 'As String()         RectSil As Boolean, SilStraight As Boolean, SilPodded As Boolean, Qseal As Boolean,
Dim found As Boolean
Dim targetAddr() As String
Dim targetRw As Integer
Dim silAddr() As String
Dim silRw As Integer
Dim TestLevel As Double

targetAddr = Split(targetRng, "$", Len(targetRng), vbTextCompare) 'TODO error checking for row
silAddr = Split(SilRng, "$", Len(SilRng), vbTextCompare)

    If UBound(targetAddr) <> -1 Or UBound(silAddr) <> -1 Then
    
    targetRw = targetAddr(2)
    silRw = silAddr(2)
    'send to public variable
    SolverRow = silRw
    
        
        If TextFileScanned = False Then 'Scan text file with silencers
        ScanFantechTextFile
        TextFileScanned = True
        End If
    
Application.Calculation = xlCalculationManual

    'search for compliant silencers
        'place in cells
        For rw = 2 To UBound(SilencerArray)
            For Col = 6 To 13
            'Debug.Print SilencerArray(rw, Col - 4)
            Cells(silRw, Col).Value = SilencerArray(rw, Col - 4)
            Next Col
        Cells(silRw, 2).Value = SilNameArray(rw)
        Calculate
        DoEvents
        
        If Me.optNR = True Then
        'TestLevel = Cells(targetRw, 14).Value
        TestLevel = NR_rate(Range(Cells(targetRw, 5), Cells(targetRw, 13)))
        Else
        TestLevel = Round(Cells(targetRw, 4).Value, 1)
        End If
        
        If TestLevel <= NoiseGoal Then 'silencer achieves target
        Me.lstOptions.AddItem (SilNameArray(rw))
        ResizeArray (Me.lstOptions.ListCount)
        IL63(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 2)
        IL125(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 3)
        IL250(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 4)
        IL500(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 5)
        IL1k(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 6)
        IL2k(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 7)
        IL4k(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 8)
        IL8k(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 9)
        FA(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 10)
        Length(Me.lstOptions.ListCount - 1) = SilencerArray(rw, 1)
        End If
        
        Next rw
    End If 'ubound close loop
    
If Me.lstOptions.ListCount > 0 Then
Me.btnInsert.Enabled = True
Else
Me.btnInsert.Enabled = False
End If

Application.Calculation = xlCalculationAutomatic

'Calculate

End Function


Sub ScanFantechTextFile()
Dim i As Integer
Dim j As Integer
Dim ReadStr() As String
    Open FANTECH_SILENCERS For Input As #1
        i = 0 '<-line number
        found = False
        Do Until EOF(1) Or found = True
        ReDim Preserve ReadStr(i)
        Line Input #1, ReadStr(i)
        'Debug.Print ReadStr(i)
        splitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
            If Left(splitStr(0), 1) <> "*" Then
                For j = 2 To 10 'hard coded columns for FantechSilencers
                'Debug.Print splitStr(j)
                    If splitStr(j) <> "" Then
                    SilencerArray(i, j) = CDbl(splitStr(j))
                    End If
                Next j
                SilNameArray(i) = splitStr(11) 'first column has name of silencer
                SilencerArray(i, 1) = splitStr(1) 'length of silencer
            End If
        i = i + 1
        Loop
    Close #1
End Sub
