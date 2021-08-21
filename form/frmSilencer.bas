VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSilencer 
   Caption         =   "Select Fantech Silencer"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14370
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
Dim Series() As String
Dim SilencerArray(0 To 1000, 0 To 11) As Double 'TODO: make dynamic array?
Dim SilNameArray() As String 'dynamic array
Dim SilSeriesArray() As String 'dynamic array
Dim TextFileScanned As Boolean
Dim numResults As Integer

Private Sub btnHelp_Click()
GotoWikiPage ("Mechanical#silencer")
End Sub

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
SilSeries = Me.txtSupplier.Value & " " & Me.txtSeries.Value 'global variable
btnOkPressed = True
Me.Hide

    'call regen
    If Me.chkCalcRegen.Value = True Then
    frmSilencerRegen.txtTypeCode.Value = SilencerModel
        'set switch on frmSilencerRegen
        If Me.txtSupplier.Value = "Fantech" Then
        frmSilencerRegen.optFantech.Value = True
        ElseIf Me.txtSupplier.Value = "NAP" Then
        frmSilencerRegen.optNAP.Value = True
        Else
        msg = MsgBox("Supplier not recognised. No regenerated noise curves exist for this element.", vbOKOnly, "Error")
        End If
    CalcRegen = True
    Else
    CalcRegen = False
    End If

End Sub

Private Sub btnSearch_Click()
Me.lblStatus.Caption = "Getting database file path..."
GetSettings
Me.lstOptions.Clear
Me.lblStatus.Caption = "Searching database..."
SearchSilencerName (Me.txtSearchName.Value)
Me.lblStatus.Caption = "Search complete: " & numResults & " results"
End Sub

Private Sub btnSearchAll_Click()
Me.optSearch.Value = True
Me.txtSearchName.Value = "<ALL>"
btnSearch_Click
End Sub

Private Sub btnSolve_Click()
Me.lblStatus.Caption = "Getting database file path..."
GetSettings
Me.lstOptions.Clear
Me.lblStatus.Caption = "Solving..."
SolveForSilencer Me.RefSilencerRange.Value, Me.RefTargetRange.Value, Me.optNR.Value, CDbl(Me.txtNoiseGoal.Value)
Me.lblStatus.Caption = "Search complete: " & numResults & " results"
End Sub

Private Sub lstOptions_Click()
Dim i As Integer
Dim SplitSeries() As String
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
SplitSeries = Split(Series(i), " ")
Me.txtSupplier.Value = SplitSeries(0)
Me.txtSeries.Value = SplitSeries(1)
End Sub

Private Sub optCircular_Click()
EnableCircularOptions
End Sub

Private Sub optRectangular_Click()
EnableRectangularOptions
End Sub

Private Sub optdBA_Click()
Me.lblUnits = "dBA"
End Sub

Private Sub optNR_Click()
Me.lblUnits = "NR"
End Sub

Private Sub optSearch_Click()
Me.txtSearchName.Enabled = True
Me.lstOptions.Clear
Me.lblStatus.Caption = ""
EnableFrame Me.FrameSolver, False
EnableFrame Me.FrameSearch, True
End Sub

Private Sub optSolver_Click()
Me.txtSearchName.Enabled = False
Me.lstOptions.Clear
Me.lblStatus.Caption = ""
EnableFrame Me.FrameSolver, True
EnableFrame Me.FrameSearch, False
End Sub

Private Sub txtSearchName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.optSearch.Value = True
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
Me.lstOptions.Clear
Me.RefSilencerRange.Value = ""
Me.RefTargetRange.Value = ""
TextFileScanned = False

End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
End Sub

Sub EnableCircularOptions()
'Me.optOpen.Enabled = True
'Me.optPod.Enabled = True
'Me.optStraight.Enabled = False
'Me.optTapered.Enabled = False
End Sub

Sub EnableRectangularOptions()
'Me.optOpen.Enabled = False
'Me.optPod.Enabled = False
'Me.optStraight.Enabled = True
'Me.optTapered.Enabled = True
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

If SearchStr = "" Then
GoTo catcherror
End If

Open FANTECH_SILENCERS For Input As #1

    i = 0 '<-line number
    found = False
    Do Until EOF(1) Or found = True
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
    
    If SearchStr = "<ALL>" Then 'catch wildcard
    MatchFound = True
    ElseIf InStr(1, SplitStr(11), SearchStr, vbTextCompare) > 0 Then '11th column is silencer name
    MatchFound = True
    Else
    MatchFound = False
    End If
    
        If MatchFound = True And InStr(1, SplitStr(0), "*", vbTextCompare) = 0 Then 'not rows with a star
        
        Me.lstOptions.AddItem (SplitStr(11))
        
        'make IL arrays the size of the list
        ResizeArray (Me.lstOptions.ListCount)
        
        'Debug.Print (splitStr(2))
        'check for 63Hz band missing
        If IsNumeric(SplitStr(2)) Then IL63(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(2))
        IL125(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(3))
        IL250(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(4))
        IL500(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(5))
        IL1k(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(6))
        IL2k(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(7))
        IL4k(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(8))
        IL8k(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(9))
        
            'Free area
            If SplitStr(10) <> "" Then
            FA(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(10))
            Else
            FA(Me.lstOptions.ListCount - 1) = 0
            End If
            
            'Length
            If SplitStr(1) <> "" Then
            Length(Me.lstOptions.ListCount - 1) = CDbl(SplitStr(1))
            Else
            Length(Me.lstOptions.ListCount - 1) = 0
            End If
            
            'series
            If SplitStr(12) <> "" Then
            Series(Me.lstOptions.ListCount - 1) = SplitStr(12)
            End If
            
        Else
        End If
    Loop

    If Me.lstOptions.ListCount > 0 Then
    Me.btnInsert.Enabled = True
    Else
    Me.btnInsert.Enabled = False
    End If
    
numResults = Me.lstOptions.ListCount

catcherror:
Close #1
End Function

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
ReDim Preserve Series(size)
End Sub


Function SolveForSilencer(SilRng As String, TargetRng As String, NRGoal As Boolean, NoiseGoal As Double)
Dim found As Boolean
Dim targetAddr() As String
Dim targetRw As Integer
Dim SilAddr() As String
Dim SilRw As Integer
Dim TestLevel As Double
Dim SilCols As Integer

targetAddr = Split(TargetRng, "$", Len(TargetRng), vbTextCompare) 'TODO error checking for row
SilAddr = Split(SilRng, "$", Len(SilRng), vbTextCompare)

    If UBound(targetAddr) <> -1 Or UBound(SilAddr) <> -1 Then
    
    targetRw = targetAddr(2)
    SilRw = SilAddr(2)
    'send to public variable
    SolverRow = SilRw
    
    Me.lblStatus.Caption = "Scanning database...."
        If TextFileScanned = False Then 'Scan text file with silencers
        ScanFantechTextFile
        TextFileScanned = True
        End If
    
Application.Calculation = xlCalculationManual

    'search for compliant silencers
        'place in cells
        For Rw = 2 To UBound(SilencerArray)
        SilCols = 0
            For Col = T_LossGainStart + 1 To T_LossGainEnd - 1
            'Debug.Print SilencerArray(rw, Col - 4)
            Cells(SilRw, Col).Value = SilencerArray(Rw, 2 + SilCols) 'start from element 2
            SilCols = SilCols + 1 'index write row
            Next Col
            
        'Debug.Print UBound(SilNameArray)
            If UBound(SilNameArray) >= Rw Then
            Cells(SilRw, T_Description).Value = SilNameArray(Rw)
            Me.lblStatus.Caption = "Checking: " & SilNameArray(Rw)
            Else
            Me.lblStatus.Caption = ""
            End If
            
        Calculate
        DoEvents
        
        If Me.optNR = True Then
        'TestLevel = Cells(targetRw, 14).Value
        TestLevel = NR_rate(Range(Cells(targetRw, T_LossGainStart), Cells(targetRw, T_LossGainStart)))
        Else
        TestLevel = Round(Cells(targetRw, T_LossGainStart - 1).Value, 1)
        End If
        
        If TestLevel <= NoiseGoal And TestLevel >= (NoiseGoal - CDbl(Me.txtDesignTolerance.Value)) Then 'silencer achieves target, but doesn't overshoot
        Me.lstOptions.AddItem (SilNameArray(Rw))
        ResizeArray (Me.lstOptions.ListCount)
        IL63(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 2)
        IL125(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 3)
        IL250(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 4)
        IL500(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 5)
        IL1k(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 6)
        IL2k(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 7)
        IL4k(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 8)
        IL8k(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 9)
        FA(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 10)
        Length(Me.lstOptions.ListCount - 1) = SilencerArray(Rw, 1)
        Series(Me.lstOptions.ListCount - 1) = SilSeriesArray(Rw)
        End If
        
        Next Rw
    End If 'ubound close loop
    
If Me.lstOptions.ListCount > 0 Then
Me.btnInsert.Enabled = True
Else
Me.btnInsert.Enabled = False
End If

Application.Calculation = xlCalculationAutomatic

numResults = Me.lstOptions.ListCount

End Function


Sub ScanFantechTextFile()
Dim i As Integer
Dim j As Integer
Dim ReadStr() As String

Open FANTECH_SILENCERS For Input As #1 'set Global variable to text input
i = 0 '<-line number
found = False

    Do Until EOF(1) Or found = True
    ReDim Preserve ReadStr(i)
    Line Input #1, ReadStr(i)
    'Debug.Print ReadStr(i)
    SplitStr = Split(ReadStr(i), vbTab, Len(ReadStr(i)), vbTextCompare)
    
        If Left(SplitStr(0), 1) <> "*" Then
            For j = 2 To 10 'hard coded columns for Silencers
            'Debug.Print splitStr(j)
                If SplitStr(j) <> "" And SplitStr(j) <> "-" Then
                'Debug.Print UBound(SilencerArray)
                SilencerArray(i, j) = CDbl(SplitStr(j))
                End If
            Next j
            
        'resize arrays for number of elements
        ReDim Preserve SilNameArray(i)
        ReDim Preserve SilSeriesArray(i)
        SilNameArray(i) = SplitStr(11) 'column 11 has name of silencer
        SilencerArray(i, 1) = SplitStr(1) 'length of silencer
        SilSeriesArray(i) = SplitStr(12) 'column 12 has silencer series (eg Fantech, NAP etc)
        End If
        
    i = i + 1
    Loop
    
Close #1
    
End Sub
