VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlotTool 
   Caption         =   "Plot Tool"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "frmPlotTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlotTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''
'INITIALISE
'''''''''''''''

Dim FirstRun As Boolean 'switch for initialising form, set to false after initial setup
Dim ShowAdvancedSettings As Boolean



Function MarkerListBoxIndex(MarkerType As Long)
Dim SelectedValue As Long
'Dim MarkerFormatList() As Long

MarkerFormatList = Array(1, 2, 3, 4, 5, 8, 8, 10, 7, 6)

SelectedValue = -1 'catch errors
    Select Case MarkerType
    Case 1 'square
    SelectedValue = 1
    Case 2 'diamond
    SelectedValue = 2
    Case 3 'triangle
    SelectedValue = 3
    Case -4168 'square markers with X
    SelectedValue = 4
    Case 5 'square with asterisk
    SelectedValue = 5
    Case 8 'circle
    SelectedValue = 8
    Case 9 'square with plus
    SelectedValue = 8
'    Case 10
'        If SelectedValue = UBound(MarkerFormatList) Then 'next one will fall off
'        SelectedValue = 0
'        Else
'        SelectedValue = SelectedValue + 1
'        End If
    Case -4105 'automatic
    SelectedValue = 10
    Case -4115 'long bar
    SelectedValue = 7
    Case -4118 'short bar markers
    SelectedValue = 6
    Case -4142 'none
    SelectedValue = 0
    Case -4147 'picture markers
    SelectedValue = -1

    End Select
MarkerListBoxIndex = SelectedValue - 1 'stupid list indices
End Function

'Me.cBoxMarkerStyle.AddItem ("1 - Square")
'Me.cBoxMarkerStyle.AddItem ("2 - Diamond")
'Me.cBoxMarkerStyle.AddItem ("3 - Triangle")
'Me.cBoxMarkerStyle.AddItem ("4 - Cross")
'Me.cBoxMarkerStyle.AddItem ("5 - Asterisk")
'Me.cBoxMarkerStyle.AddItem ("6 - Dash")
'Me.cBoxMarkerStyle.AddItem ("7 - Big Dash")
'Me.cBoxMarkerStyle.AddItem ("8 - Circle")
'Me.cBoxMarkerStyle.AddItem ("9 - Plus")


Function IsSeriesNameActivated(SeriesName As String)
Dim i As Integer
IsSeriesNameActivated = False ' Initialize the function to return False

' Loop through each item in the ListBox
For i = 0 To Me.lstSeries.ListCount - 1
    If Me.lstSeries.List(i) = SeriesName And Me.lstSeries.Selected(i) Then
        IsSeriesNameActivated = True ' Set to True if SeriesName is found
        Exit Function ' Exit the function early as the SeriesName is found
    End If
Next i

End Function


Sub GetCurrentChartValues()

Dim YAxisStandardValues As Boolean
Dim XAxisStandardValues As Boolean

'catch error
If ActiveChart Is Nothing Then
    msg = MsgBox("No chart selected", vbOKOnly, "Error")
    End
End If

FirstRun = True

'Debug.Print ActiveChart.ChartType
    
'only certain chart types allowed
If ActiveChart.ChartType <> xlLineMarkers And ActiveChart.ChartType <> xlLine Then
    msg = MsgBox("Chart Type: " & ActiveChart.ChartType & chr(10) & "Sorry, only line charts allowed.", vbOKOnly, "Error")
    End
End If

'markers
If ActiveChart.FullSeriesCollection(1).MarkerStyle = -4142 Then 'no markers
    Me.chkShowMarkers.Value = False
Else
'Me.cBoxMarkerStyle.ListIndex = ActiveChart.FullSeriesCollection(1).MarkerStyle - 1 'list item number 1 is marker style 0?
    Me.chkShowMarkers.Value = True
    Me.cBoxMarkerStyle.ListIndex = MarkerListBoxIndex(ActiveChart.FullSeriesCollection(1).MarkerStyle) 'TODO: will have to update this for multi-markers
    Me.txtMarkerSize.Value = ActiveChart.FullSeriesCollection(1).MarkerSize

End If

'lines
If ActiveChart.FullSeriesCollection(1).Format.Line.Visible = msoFalse Then 'no line
    Me.txtLineThickness.Value = ""
    Me.chkShowLines.Value = False
Else
    Me.txtLineThickness.Value = ActiveChart.FullSeriesCollection(1).Format.Line.Weight
    Me.txtTransparency.Value = ActiveChart.FullSeriesCollection(1).Format.Line.Transparency * 100
End If
    
    
    'y-axis
With ActiveChart.Axes(xlValue, xlPrimary)
    If .HasTitle Then
    
        If InStr(1, .AxisTitle.text, "Sound Pressure Level", vbTextCompare) > 0 Then
            Me.optSPL.Value = True
            YAxisStandardValues = True
        End If
        
        If InStr(1, .AxisTitle.text, "Sound Power Level", vbTextCompare) > 0 Then
            Me.optSWL.Value = True
            YAxisStandardValues = True
        End If
        
        If InStr(1, .AxisTitle.text, "Transmission Loss", vbTextCompare) > 0 Then
            Me.optTL.Value = True
            YAxisStandardValues = True
        End If
        
        If InStr(1, .AxisTitle.text, "Insertion Loss", vbTextCompare) > 0 Then
            Me.optIL.Value = True
            YAxisStandardValues = True
        End If
        
        If YAxisStandardValues = False Then
            Me.optYOther = True
        End If
        
    Me.txtYAxis.Value = .AxisTitle.text
    End If
End With
    
With ActiveChart.Axes(xlValue)
    'major y-axis gridlines
    If .HasMajorGridlines = True Then
        Me.chkMajor.Value = True
        If .MajorUnit = 10 Then
            Me.optMajor10.Value = True
        ElseIf .MajorUnit = 5 Then
            Me.optMajor5.Value = True
        ElseIf .MajorUnit = 1 Then
            Me.optMajor1.Value = True
        Else
            Me.optMajorOther.Value = True
            Me.txtMajorGridValue = .MajorUnit
        End If
    Else
        Me.chkMajor.Value = False
    End If

        
        'minor y-axis gridlines
    If .HasMinorGridlines = True Then
        Me.chkMinor.Value = True
        If .MinorUnit = 10 Then
            Me.optMinor10.Value = True
        ElseIf .MinorUnit = 5 Then
            Me.optMinor5.Value = True
        ElseIf .MinorUnit = 1 Then
            Me.optMinor1.Value = True
        Else
            Me.optMinorOther.Value = True
            Me.txtMinorGridValue = .MinorUnit
        End If
    Else
        Me.chkMinor.Value = False
    End If
        
    'Y-axis number format
    numFormat = .TickLabels.NumberFormat
    numDecimal = Split(numFormat, ".", Len(numFormat), vbTextCompare)
        If UBound(numDecimal) >= 1 Then
            Me.txtDecimal.Value = Len(numDecimal(1))
        Else
            Me.txtDecimal.Value = 0
        End If
        
    'get Y-axis ranges
    Me.txtYRangeMax.Value = .MaximumScale
    Me.txtYRangeMin.Value = .MinimumScale
    'Me.txtYRangeMin.Value = .MaximumScale - 60 'Impose a 60dB range by default
        
    'check if Y-axis is log scale
    If .ScaleType = xlScaleLogarithmic Then Me.chkLogScale.Value = True
        
End With
    
'x-axis
With ActiveChart.Axes(xlCategory, xlPrimary)
    If .HasTitle Then
            
        If InStr(1, .AxisTitle.text, "Octave Band Centre Frequency, Hz", vbTextCompare) > 0 Then
            Me.optOct.Value = True
            XAxisStandardValues = True
        End If
            
        If InStr(1, .AxisTitle.text, "One-Third Octave Band Centre Frequency, Hz", vbTextCompare) > 0 Then
            Me.optOToct.Value = True
            XAxisStandardValues = True
        End If
            
        If XAxisStandardValues = False Then
            Me.txtXaxis.Enabled = True
        End If
        Me.txtXaxis.Value = .AxisTitle.text
    End If
    'ticks, on values by default
    .AxisBetweenCategories = False
    .MajorTickMark = xlInside
End With
    
    'Chart title
If ActiveChart.HasTitle = True Then
    Me.txtChartTitle.Value = ActiveChart.ChartTitle.text
    Me.chkChartTitle.Value = True
End If
    
    'legend
If ActiveChart.HasLegend Then
    If ActiveChart.Legend.Position = xlLegendPositionBottom Then
        Me.optLegendBottom.Value = True
    ElseIf ActiveChart.Legend.Position = xlLegendPositionRight Then
        Me.optLegendRight.Value = True
    End If
Else
    Me.optLegendNone.Value = True
End If
    
'lstSeries
For i = 1 To ActiveChart.FullSeriesCollection.Count
    Me.lstSeries.AddItem (ActiveChart.FullSeriesCollection(i).Name)
    Me.lstSeries.Selected(i - 1) = True
Next i

FirstRun = False 'end of setup
    
End Sub

'show the advanced panel
Private Sub btnAdvanced_Click()
If ShowAdvancedSettings = False Then 'show
Me.width = 780
Me.FrameSeries.Visible = True
Me.btnAdvanced.Caption = "<<<Hide"
ShowAdvancedSettings = True
Else 'hide
Me.width = 575
Me.FrameSeries.Visible = False
Me.btnAdvanced.Caption = "Advanced>>>"
ShowAdvancedSettings = False
End If

End Sub

Private Sub btnAllSeries_Click()
Dim r As Integer

For r = 0 To Me.lstSeries.ListCount - 1
    Me.lstSeries.Selected(r) = True
Next r

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub btnDecMinus1_Click()
Me.txtDecimal.Value = Me.txtDecimal.Value - 1
End Sub

Private Sub btnDecPlus1_Click()
Me.txtDecimal.Value = Me.txtDecimal.Value + 1
End Sub

Private Sub btnDone_Click()
'Me.Hide
Unload Me
End Sub

Private Sub btnFindAndReplace_Click()
Dim FindText As String
Dim ReplaceText As String
Dim OldFormula As String
Dim NewFormula As String

On Error GoTo catch:

FindText = Me.txtFindText.Value
ReplaceText = Me.txtReplaceText.Value

    With ActiveChart
        For i = 1 To .SeriesCollection.Count
        OldFormula = .SeriesCollection(i).Formula
        'Debug.Print OldFormula
        NewFormula = Replace(OldFormula, FindText, ReplaceText, 1, Len(OldFormula), vbTextCompare)
        'Debug.Print NewFormula
        .SeriesCollection(i).Formula = NewFormula
        Next i
    End With

'update the scan of the formulas
btnScanSeries_Click

catch:
    If Err.Number = 1004 Then
    ActiveChart.SeriesCollection(i).Formula = OldFormula
    End If

End Sub


Private Sub btnHelp_Click()
GotoWikiPage ("Sheet-Functions#plot")
End Sub

Private Sub btnMakeDashLines_Click()

ActiveChart.ChartArea.Select
    For i = 1 To ActiveChart.FullSeriesCollection.Count
    '.Select
        With ActiveChart.FullSeriesCollection(i).Format.Line
            If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
                Select Case i Mod 7 '7 options of dash
                Case Is = 0
                .DashStyle = msoLineLongDashDot
                Case Is = 1 'note this one happens first, because i=1 and 1 mod 7 is 1
                .DashStyle = msoLineSolid
                Case Is = 2
                .DashStyle = msoLineSysDot
                Case Is = 3
                .DashStyle = msoLineSysDash
                Case Is = 4
                .DashStyle = msoLineDash
                Case Is = 5
                .DashStyle = msoLineDashDot
                Case Is = 6
                .DashStyle = msoLineLongDash
                End Select
            End If
        End With
     Next i

End Sub

Private Sub btnMakeSolidLines_Click()
ActiveChart.ChartArea.Select
    For i = 1 To ActiveChart.FullSeriesCollection.Count
        ActiveChart.FullSeriesCollection(i).Format.Line.DashStyle = msoLineSolid
     Next i
End Sub

Private Sub btnMarkMinus1_Click()
Me.txtMarkerSize.Value = Me.txtMarkerSize.Value - 1
ApplyMarkerSize
End Sub

Private Sub btnMarkPlus1_Click()
    If Me.txtMarkerSize.Value = "" Then
    Me.txtMarkerSize.Value = 0
    Else
    Me.txtMarkerSize.Value = Me.txtMarkerSize.Value + 1
    End If
ApplyMarkerSize
End Sub

Private Sub btnNoneSeries_Click()
Dim r As Integer

For r = 0 To Me.lstSeries.ListCount - 1
    Me.lstSeries.Selected(r) = False
Next r
End Sub

Private Sub btnOpaque_Click()
Me.txtTransparency.Value = 0
End Sub

Private Sub btnScanSeries_Click()
Dim PrintStr As String
PrintStr = ""
With ActiveChart
    For i = 1 To .SeriesCollection.Count
        PrintStr = PrintStr & .SeriesCollection(i).Formula & chr(10) & chr(10)
    Next i
    Me.txtScanResults.Value = PrintStr
End With
Me.Repaint
End Sub

Private Sub btnThickMinusHalf_Click()
Me.txtLineThickness.Value = Me.txtLineThickness.Value - 0.5
ApplyLineWeight (Me.txtLineThickness.Value)
End Sub

Private Sub btnThickPlusHalf_Click()
Me.txtLineThickness.Value = Me.txtLineThickness.Value + 0.5
ApplyLineWeight (Me.txtLineThickness.Value)
End Sub

Private Sub btnTrans25_Click()
Me.txtTransparency.Value = 25
End Sub

Private Sub btnTrans50_Click()
Me.txtTransparency.Value = 50
End Sub

Private Sub btnYAxisNoDec_Click()
Me.txtDecimal.Value = 0
End Sub

Private Sub btnYMinus5_Click()
If Me.chkLogScale.Value = False Then
    Me.txtYRangeMax.Value = Me.txtYRangeMax.Value - 5
    Me.txtYRangeMin.Value = Me.txtYRangeMin.Value - 5
Else 'true: log scale is on
    Me.txtYRangeMin.Value = Me.txtYRangeMin.Value / 10
    Me.txtYRangeMax.Value = Me.txtYRangeMax.Value / 10
End If
End Sub

Private Sub btnYPlus5_Click()
If Me.chkLogScale.Value = False Then
    Me.txtYRangeMax.Value = Me.txtYRangeMax.Value + 5
    Me.txtYRangeMin.Value = Me.txtYRangeMin.Value + 5
Else 'true: log scale is on
    Me.txtYRangeMin.Value = Me.txtYRangeMin.Value * 10
    Me.txtYRangeMax.Value = Me.txtYRangeMax.Value * 10
End If
End Sub

Private Sub cBoxLineColours_Click()

ActiveChart.ClearToMatchColorStyle

    Select Case Me.cBoxLineColours.Value
    Case Is = "Blue"
    ActiveChart.ChartColor = 14 'blue
    Case Is = "Orange"
    ActiveChart.ChartColor = 15 'orange
    Case Is = "Yellow"
    ActiveChart.ChartColor = 17 'yellow
    Case Is = "Green"
    ActiveChart.ChartColor = 19 'green
    Case Is = "Grey"
    ActiveChart.ChartColor = 23 'grey
    Case Is = "Default (rainbow)"
    ActiveChart.ChartColor = 10 'RAINBOOOOOW MUTHAFUCKAAAAAAAAAA
    End Select
    
End Sub


Private Sub cBoxMarkerStyle_Click()
ApplyMarkerStyle
End Sub

Private Sub chkChartTitle_Click()
    If Me.chkChartTitle.Value = True Then
    Me.txtChartTitle.Enabled = True
    ApplyChartTitle
    Else
    Me.txtChartTitle.Enabled = False
    ActiveChart.ChartTitle.Delete
    End If
End Sub

Private Sub chkLogScale_Click()

    'turn off/on gridline controls
    If Me.chkLogScale.Value = True Then
    
    Me.chkMajor.Enabled = False
    Me.optMajor10.Enabled = False
    Me.optMajor5.Enabled = False
    Me.optMajor1.Enabled = False
    Me.txtMajorGridValue.Enabled = False
    
    Me.chkMinor.Enabled = False
    Me.optMinor10.Enabled = False
    Me.optMinor5.Enabled = False
    Me.optMinor1.Enabled = False
    Me.txtMinorGridValue.Enabled = False
    
    Me.btnYMinus5.Caption = "/ 10"
    Me.btnYPlus5.Caption = "X 10"
    
    Else 'false
    Me.chkMajor.Enabled = True
    Me.optMajor10.Enabled = True
    Me.optMajor5.Enabled = True
    Me.optMajor1.Enabled = True
    Me.txtMajorGridValue.Enabled = True
    
    Me.chkMinor.Enabled = True
    Me.optMinor10.Enabled = True
    Me.optMinor5.Enabled = True
    Me.optMinor1.Enabled = True
    Me.txtMinorGridValue.Enabled = True
    
    Me.btnYMinus5.Caption = "-5dB"
    Me.btnYPlus5.Caption = "+5dB"
    
    End If

'set scales
With ActiveChart.Axes(xlValue, xlPrimary)

    If Me.chkLogScale.Value = True Then
        .ScaleType = xlScaleLogarithmic
    Else
        .ScaleType = xlScaleLinear
    End If
    
    Me.txtYRangeMax.Value = .MaximumScale
    Me.txtYRangeMin.Value = .MinimumScale
    
End With

End Sub

Private Sub chkMajor_Click()
ApplyMajorGridlines
End Sub

Private Sub chkMinor_Click()
ApplyMinorGridlines
End Sub

Private Sub chkShowMarkers_Click()
    If FirstRun Then
    Me.cBoxMarkerStyle.ListIndex = MarkerListBoxIndex(ActiveChart.FullSeriesCollection(1).MarkerStyle)
    End If
    If Len(Me.cBoxMarkerStyle.Value) = 0 Then
    Me.cBoxMarkerStyle.Value = "1 - Square" 'default to square
    End If
Me.txtMarkerSize.Value = ActiveChart.FullSeriesCollection(1).MarkerSize
ApplyMarkerStyle
ApplyMarkerSize
End Sub

Private Sub optdB_Click()
YaxisLabel
End Sub

Private Sub optdBA_Click()
YaxisLabel
End Sub

Private Sub optIL_Click()
Me.txtYAxis.Enabled = False
Me.optdB.Value = True 'defaults to dB
Me.optdB.Enabled = False
Me.optdBA.Enabled = False
YaxisLabel
End Sub

'legend''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optLegendNone_Click()
ActiveChart.Legend.Delete
End Sub

Private Sub optLegendBottom_Click()
ActiveChart.Legend.Position = xlLegendPositionBottom
End Sub

Private Sub optLegendRight_Click()
ActiveChart.Legend.Position = xlLegendPositionRight
End Sub

Private Sub optLegendTopRight_Click()
ActiveChart.Legend.Position = xlLegendPositionCorner
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub optMajor10_Click()
ApplyMajorGridlines
End Sub

Private Sub optMajor1_Click()
ApplyMajorGridlines
End Sub

Private Sub optMajor5_Click()
ApplyMajorGridlines
End Sub

Private Sub optMajorOther_Click()
ApplyMajorGridlines
End Sub

Private Sub optMarkerFill_Click()
ApplyMarkerFill
End Sub

Private Sub optMarkerHollow_Click()
ApplyMarkerFill
End Sub

Private Sub optMinor10_Click()
ApplyMinorGridlines
End Sub

Private Sub optMinor1_Click()
ApplyMinorGridlines
End Sub

Private Sub optMinor5_Click()
ApplyMinorGridlines
End Sub

Private Sub optMinorOther_Click()
ApplyMinorGridlines
End Sub

Private Sub optOct_Click()
Me.txtXaxis.Enabled = False
XaxisLabel
End Sub

Private Sub optOToct_Click()
Me.txtXaxis.Enabled = False
XaxisLabel
End Sub

Private Sub optSPL_Click()
Me.txtYAxis.Enabled = False
Me.optdB.Enabled = True
Me.optdBA.Enabled = True
YaxisLabel
End Sub

Private Sub optSWL_Click()
Me.txtYAxis.Enabled = False
Me.optdB.Enabled = True
Me.optdBA.Enabled = True
YaxisLabel
End Sub

Private Sub optTL_Click()
Me.txtYAxis.Enabled = False
Me.optdB.Value = True 'defaults to dB
Me.optdB.Enabled = False
Me.optdBA.Enabled = False
YaxisLabel
End Sub

Private Sub optXNone_Click()
Me.txtXaxis.Enabled = False
XaxisLabel
End Sub

Private Sub optXOther_Click()
Me.txtXaxis.Enabled = True
XaxisLabel
End Sub

Private Sub optYNone_Click()
Me.txtYAxis.Enabled = False
YaxisLabel
End Sub

Private Sub optYOther_Click()
Me.txtYAxis.Enabled = True
Me.optdB.Enabled = False
Me.optdBA.Enabled = False
YaxisLabel
End Sub

Private Sub txtChartTitle_Change()
ApplyChartTitle
End Sub

Private Sub txtDecimal_Change()

Dim FormatString As String

FormatString = "0."

    If Me.txtDecimal.Value <> "" Then
    
        For i = 0 To Me.txtDecimal.Value - 1
        FormatString = FormatString & "0"
        Next i
        
        If Right(FormatString, 1) = "." Then '0 decimals
        FormatString = "0"
        End If
        
    ActiveChart.Axes(xlValue).TickLabels.NumberFormat = FormatString
    End If
    
End Sub

Private Sub txtLineThickness_Click()
ApplyLineWeight (Me.txtLineThickness.Value)
End Sub

Private Sub txtLineThickness_Change()
ApplyLineWeight (Me.txtLineThickness.Value)
End Sub

Private Sub txtMajorGridValue_Change()
Me.optMajorOther.Value = True
ApplyMajorGridlines
End Sub

Private Sub txtMarkerSize_Click()
ApplyMarkerSize
End Sub

Function IsVal(InValue As Variant)
IsVal = Application.WorksheetFunction.IsNumber(InValue)
End Function

Private Sub txtMinorGridValue_Change()
Me.optMinorOther.Value = True
ApplyMinorGridlines
End Sub

Private Sub txtTransparency_Change()
ApplyLineTransparency (Me.txtTransparency.Value)
End Sub

Private Sub txtXaxis_Change()
'Me.optXOther.Value = True
XaxisLabel
End Sub

Private Sub txtYAxis_Change()
'Me.optYOther.Value = True
YaxisLabel
End Sub

Private Sub txtYRangeMax_Change()
    If IsNumeric(Me.txtYRangeMax.Value) Then
    ActiveChart.Axes(xlValue).MaximumScale = Me.txtYRangeMax.Value
    End If
End Sub

Private Sub txtYRangeMin_Change()
    If IsNumeric(Me.txtYRangeMin.Value) Then
        If Me.chkLogScale.Value = True And txtYRangeMin.Value = 0 Then 'skip
        Else
        ActiveChart.Axes(xlValue).MinimumScale = Me.txtYRangeMin.Value
        End If
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False
'position in top corner
    With Me
    .Left = Application.Left + 20 '+ (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + 20 '+ (0.5 * Application.Height) - (0.5 * .Height)
    End With

PopulateMarkerComboBox

PopulateLineColourComboBox

GetCurrentChartValues

End Sub

'''''''''''''''''''''''''''''''''''
'Apply stuff to things and stuff
'''''''''''''''''''''''''''''''''''

Sub PopulateMarkerComboBox()
Me.cBoxMarkerStyle.Clear
Me.cBoxMarkerStyle.AddItem ("1 - Square")
Me.cBoxMarkerStyle.AddItem ("2 - Diamond")
Me.cBoxMarkerStyle.AddItem ("3 - Triangle")
Me.cBoxMarkerStyle.AddItem ("4 - Cross")
Me.cBoxMarkerStyle.AddItem ("5 - Asterisk")
Me.cBoxMarkerStyle.AddItem ("6 - Dash")
Me.cBoxMarkerStyle.AddItem ("7 - Big Dash")
Me.cBoxMarkerStyle.AddItem ("8 - Circle")
Me.cBoxMarkerStyle.AddItem ("9 - Plus")
Me.cBoxMarkerStyle.AddItem ("10 - Multi-marker")
End Sub

Sub PopulateLineColourComboBox()
Me.cBoxLineColours.Clear
Me.cBoxLineColours.AddItem ("Default (rainbow)")
Me.cBoxLineColours.AddItem ("Blue")
Me.cBoxLineColours.AddItem ("Orange")
Me.cBoxLineColours.AddItem ("Yellow")
Me.cBoxLineColours.AddItem ("Green")
Me.cBoxLineColours.AddItem ("Grey")
End Sub

Sub ApplyMarkerSize()

If Me.txtMarkerSize.Value <> "" Then
s = CInt(Me.txtMarkerSize.Value)
    If s <> "" And s > 1 And ActiveChart.FullSeriesCollection(1).MarkerStyle <> xlMarkerStyleNone Then
        ActiveChart.ChartArea.Select
        For i = 1 To ActiveChart.FullSeriesCollection.Count
            If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
                With ActiveChart.FullSeriesCollection(i)
                    .MarkerSize = s
                End With
            End If
        Next i
    End If
End If

End Sub

Sub ApplyMarkerStyle()
Dim getMarkerIndex() As String
Dim m_style As Integer
Dim MultiMarkerNo As Integer

MultiMarkerNo = 1
'split cBox into array, first thing is the number
getMarkerIndex = Split(Me.cBoxMarkerStyle.Value, " ", Len(Me.cBoxMarkerStyle.Value), vbTextCompare)

If UBound(getMarkerIndex) <> -1 Then 'nothing selected
    If Me.chkShowMarkers = True Then
        m_style = getMarkerIndex(0) 'first element
    Else
        m_style = 0
    End If
End If
    
If m_style = 10 Then 'multi-marker!
    'set sizes, loop through all series
    For i = 1 To ActiveChart.FullSeriesCollection.Count
        
        If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
            ActiveChart.FullSeriesCollection(i).MarkerStyle = MultiMarkerNo
        End If
        
        If MultiMarkerNo >= 9 Then
            MultiMarkerNo = 1 'loop around
        Else
            MultiMarkerNo = MultiMarkerNo + 1 'index up
        End If
        
    Next i
Else
    'set sizes, loop through all series
    For i = 1 To ActiveChart.FullSeriesCollection.Count
    Debug.Print ActiveChart.FullSeriesCollection(i).Name & " " & IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name)
        If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
            ActiveChart.FullSeriesCollection(i).MarkerStyle = m_style
        End If
    Next i
End If

End Sub

Sub ApplyMarkerFill()
Dim colorIndex As Integer
For i = 1 To ActiveChart.FullSeriesCollection.Count
    If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
        C = ActiveChart.FullSeriesCollection(i).Format.Line.ForeColor
        If Me.optMarkerFill.Value = True Then 'fill markers
            ActiveChart.FullSeriesCollection(i).MarkerForegroundColorIndex = 0
            ActiveChart.FullSeriesCollection(i).MarkerBackgroundColor = C
        Else 'hollow
            ActiveChart.FullSeriesCollection(i).MarkerForegroundColor = C
            ActiveChart.FullSeriesCollection(i).MarkerBackgroundColorIndex = 0
        End If
    End If
Next i
    
End Sub

Sub ApplyChartTitle()
    If ActiveChart.HasTitle = False Then
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    End If
ActiveChart.ChartTitle.text = Me.txtChartTitle.Value
End Sub

Sub MarkerBordersOff()
ActiveChart.ChartArea.Select
For i = 1 To ActiveChart.FullSeriesCollection.Count
    If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
        ActiveChart.FullSeriesCollection(i).MarkerForegroundColorIndex = xlColorIndexNone
    End If
Next i
End Sub

Sub ApplyLabels()
For i = 1 To ActiveChart.FullSeriesCollection.Count
    If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
        ActiveChart.FullSeriesCollection(i).Select
        ActiveChart.FullSeriesCollection(i).ApplyDataLabels
            With ActiveChart.FullSeriesCollection(i).DataLabels
            .ShowSeriesName = True
            .ShowValue = False
            .Position = xlLabelPositionAbove
            .Orientation = xlDownward
            .Format.TextFrame2.Orientation = msoTextOrientationDownward
            End With
    End If
Next i
End Sub

Sub ApplyLineWeight(s)

If IsNumeric(s) And s > 0.5 Then
    ActiveChart.ChartArea.Select
    For i = 1 To ActiveChart.FullSeriesCollection.Count
        If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
            ActiveChart.FullSeriesCollection(i).Format.Line.Weight = s
        End If
    Next i
End If

End Sub

Sub ApplyLineTransparency(tVal)
    If tVal <> "" And tVal <= 100 And tVal >= 0 Then
        For i = 1 To ActiveChart.FullSeriesCollection.Count
            If IsSeriesNameActivated(ActiveChart.FullSeriesCollection(i).Name) Then
                With ActiveChart.FullSeriesCollection(i)
                .Format.Line.DashStyle = 1
                .Format.Line.Transparency = tVal / 100
                End With
            End If
        Next i
    End If
End Sub


Sub ApplyMinorGridlines()
    If ActiveChart.Axes(xlValue).ScaleType <> xlScaleLogarithmic Then
        If Me.chkMinor.Value = True Then
        ActiveChart.Axes(xlValue).HasMinorGridlines = True
            If Me.optMinor10.Value = True Then
            ActiveChart.Axes(xlValue).MinorUnit = 10
            ElseIf Me.optMinor5.Value = True Then
            ActiveChart.Axes(xlValue).MinorUnit = 5
            ElseIf Me.optMinor1.Value = True Then
            ActiveChart.Axes(xlValue).MinorUnit = 1
            Else
                If IsNumeric(Me.txtMinorGridValue.Value) And Me.txtMinorGridValue.Value <> 0 Then
                ActiveChart.Axes(xlValue).MinorUnit = Me.txtMinorGridValue.Value
                End If
            End If
        Else
        ActiveChart.Axes(xlValue).HasMinorGridlines = False
        End If
    End If
End Sub

Sub ApplyMajorGridlines()
    If Me.chkMajor.Value = True Then
    ActiveChart.Axes(xlValue).HasMajorGridlines = True
        If Me.optMajor10.Value = True Then
        ActiveChart.Axes(xlValue).MajorUnit = 10
        ElseIf Me.optMajor5.Value = True Then
        ActiveChart.Axes(xlValue).MajorUnit = 5
        ElseIf Me.optMajor1.Value = True Then
        ActiveChart.Axes(xlValue).MajorUnit = 1
        Else
            If IsNumeric(Me.txtMajorGridValue.Value) And Me.txtMajorGridValue.Value <> 0 Then
            ActiveChart.Axes(xlValue).MajorUnit = Me.txtMajorGridValue.Value
            End If
        End If
    Else
    ActiveChart.Axes(xlValue).HasMajorGridlines = False
    End If
End Sub

Sub YaxisLabel()
'guard clause
If FirstRun = True Then Exit Sub

    If Me.optYNone.Value = False Then
        If Me.optYOther.Value = False Then 'only if not 'other'
            If Me.optSPL.Value = True Then
                If Me.optdB.Value = True Then
                Me.txtYAxis.text = "Sound Pressure Level, dB"
                Else
                Me.txtYAxis.text = "Sound Pressure Level, dBA"
                End If
            ElseIf Me.optSWL.Value Then
                If Me.optdB.Value = True Then
                Me.txtYAxis.text = "Sound Power Level, dB"
                Else
                Me.txtYAxis.text = "Sound Power Level, dBA"
                End If
            ElseIf Me.optTL.Value = True Then
            Me.txtYAxis.text = "Transmission Loss, dB"
            ElseIf Me.optIL.Value = True Then
            Me.txtYAxis.text = "Insertion Loss, dB"
            End If
        End If
        

        With ActiveChart.Axes(xlValue, xlPrimary)
            If .HasTitle = False Then
            .HasTitle = True
            End If
        .AxisTitle.text = Me.txtYAxis.text 'this line is problematic
        End With

    Else
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = False
    End If
        
End Sub

Sub XaxisLabel()
'guard clause
If FirstRun = True Then Exit Sub

If Me.optXNone.Value = False Then

    With ActiveChart.Axes(xlCategory, xlPrimary)
        If .HasTitle = False Then
        .HasTitle = True
        End If
    .AxisTitle.text = Me.txtXaxis.text
    End With
        
    'Put labels in textbox
    If Me.optXOther.Value = False Then 'only if not 'other'
        If Me.optOct.Value = True Then
        Me.txtXaxis.text = "Octave Band Centre Frequency, Hz"
        Else
        Me.txtXaxis.text = "One-Third Octave Band Centre Frequency, Hz"
        End If
    End If
        
Else
ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = False
End If

End Sub

