VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScheduleBuilder 
   Caption         =   "Schedule Builder"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "frmScheduleBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmScheduleBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SearchForSymbol As String
Dim SearchGroupName As String
Dim HomeSheetName As String
Dim SkipCheck As Boolean

Private Sub btnHelp_Click()
GotoWikiPage "Sheet-Functions#schedule-builder"
End Sub

Private Sub chkCreateHeading_Click()
CountMarkersAllSheets
End Sub

Private Sub UserForm_Initialize()
Me.cBoxApplyStyle.AddItem ("None")
Me.cBoxApplyStyle.AddItem ("Normal")
Me.cBoxApplyStyle.AddItem ("Reference")
'set default
Me.cBoxApplyStyle.Value = "Reference"
End Sub

Private Sub btnCancel_Click()
btnOkPressed = False
Me.Hide
Unload Me
End Sub

Private Sub CountMarkersAllSheets()
Dim s As Integer
Dim FirstRow As Integer
Dim LastRow As Integer
Dim TotalInstances As Integer

    If SkipCheck = True Then Exit Sub

HomeSheetName = ActiveSheet.Name

SetSymbol
If SearchForSymbol = "" Then Exit Sub

    If Me.RefTargetRng.Value = "" Then
    FirstRow = Selection.Row
    Else
    FirstRow = ExtractAddressElement(Me.RefTargetRng.Value, 2)
    End If

TotalInstances = 0
'loop over all sheets
For s = 1 To Me.lstSheets.ListCount
    If Me.lstSheets.Selected(s - 1) = True Then
        ActiveWorkbook.Sheets(s).Activate
        TotalInstances = TotalInstances + CountMarkers(SearchForSymbol)
    End If
Next s

Me.lblCount.Visible = True
Me.lblCount.Caption = "Count: " & CStr(TotalInstances)
Sheets(HomeSheetName).Activate
'select range
    If TotalInstances > 1 Then
    LastRow = FirstRow + TotalInstances - 1
        If Me.chkCreateHeading.Value = True Then LastRow = LastRow + 1
    Range(Cells(FirstRow, T_Description), Cells(LastRow, T_LossGainEnd)).Select
    End If
End Sub

Sub SetSymbol()

If Me.optLouvre.Value = True Then
    SearchForSymbol = ChrW(T_MrkLouvre)
    SearchGroupName = "Louvre"
ElseIf Me.optSilencer.Value = True Then
    SearchForSymbol = ChrW(T_MrkSilencer)
    SearchGroupName = "Silencer"
ElseIf Me.optResult.Value = True Then
    SearchGroupName = "Key Element"
    SearchForSymbol = ChrW(T_MrkResult)
Else
    Me.lblCount.Visible = True
    Me.lblCount.Caption = "No symbol selected!"
End If

End Sub

Private Sub btnOK_Click()

Dim StartRw As Integer

If Me.RefTargetRng.Value = "" Then
    MsgBox "No destination range selected.", vbOKOnly, "Error - Destination Range"
    Exit Sub
Else
    StartRw = ExtractAddressElement(Me.RefTargetRng.Value, 2)
End If

SetSymbol

    'set heading
    If Me.chkCreateHeading.Value = True Then
    SetDescription SearchGroupName & " Schedule", StartRw, True
    SetTraceStyle ("Title")
    StartRw = StartRw + 1
    End If

HomeSheetName = ActiveSheet.Name '<- ok for now but may need to update to get from destination sheet name

BuildSchedule StartRw

btnOkPressed = True
Me.Hide
Unload Me
End Sub

Private Sub btnSelectAll_Click()
Dim i As Integer
SkipCheck = True
For i = 0 To Me.lstSheets.ListCount - 1 'stupid zero index
    Me.lstSheets.Selected(i) = True
Next i
SkipCheck = False
CountMarkersAllSheets
End Sub

Private Sub btnSelectNone_Click()
Dim i As Integer
SkipCheck = True
For i = 0 To Me.lstSheets.ListCount - 1 'stupid zero index
    Me.lstSheets.Selected(i) = False
Next i
SkipCheck = False
CountMarkersAllSheets
End Sub

Private Sub lstSheets_Change()
CountMarkersAllSheets
End Sub

Private Sub optLouvre_Click()
CountMarkersAllSheets
End Sub

Private Sub optResult_Click()
CountMarkersAllSheets
End Sub

Private Sub optSilencer_Click()
CountMarkersAllSheets
End Sub

Private Sub UserForm_Activate()
Dim s As Integer
btnOkPressed = False
    With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

'list all sheets in listbox
For s = 1 To ActiveWorkbook.Sheets.Count
    Me.lstSheets.AddItem ActiveWorkbook.Sheets(s).Name
Next s
SearchForSymbol = ""
End Sub

'==============================================================================
' Name:     BuildSchedule
' Author:   PS
' Desc:     Counts how many of the selected symbol there is in the sheets
' Args:     WriteRw -  first row of range
' Comments: (1)
'==============================================================================
Sub BuildSchedule(WriteRw As Integer)
Dim overflow As Integer
Dim c As Integer
Dim ShNm As String
Dim ScanRw As Integer
Dim TargetLGCol As Integer
Dim TargetDescCol As Integer

'loop over all sheets
For s = 1 To Me.lstSheets.ListCount
    If Me.lstSheets.Selected(s - 1) = True Then
    Sheets(s).Activate
    ShNm = "'" & Sheets(s).Name & "'!"
    overflow = 0
    ScanRw = 8
    SetSheetTypeControls ScanRw 'from the first line
    
        While overflow < 100 'exit on 100 blank rows
            If Cells(ScanRw, T_Description).Value = "" And Cells(ScanRw, 1).Value = "" Then
                overflow = overflow + 1
            Else
                overflow = 0 'reset to 0
                If Cells(ScanRw, 1).Value = SearchForSymbol Then
                
                TargetLGCol = T_LossGainStart
                TargetDescCol = T_Description
                
                Sheets(HomeSheetName).Activate
                SetSheetTypeControls ScanRw
                
                'insert references
                Cells(WriteRw, T_Description).Select 'move top row so ExtendFunction works
                Cells(WriteRw, T_Description).Value = "=" & ShNm & _
                    Cells(ScanRw, TargetDescCol).Address(False, False)
                Cells(WriteRw, T_LossGainStart).Value = "=" & ShNm & _
                    Cells(ScanRw, TargetLGCol).Address(True, False)
                ExtendFunction
                
                    'style/marker/comment
                    If Me.cBoxApplyStyle.Value <> "" Then
                    SetTraceStyle Me.cBoxApplyStyle.Value
                    End If
                ApplyTraceMarker "Schedule"
                InsertComment ShNm, T_Description, False
                
                'go back to sheet
                Sheets(s).Activate
                WriteRw = WriteRw + 1
                End If
            End If
        ScanRw = ScanRw + 1
        Wend 'end of loop: overflow
    End If
Next s
Sheets(HomeSheetName).Activate
End Sub

'==============================================================================
' Name:     CountMarkers
' Author:   PS
' Desc:     Counts how many of the selected marker symbol there is in the sheets
' Args:     Symbol - what to look for
' Comments: (1) Uses an overflow to detect the end of calcs
'==============================================================================
Function CountMarkers(Symbol As String)
Dim overflow As Integer
Dim c As Integer
Dim rw As Integer

c = 0
rw = 7

While overflow < 100 'exit on 100 blank rows
    If Cells(rw, T_Description).Value = "" And Cells(rw, 1).Value = "" Then
        overflow = overflow + 1
    Else
        overflow = 0 'reset to 0
        If Cells(rw, 1).Value = Symbol Then
            c = c + 1
        End If
    End If
rw = rw + 1
Wend

CountMarkers = c

End Function

