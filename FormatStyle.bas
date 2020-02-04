Attribute VB_Name = "FormatStyle"
Public targetType As String
Public targetRange As String
Public targetLimitValue As Single
Public targetMarginValue As Single
Public targetCompliantValue As Single
Public targetRoundingWholeNumber As Boolean
Public targetLimitColour As Long
Public targetMarginColour As Long
Public targetCompliantColour As Long

'Sub FormatAs_CellReference(SheetType As String)
'Call GetSettings
'    If Left(SheetType, 3) = "OCT" Then
'    Range(Cells(Selection.Row, 2), Cells(Selection.Row, 15)).Font.Color = fmtReference
'    ElseIf Left(SheetType, 2) = "TO" Then
'    Range(Cells(Selection.Row, 2), Cells(Selection.Row, 25)).Font.Color = fmtReference
'    End If
'End Sub

Sub FormatAs_Total() '<-legacy
Call GetSettings
End Sub

Sub FormatAs_UserInput() '<-legacy
Call GetSettings
End Sub

''''''''''''''''''''''''''

Sub fmtTitle(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Title", SheetType, Selection.Row
End Sub

Sub fmtUnmiti(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Unmitigated", SheetType, Selection.Row
End Sub


Sub fmtMiti(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Mitigated", SheetType, Selection.Row
End Sub

Sub fmtSource(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Lw Source", SheetType, Selection.Row
End Sub

Sub fmtSilencer(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Silencer", SheetType, Selection.Row
End Sub

Sub fmtReference(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Reference", SheetType, Selection.Row
End Sub

Sub fmtSubtotal(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Subtotal", SheetType, Selection.Row
End Sub

Sub fmtTotal(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Total", SheetType, Selection.Row
End Sub

Sub fmtUserInput(SheetType As String, Optional isParamCol As Boolean)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Input", SheetType, Selection.Row, isParamCol
End Sub

Sub fmtComment(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Comment", SheetType, Selection.Row
End Sub

Sub fmtNormal(SheetType As String)
CheckRow (Selection.Row)
ApplyTraceStyle "Trace Normal", SheetType, Selection.Row
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'UNITS
'''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Unit_m(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""m"""
End Sub

Sub Unit_m2(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""m" & chr(178) & """"
End Sub

Sub Unit_m2ps(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""m" & chr(178) & "/s"""
End Sub

Sub Unit_m3ps(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""m" & chr(179) & "/s"""
End Sub

Sub Unit_mm(colStart As Integer, Optional colEnd As Integer, Optional numDigits As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""m" & chr(179) & "/s"""
End Sub

Sub Unit_dB(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""dB"""
End Sub

Sub Unit_dBA(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""dBA"""
End Sub

Sub Unit_kW(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""kW"""
End Sub

Sub Unit_Pa(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "0 ""Pa"""
End Sub

Sub Unit_Q(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "Q=0"
End Sub

Sub Unit_Clear(colStart As Integer, Optional colEnd As Integer)
If colEnd = 0 Then colEnd = colStart
Range(Cells(Selection.Row, colStart), Cells(Selection.Row, colEnd)).NumberFormat = "General"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Target(SheetType As String)

Dim numconditions As Integer

CheckRow (Selection.Row)
frmTarget.Show
If btnOkPressed = False Then End

If targetType = "" Then End

    Select Case targetType
    Case Is = "dB"
    targetRange = Cells(Selection.Row, 3).Address
    Case Is = "dBA"
    'Cells(Selection.Row, 4).Value = targetLimitValue
    targetRange = Cells(Selection.Row, 4).Address
    Case Is = "dBC" '<- TODO check for C-weighted range
    targetRange = Cells(Selection.Row, 4).Address
    Case Is = "NR"
    PutNR (SheetType)
    Case Is = "Band"
    '<- TODO band limits
    End Select

Range(targetRange).Select

    If Selection.FormatConditions.Count > 0 Then
    Selection.FormatConditions.Delete
    End If

numconditions = 1

    If targetLimitValue <> 0 Then
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & targetLimitValue
    Selection.FormatConditions(numconditions).Interior.Color = targetLimitColour
    numconditions = numconditions + 1
    End If
    
    If targetCompliantValue <> 0 Then
    'margin format
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:=targetMarginValue, Formula2:=targetLimitValue
    Selection.FormatConditions(numconditions).Interior.Color = targetMarginColour
    numconditions = numconditions + 1
    'compliant format
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=" & targetCompliantValue
    Selection.FormatConditions(numconditions).Interior.Color = targetCompliantColour
    End If

Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 'I think this line is just good housekeeping?

    
End Sub

Sub ImportTraceStyles()
Dim BlankBookName As String
Dim CurrentBookName As String
GetSettings
CurrentBookName = ActiveWorkbook.Name
Application.Workbooks.Open (TEMPLATELOCATION)
BlankBookName = ActiveWorkbook.Name
Workbooks(CurrentBookName).Activate
ActiveWorkbook.Styles.Merge (BlankBookName)
Workbooks(BlankBookName).Close (False)
End Sub

Sub ApplyTraceStyle(StyleName As String, SheetType As String, InputRw As Integer, Optional isParamCol As Boolean)
    If StyleExists(StyleName) = False Then
    askforstyleimport = MsgBox("No styles in this document. Do you want to import?", vbYesNo, "Style(ish)!")
        If askforstyleimport = vbYes Then
        ImportTraceStyles
        Else
        End
        End If
    End If
Range(GetStyleRange(SheetType, InputRw, isParamCol)).Style = StyleName
End Sub


'Sub DeleteOtherStyles()
'  For i = 1 To ActiveWorkbook.Styles.Count
'    Debug.Print ActiveWorkbook.Styles(i).Name
'        If InStr(1, "Trace", ActiveWorkbook.Styles(i).Name, vbTextCompare) = 0 Then
'        ActiveWorkbook.Styles(i).Delete
'        i = i - 1 'don't skip
'        End If
'    Next i
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetStyleRange(SheetType As String, InputRw As Integer, Optional isParamCol As Boolean) As String
    If Left(SheetType, 3) = "OCT" Then
        If isParamCol Then 'isParamCol defaults to false
        GetStyleRange = Range(Cells(InputRw, 14), Cells(Selection.Row, 15)).Address
        Else
        GetStyleRange = Range(Cells(InputRw, 2), Cells(Selection.Row, 13)).Address
        End If
        Cells(Selection.Row, 4).Font.Bold = True
    ElseIf Left(SheetType, 2) = "TO" Then
        If isParamCol Then
        GetStyleRange = Range(Cells(InputRw, 26), Cells(Selection.Row, 27)).Address
        Else
        GetStyleRange = Range(Cells(InputRw, 2), Cells(Selection.Row, 25)).Address
        Cells(Selection.Row, 4).Font.Bold = True
        End If
    ElseIf Left(SheetType, 2) = "LF" Then
        If isParamCol Then
        GetStyleRange = Range(Cells(InputRw, 32), Cells(Selection.Row, 33)).Address
        Else
        GetStyleRange = Range(Cells(InputRw, 2), Cells(Selection.Row, 31)).Address
        Cells(Selection.Row, 4).Font.Bold = True
        End If
    ElseIf SheetType = "CVT" Then 'no parameter columns
        GetStyleRange = Range(Cells(InputRw, 2), Cells(Selection.Row, 44)).Address
        Cells(Selection.Row, 4).Font.Bold = True
    Else
    msg = MsgBox("Sheet Type '" & SheetType & "' is not supported. Try applying the style manually.", vbOKOnly, "Style application error")
    End
    End If
End Function

Function StyleExists(StyleName As String) As Boolean
StyleExists = False
    For i = 1 To ActiveWorkbook.Styles.Count
    'Debug.Print ActiveWorkbook.Styles(i).Name
        If ActiveWorkbook.Styles(i).Name = StyleName Then
        StyleExists = True
        End If
    Next i
End Function



