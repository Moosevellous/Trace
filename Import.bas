Attribute VB_Name = "Import"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     GetClipBoardText
' Author:   PS
' Desc:     Stores clipboard data to string
' Args:     None
' Comments: (1)
'==============================================================================
Function GetClipBoardText() As String
Dim DataObj As MSForms.DataObject
Set DataObj = New MSForms.DataObject

On Error GoTo Whoa

'~~> Get data from the clipboard.
DataObj.GetFromClipboard

'~~> Get clipboard contents, pass to output of function
GetClipBoardText = DataObj.GetText(1)

Exit Function
Whoa:
   If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     ImportFantech
' Author:   PS
' Desc:     Imports Sound Power Data from Fantech export
' Args:     None
' Comments: (1) Assumes files are exported as .xlsx format
'==============================================================================
Sub ImportFantech()

On Error GoTo errHandler

Dim SplitStr() As String
Dim RawBookName As String
Dim WriteRw As Integer
Dim PercentDone As Single
Dim PasteCol As Integer
Dim FindRw As Integer
Dim FoundSWL As Boolean

    If Left(T_BandType, 2) = "TO" Then 'TO or TOA
    ErrorOctOnly
    msg = MsgBox("Fantech import is not available for one-third octave band sheets", _
    vbOKOnly, "Import: Impossible")
    End
    End If

Application.DefaultFilePath = ActiveWorkbook.Path
WorkbookName = ActiveWorkbook.Name
WriteRw = Selection.Row
WriteCol = 2
doneFiles = 0
runonce = False

    'select where to put sound power
    If T_SheetType = "MECH" Then
    PasteCol = T_RegenStart + 1
    Else
    PasteCol = T_LossGainStart + 1
    End If

Application.ScreenUpdating = False

'Open Raw File
File = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx),*.xlsx", _
ButtonText:="Please select file (in XLSX format)...", _
MultiSelect:=True)

    If Not IsArray(File) Then End

numFiles = UBound(File)

    For fnumber = 1 To UBound(File)
    'open file and get ready
    Workbooks.Open File(fnumber)
    DoEvents
    RawBookName = ActiveWorkbook.Name
    Application.StatusBar = RawBookName
    
    'Inlet
    FanType = Cells(7, 2).Value
    FindRw = 3
    FoundSWL = False
        While FoundSWL = False
            If InStr(1, Cells(FindRw, 1).Value, "Sound Power Inlet", vbTextCompare) > 0 Then
            FoundSWL = True
            Else
            FindRw = FindRw + 1
            End If
        Wend
    Range(Cells(FindRw, 2), Cells(FindRw + 7, 2)).Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, PasteCol).PasteSpecial Paste:=xlValues, Transpose:=True
    SetDescription FanType & " - Inlet", WriteRw
    CreateSparkline Selection.Row, 0
    WriteRw = WriteRw + 1
    
    'Outlet
    Workbooks(RawBookName).Activate
    FoundSWL = False
        While FoundSWL = False
            If InStr(1, Cells(FindRw, 1).Value, "Sound Power Outlet", vbTextCompare) > 0 Then
            FoundSWL = True
            Else
            FindRw = FindRw + 1
            End If
        Wend
    Range(Cells(FindRw, 2), Cells(FindRw + 7, 2)).Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, PasteCol).PasteSpecial Paste:=xlValues, Transpose:=True
    SetDescription FanType & " - Outlet", WriteRw
    CreateSparkline Selection.Row, 0
    
    'clean up
    Application.CutCopyMode = False
    Workbooks(RawBookName).Close (False)
    
    'Status
    doneFiles = doneFiles + 1
    PercentDone = (doneFiles / numFiles)
    WriteRw = WriteRw + 1
    Next fnumber
    
Application.StatusBar = False
'msg = MsgBox("Import Complete. " & fnumber - 1 & " files imported.", _
vbOKOnly, "Fantech Import")

Exit Sub

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly, "Error"

End Sub

'==============================================================================
' Name:     ImportInsul
' Author:   PS
' Desc:     Imports TL data from the clipboard
' Args:     None
' Comments: (1) Assumes data has been copied to the clipboard from INSUL software
'==============================================================================
Sub ImportInsul()
Dim PasteCol As Integer
Dim ClipData As String

    If Left(T_SheetType, 2) = "TO" Then 'TO or TOA
    ParameterUnmerge (Selection.Row)
    
    ClipData = GetClipBoardText
    'Debug.Print ClipData
        'catch ZORBA Import
        If InStr(1, ClipData, "NRC", vbTextCompare) > 0 Then
        msg = MsgBox("This looks like ZORBA data. Did you mean the other button?", _
            vbOKOnly, "Error - Data mismatch")
        End
        End If
    
    splitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
    
    PasteCol = 2 '<--TODO make this flexible based on Sheet Type
    
        For i = 0 To 21 'stop at 5k band
        'split on tab
        splitLine = Split(splitData(i), vbTab, Len(splitData(i)), vbTextCompare)
        Cells(Selection.Row, PasteCol).Value = splitLine(UBound(splitLine))
            If PasteCol = 2 Then 'construction description is here
            TitleStr = splitLine(UBound(splitLine))
            PasteCol = PasteCol + 2 'skip to column D, then +1 to E later
            ElseIf PasteCol = 25 Then 'put in the ratings
                If InStr(1, TitleStr, "FLOOR", vbTextCompare) = 0 Then 'not floor
                Cells(Selection.Row, T_ParamStart).Value = "=RwRate(H" & Selection.Row & _
                    ":W" & Selection.Row & ")" 'Rw
                Cells(Selection.Row, T_ParamStart).NumberFormat = """Rw ""0"
                Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(H" & Selection.Row & _
                    ":W" & Selection.Row & ",Z" & Selection.Row & ")" 'Ctr
                Cells(Selection.Row, T_ParamStart + 1).NumberFormat = ";Ct\r -0;"
                End If
            End If
            PasteCol = PasteCol + 1
        Next i
    'sparkline
    CreateSparkline Selection.Row, 0
    Else
    ErrorThirdOctOnly
    End If
End Sub

'==============================================================================
' Name:     ImportZorba
' Author:   PS
' Desc:     Imports absorption data from the clipboard
' Args:     None
' Comments: (1) Assumes data has been copied to the clipboard from Zorba Software
'==============================================================================
Sub ImportZorba()
Dim PasteCol As Integer
Dim ClipData As String

    If Left(T_SheetType, 2) = "TO" Then 'TO or TOA
    ClipData = GetClipBoardText
    
        'catch INSUL Import
        If InStr(1, ClipData, "Wall", vbTextCompare) > 0 Or _
        InStr(1, ClipData, "Floor", vbTextCompare) > 0 Or _
        InStr(1, ClipData, "Ceiling", vbTextCompare) > 0 Or _
        InStr(1, ClipData, "Roof", vbTextCompare) > 0 Or _
        InStr(1, ClipData, "Glazing", vbTextCompare) > 0 Or _
        InStr(1, ClipData, "Porous", vbTextCompare) > 0 Then
        msg = MsgBox("This looks like INSUL data. Did you mean the other button?", _
            vbOKOnly, "Error - Data mismatch")
        End
        End If
    
    splitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
    PasteCol = 5 '<--TODO make this flexible based on Sheet Type
    
        For i = 0 To UBound(splitData) 'skip last two lines
        'split on tab
        splitLine = Split(splitData(i), vbTab, Len(splitData(i)), vbTextCompare)
            If i <= 21 Then 'first 21 rows contain data
            Cells(Selection.Row, PasteCol).Value = splitLine(UBound(splitLine)) 'last element of line
            PasteCol = PasteCol + 1
            ElseIf i = 22 Then 'NRC value
            SetDescription "Import from ZORBA - NRC " & splitLine(UBound(splitLine))
            Else
            'do nothing
            End If
        Next i
        
    'single value rating: NRC
    ParameterMerge (Selection.Row)
    Cells(Selection.Row, T_ParamStart).NumberFormat = """NRC ""0.00"
    'sparkline
    CreateSparkline Selection.Row, 0
    Else
    ErrorThirdOctOnly
    End If
    
End Sub
