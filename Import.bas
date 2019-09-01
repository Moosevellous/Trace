Attribute VB_Name = "Import"
Sub IMPORT_FANTECH_DATA(SheetType As String)

On Error GoTo errHandler

Dim SplitStr() As String
Dim RawBookName As String
Dim WriteRw As Integer
Dim PercentDone As Single

    If Left(SheetType, 2) = "TO" Then
    msg = MsgBox("Fantech import is not available for one-third octave band sheets", vbOKOnly, "Import: Impossible")
    End
    End If


Application.DefaultFilePath = ActiveWorkbook.Path
WorkbookName = ActiveWorkbook.Name
WriteRw = Selection.Row
WriteCol = 2
doneFiles = 0
runonce = False

Application.ScreenUpdating = False


'Open Raw File
File = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx),*.xlsx", _
ButtonText:="Please select file (in XLSX format)...", _
MultiSelect:=True)

If Not IsArray(File) Then End

numfiles = UBound(File)

    For fnumber = 1 To UBound(File)
    Workbooks.Open File(fnumber)
    DoEvents
    RawBookName = ActiveWorkbook.Name
    Application.StatusBar = RawBookName
    'Inlet
    FanType = Cells(7, 2).Value
    Range("B33:B40").Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, 2) = FanType & " - Inlet"
    Cells(WriteRw, 6).PasteSpecial Paste:=xlValues, Transpose:=True
    WriteRw = WriteRw + 1
    
    'Outlet
    Workbooks(RawBookName).Activate
    Range("B42:B49").Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, 2) = FanType & " - Outlet"
    Cells(WriteRw, 6).PasteSpecial Paste:=xlValues, Transpose:=True
        
    'Breakout
    'TODO BREAKOUT
    
    Application.CutCopyMode = False
    Workbooks(RawBookName).Close (False)
    
    'Status
    doneFiles = doneFiles + 1
    PercentDone = (doneFiles / numfiles)
    Application.ScreenUpdating = True
    'Cells(2, 11).Value = PercentDone
    Application.ScreenUpdating = False
    WriteRw = WriteRw + 1
    Next fnumber
    
Application.StatusBar = False
msg = MsgBox("Import Complete. " & fnumber - 1 & " files imported.", vbOKOnly, "Fantech Import")

Exit Sub

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.DESCRIPTION, vbOKOnly, "Error"

End Sub

Sub Import_INSUL(SheetType As String)
Dim PasteCol As Integer
Dim ClipData As String
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 2) = "TO" Then
    Call ParameterUnmerge(Selection.Row, SheetType)
    
    ClipData = GetClipBoardText
    splitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
    PasteCol = 2
        For i = 0 To UBound(splitData)
        'split on tab
        splitLine = Split(splitData(i), vbTab, Len(splitData(i)), vbTextCompare)
        Cells(Selection.Row, PasteCol).Value = splitLine(UBound(splitLine))
            If PasteCol = 2 Then
            TitleStr = splitLine(UBound(splitLine))
            PasteCol = 5 'skip straight to colume E
            ElseIf PasteCol = 25 Then
                If InStr(1, TitleStr, "FLOOR", vbTextCompare) = 0 Then 'not floor
                Cells(Selection.Row, 26).Value = "=RwRate(H" & Selection.Row & ":Y" & Selection.Row & ")" 'Rw
                Cells(Selection.Row, 26).NumberFormat = """Rw ""0"
                Cells(Selection.Row, 27).Value = "=CtrRate(H" & Selection.Row & ":Y" & Selection.Row & ",Z" & Selection.Row & ")" 'Ctr
                Cells(Selection.Row, 27).NumberFormat = ";Ct\r -0;"
                End If
            End
            Else
            PasteCol = PasteCol + 1
            End If
        Next i
    Else
    msg = MsgBox("Implemented for one-third octave sheet only. Please try again later.", vbOKOnly, "To be continued...")
    End If
End Sub

Sub Import_Zorba(SheetType As String)
Dim PasteCol As Integer
Dim ClipData As String
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 2) = "TO" Then
    ClipData = GetClipBoardText
    'Debug.Print ClipData
    
    'catch INSUL Import
    If InStr(1, ClipData, "Wall", vbTextCompare) > 0 Or _
    InStr(1, ClipData, "Floor", vbTextCompare) > 0 Or _
    InStr(1, ClipData, "Ceiling", vbTextCompare) > 0 Or _
    InStr(1, ClipData, "Roof", vbTextCompare) > 0 Or _
    InStr(1, ClipData, "Glazing", vbTextCompare) > 0 Or _
    InStr(1, ClipData, "Porous", vbTextCompare) > 0 Then
    msg = MsgBox("This looks like INSUL data. Did you mean the other button?", vbOKOnly, "Error - Data mismatch")
    End
    End If
    
    splitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
    PasteCol = 5
        For i = 0 To UBound(splitData) 'skip last two lines
        'split on tab
        'Debug.Print splitData(i)
        splitLine = Split(splitData(i), vbTab, Len(splitData(i)), vbTextCompare)
            If i <= 21 Then 'first 21 rows contain data
            Cells(Selection.Row, PasteCol).Value = splitLine(UBound(splitLine)) 'last element of line
            PasteCol = PasteCol + 1
            ElseIf i = 22 Then 'NRC value
            Cells(Selection.Row, 2).Value = "Import from ZORBA - NRC " & splitLine(UBound(splitLine))
            Else
            'do nothing
            End If
        Next i

    Else
    msg = MsgBox("Implemented for one-third octave sheet only. Please try again later.", vbOKOnly, "To be continued...")
    End If
End Sub

Function GetClipBoardText()
   Dim DataObj As MSForms.DataObject
   Set DataObj = New MSForms.DataObject

   On Error GoTo Whoa

   '~~> Get data from the clipboard.
   DataObj.GetFromClipboard

   '~~> Get clipboard contents
   GetClipBoardText = DataObj.GetText(1)

   Exit Function
Whoa:
   If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
End Function

