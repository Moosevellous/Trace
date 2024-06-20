Attribute VB_Name = "DataAndImport"
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
   
If Err <> 0 Then
   MsgBox "Data on clipboard is not text or is empty"
   End
End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Database_SWL()
ErrorDoesNotExist
End Sub

Sub Database_TL()
ErrorDoesNotExist
End Sub

Sub Database_alpha()
ErrorDoesNotExist
End Sub

Sub CeilingIL()
ErrorDoesNotExist
End Sub


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
Dim HomeWorkbookName As String
Dim FanType As String
Dim FindRw As Integer
Dim FoundSWL As Boolean
Dim PercentDone As Single
Dim PasteCol As Integer
Dim WriteRw As Integer
Dim WriteCol As Integer

If Left(T_BandType, 2) = "to" Then 'TO or TOA
    ErrorOctOnly
    msg = MsgBox("Fantech import is not available for one-third octave band sheets", _
    vbOKOnly, "Import: Impossible")
    End
End If

Application.DefaultFilePath = ActiveWorkbook.Path
HomeWorkbookName = ActiveWorkbook.Name
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

'Open Raw File
File = Application.GetOpenFilename(Filefilter:="Excel Files (*.xlsx),*.xlsx", _
ButtonText:="Please select file (in XLSX format)...", _
MultiSelect:=True)

    If Not IsArray(File) Then End
    
Application.ScreenUpdating = False
numFiles = UBound(File)
frmLoading.Show (False)
frmLoading.lblStartTime.Caption = Now

For fnumber = 1 To numFiles
    'open file and get ready
    Workbooks.Open File(fnumber)
    DoEvents
    RawBookName = ActiveWorkbook.Name
    frmLoading.lblFileName.Caption = RawBookName
    
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
    Workbooks(HomeWorkbookName).Activate
    
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
    Workbooks(HomeWorkbookName).Activate
    Cells(WriteRw, PasteCol).PasteSpecial Paste:=xlValues, Transpose:=True
    SetDescription FanType & " - Outlet", WriteRw
    CreateSparkline Selection.Row, 0
    
    'clean up
    Application.CutCopyMode = False
    Workbooks(RawBookName).Close (False)
    
    'Status
    doneFiles = doneFiles + 1
    PercentDone = (doneFiles / numFiles)
    frmLoading.lblStatus.Caption = "(" & doneFiles & "/" & numFiles & ") " & _
        Round(PercentDone * 100, 1) & "%"
    WriteRw = WriteRw + 1
Next fnumber
    
frmLoading.Hide
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
Dim RngRw As String
Dim i As Integer

    If T_BandType = "oct" Then ErrorThirdOctOnly
    
ParameterUnmerge (Selection.Row)

ClipData = GetClipBoardText
'Debug.Print ClipData

    'catch ZORBA Import
    If InStr(1, ClipData, "NRC", vbTextCompare) > 0 Then
    msg = MsgBox("This looks like ZORBA data. Did you mean the other button?", _
        vbOKOnly, "Error - Data mismatch")
    End
    End If

SplitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)

PasteCol = T_Description

    For i = 0 To 21 'stop at 5k band
    Debug.Print SplitData(i)
    'split on tab
    splitLine = Split(SplitData(i), vbTab, Len(SplitData(i)), vbTextCompare)
    Cells(Selection.Row, PasteCol).Value = splitLine(UBound(splitLine))
        If PasteCol = 2 Then 'construction description is here
        TitleStr = splitLine(UBound(splitLine))
        PasteCol = FindFrequencyBand("50") - 1 'skip to column D, then +1 to E later
        ElseIf PasteCol = 25 Then 'put in the ratings
        
        RngRw = Range(Cells(Selection.Row, FindFrequencyBand("100")), _
            Cells(Selection.Row, FindFrequencyBand("3.15k"))).Address(False, True)
            
            If InStr(1, TitleStr, "FLOOR", vbTextCompare) = 0 Then 'not floor
            'Rw
            Cells(Selection.Row, T_ParamStart).Value = "=RwRate(" & RngRw & ")"
            Cells(Selection.Row, T_ParamStart).NumberFormat = """Rw ""0"
            'Ctr
            Cells(Selection.Row, T_ParamStart + 1).Value = "=CtrRate(" & RngRw & _
                "," & T_ParamRng(0) & ")"
            Cells(Selection.Row, T_ParamStart + 1).NumberFormat = ";Ct\r -0;"
            End If
            
        End If
        PasteCol = PasteCol + 1
    Next i
'sparkline
CreateSparkline Selection.Row, 0

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
Dim i As Integer

    If T_BandType = "oct" Then ErrorThirdOctOnly
    
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

SplitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
PasteCol = FindFrequencyBand("31.5")

    For i = 0 To UBound(SplitData) 'skip last two lines
    'split on tab
    splitLine = Split(SplitData(i), vbTab, Len(SplitData(i)), vbTextCompare)
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
    
End Sub


'==============================================================================
' Name:     ConvertNGARAcsv2stats
' Author:   PS
' Desc:     Takes raw NGARA data and converts into statistics at the desired interval
' Args:
' Comments: (1)
'==============================================================================
Sub ConvertNGARAcsv2stats()
Dim folderPath As String
Dim csvFileName As String
Dim csvFilePath As String
Dim textFile As Object
Dim textStream As Object
Dim lineText As String
Dim lineNumber As Integer

' Prompt the user to select a folder
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a NGARA session"
    .Show
    If .SelectedItems.Count <> 0 Then
        folderPath = .SelectedItems(1) & "\"
    Else
        MsgBox "No folder selected. Exiting the macro.", vbExclamation
        Exit Sub
    End If
End With

' Check if the selected folder contains CSV files
If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
If Dir(folderPath & "*.csv") = "" Then
    MsgBox "No CSV files found in the selected folder. Exiting the macro.", vbExclamation
    Exit Sub
End If

' Loop through all CSV files in the selected folder
csvFileName = Dir(folderPath & "*.csv")
Do While csvFileName <> ""
    csvFilePath = folderPath & csvFileName
    
    ' Open the CSV file as a text file
    Set textFile = CreateObject("Scripting.FileSystemObject")
    Set textStream = textFile.OpenTextFile(csvFilePath)
    
    ' Loop through each line in the text file
    lineNumber = 1
    Do While Not textStream.AtEndOfStream
        lineText = textStream.ReadLine
        
        ' Place your code to process each line here
        ' Example: Debug.Print "Line " & lineNumber & ": " & lineText
        
        lineNumber = lineNumber + 1
    Loop
    
    ' Close the text file
    textStream.Close
    
    ' Get the next CSV file in the folder
    csvFileName = Dir
Loop

MsgBox "All CSV files in the folder have been processed as text files.", vbInformation
End Sub
