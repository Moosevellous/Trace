Attribute VB_Name = "Working"


Sub Print_Unicode_Table()
Offset = 0
rw = 2

For x = 1 To 65535
    If rw >= 102 Then
        Offset = Offset + 3
        rw = 2
        'Range(Offset).ColumnWidth = 2
    End If
    Cells(rw, 1 + Offset).Value = x
    Cells(rw, 2 + Offset).Value = ChrW(x)
    Application.StatusBar = x & " " & ChrW(x)
    rw = rw + 1
Next x

Application.StatusBar = False

End Sub

