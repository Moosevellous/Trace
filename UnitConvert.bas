Attribute VB_Name = "UnitConvert"
'Functions to be written:

Function Inches2mm(inputValue As Variant)
Inches2mm = inputValue * 25.4
End Function

Function mm2Inches(inputValue As Variant)
mm2Inches = inputValue / 25.4
End Function

Function Metres2Feet(inputValue As Variant)
Metres2Feet = inputValue * 3.28084
End Function

Function Feet2Metres(inputValue As Variant)
Feet2Metres = inputValue / 3.28084
End Function

'FORWARDING FUNCTIONS
Function DuctAtten_ASHRAE_IU(freq As String, DuctHeight As Long, DuctWidth As Long, DuctType As String, length As Double)
Dim H_metric As Long
Dim W_metric As Long
Dim Length_metric As Double

'convert units
H_metric = Inches2mm(DuctHeight)
W_metric = Inches2mm(DuctWidth)
Length_metric = Metres2Feet(length)

'call metric function
DuctAtten_ASHRAE_IU = DuctAtten_ASHRAE(freq, H_metric, W_metric, DuctType, Length_metric)

End Function


