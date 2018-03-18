Attribute VB_Name = "RibbonControls"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'LOAD
Sub btnLoad(control As IRibbonControl)
On Error Resume Next
New_Tab
End Sub

Sub btnSameType(control As IRibbonControl)
On Error Resume Next
Same_Type
End Sub

Sub btnStandardCalc(control As IRibbonControl)
On Error Resume Next
StandardCalc
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IMPORT FROM
Sub btnImportFantech(control As IRibbonControl)
On Error Resume Next
IMPORT_FANTECH_DATA (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnImportInsul(control As IRibbonControl)
On Error Resume Next
Import_INSUL (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ROW FUNCTIONS

Sub btnClearRw(control As IRibbonControl)
On Error Resume Next
ClearRw (ActiveSheet.Range("TYPECODE").Value)
End Sub

'Sub btnAWeight(control As IRibbonControl)
'On Error Resume Next
'A_weight_oct (ActiveSheet.Range("TYPECODE").Value)
'End Sub

Sub btnFlipSign(control As IRibbonControl)
On Error Resume Next
FlipSign (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnMoveUp(control As IRibbonControl)
On Error Resume Next
MoveUp (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnMoveDown(control As IRibbonControl)
On Error Resume Next
MoveDown (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnSingleCorrection(control As IRibbonControl)
On Error Resume Next
SingleCorrection (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnAutoSum(control As IRibbonControl)
On Error Resume Next
AutoSum (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnManual_ExtendFunction(control As IRibbonControl)
On Error Resume Next
Manual_ExtendFunction (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnTenLogN(control As IRibbonControl)
On Error Resume Next
TenLogN (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnTenLogOneOnT(control As IRibbonControl)
On Error Resume Next
TenLogOneOnT (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnOneThirdsToOctave(control As IRibbonControl)
On Error Resume Next
OneThirdsToOctave (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnConvertAWeight(control As IRibbonControl)
On Error Resume Next
ConvertToAWeight (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnRowReference(control As IRibbonControl)
On Error Resume Next
RowReference (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOISE FUNCTIONS

Sub btnAirAbsorption(control As IRibbonControl)
On Error Resume Next
AirAbsorption (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDistance(control As IRibbonControl)
On Error Resume Next
Distance (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDistanceLine(control As IRibbonControl)
On Error Resume Next
DistanceLine (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnAreaCorrection(control As IRibbonControl)
On Error Resume Next
Area (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDuctSplit(control As IRibbonControl)
On Error Resume Next
DuctSplit (ActiveSheet.Range("TYPECODE").Value)
DoEvents
End Sub

Sub btnASHRAE_Duct(control As IRibbonControl)
On Error Resume Next
ASHRAE_DUCT (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnFlexDuct(control As IRibbonControl)
On Error Resume Next
FlexDuct (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnERL(control As IRibbonControl)
On Error Resume Next
ERLoss (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnElbow(control As IRibbonControl)
On Error Resume Next
ElbowLoss (ActiveSheet.Range("TYPECODE").Value)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'ROOM LOSS GROUP
Sub btnRoomLoss(control As IRibbonControl)
On Error Resume Next
RoomLoss (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnRoomLossRC(control As IRibbonControl)
On Error Resume Next
RoomLossRC (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnRoomLossRT(control As IRibbonControl)
On Error Resume Next
RoomLossRT (ActiveSheet.Range("TYPECODE").Value)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub btnRegenNoise(control As IRibbonControl)
On Error Resume Next
RegenNoise (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDirectReverberantSum(control As IRibbonControl)
On Error Resume Next
DirRevSum (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnSilencer(control As IRibbonControl)
On Error Resume Next
'msg = MsgBox("Feature does not exist yet - please try again later", vbOKOnly, "You wanna build it? Go right ahead.")
Silencer (ActiveSheet.Range("TYPECODE").Value)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CURVE FUNCTIONS
Sub btnNRcurve(control As IRibbonControl)
On Error Resume Next
PutNR (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnNCcurve(control As IRibbonControl)
On Error Resume Next
PutNC (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnRwCurve(control As IRibbonControl)
On Error Resume Next
PutRw (ActiveSheet.Range("TYPECODE").Value)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SHEET FUNCTIONS
Sub btnHeaderBlock(control As IRibbonControl)
On Error Resume Next
HeaderBlock (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnClearHeaderBlock(control As IRibbonControl)
On Error Resume Next
ClearHeaderBlock (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnFormatBorders(control As IRibbonControl)
On Error Resume Next
FormatBorders
End Sub

Sub btnPlot(control As IRibbonControl)
On Error Resume Next
Plot (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnHeatMap(control As IRibbonControl)
'On Error Resume Next
'msg = MsgBox("Feature does not exist yet - please try again later", vbOKOnly, "You wanna build it? Go right ahead.")
HeatMap (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnFixReferences(control As IRibbonControl)
On Error Resume Next
FixReferences (ActiveSheet.Range("TYPECODE").Value)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'named range for sheet type would enable error catching
'if named range throws error, can catch at button press
Public Sub ErrorTypeCode()
msg = MsgBox("Function only possible in template sheets.", vbOKOnly, "Waggling finger of shame")
End
End Sub
