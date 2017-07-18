# Trace #
An open source calculation toolkit for Acoustics 
### 1. Introduction ###
Trace is a custom designed Add-In for Excel to assist in the calculation of noise levels, particularly for mechanical sources. The methods used generally implement ASHRAE methodology, version 2011. 

### 2. Installation (.xlam file) ### 
To install in Excel go to `File>Options>Add-ins.` Then click on `Go…`
Then browse to where the NoiseCalc folder has been saved, eg.
Z:\Specialists\Acoustics\1 - Technical Library\Excel Add-in\NoiseCalc
And add the file `NoiseCalc.xlam`

**Optional - Central Locations**
For business groups, the toolkit may be placed centrally on the network with all users pointing to it. If this is the case, then when prompted as follows:

`Copy Trace.xlam to the AddIns folder for USERNAME? `

Click `No`

Updates to the central AddIn file can then be rolled out to all users automatically.


### 3. Design Principles ###
The Add-In can be used for a variety of calculations, mainly centred around fundamental acoustic equations. Where possible, the full equation is written out, except where:

- Empirical relationships exist / lookup tables are required. 
- The formula is difficult to read as an inline function.

Other design principles of the toolkit are:
- All losses are to be inserted as negative numbers. In general the final formula should be the sum of all lines of the calculation. 
- Sheets are designed to be printable as an Appendix to a report. Not all lines of the calculation are displayed in the final presented sheet, but should be sufficient for a review. 
- User input cells is to be ‘post-it’ yellow `RGB (251, 251, 143)`.
- Final answer is be in ‘friendly’ blue `RGB (146, 205, 220)`.
- Units (10 m<sup>2</sup>,Q=2) are presented as custom formatting with the cell.

#### Blank Calculation Sheets ####
Blank Calculation Sheets are laid out for the purposes of setting up calculations from scratch. There are cur-rently four such sheets:
- Octave Band (Linear Input) – 31.5Hz to 8kHz bands
- Octave Band (A-weighted input) – 31.5Hz to 8kHz bands
- One-Third Octave Band (Linear Input) – 50Hz to 5kHz bands 
- One-Third Octave Band (A-weighted Input) – 50Hz to 5kHz bands

There are no formulas in Blank Calculation Sheets, except for a running linear and A-weighted total in each line. Users can utilise the in-built functions to undertake a calculation – this may be supplemented by any equation in the usual manner. Additional sheets are added to the current workbook. The named range TYPECODE denotes the type of Blank Calculation Sheet – this is fed into other logic structures within the code.

####	Standard Calculation Sheets ####
These sheets are set up to calculate common acoustic problems. The ‘Standard Calculation’ button on the ribbon scans the directory (Trace/Standard Calc Sheets) and populates a list in a form. The user selects from the list and is prompted to save a copy, which is date stamped by default.

As with Blank Calculation Sheets, the Standard Calculation Sheets are laid out to fit a on a page when printed. Not all lines of the calculation are displayed in the final presented sheet, but should be sufficient for a review. Table 1 shows a list of the Sheets that have been developed to date. Other sheets can be developed and shared within an organisation by simply copying them into the appropriate folder, however sharing calculation sheets to the entire development pool for common acoustic calculations is strongly encouraged. 

**Table 1: Standard Calculation Sheets (Currently available)** 

| CODE | Name                        | Description                                                                                                                   | Latest Version |
|------|-----------------------------|-------------------------------------------------------------------------------------------------------------------------------|----------------|
| BA   | Barrier Attenuation         | Implements Maekawa's formula to predict barrier insertion loss                                                                | 1.3            |
| GLZ  | Noise Ingress Glazing       | Predicts internal noise level of external noise source through glazing in one-third octave bands                              | 1.2            |
| LG   | Logger Grapher              | Presents ARL316 logger data                                                                                                   | 1.1            |
| N1L  | SEPP N-1 Limits             | Determines Noise Limits under State Environment Protection Policy No N-1 - Control of Noise from commerce, industry and trade | 1.1            |
| NR1L | NIRV Limits                 | Determines Recommended Maximum Noise Levels (RMNLs) under EPA Publication 1411 - Noise in Regional Victoria                   | 1.0            |
| P2W  | SWL to SPL Conversion       | Converts SPL to SWL for up to four sources. Converts SWL to SPL based on RT of room                                           | 1.0            |
| R2R  | Room to Room Transmission   | Predicts noise levels in a room, transmitted through a partition from another room                                            | 1.0            |
| RT   | RT Calculation              | Estimates Reverberation Time using 3 methods based on room geometry and finishes.                                             | 1.6            |
| SIA  | Sound Insulation - Airborne | Calculates Dw and DnTw for measured wall/floor systems                                                                        | 1.4            |
