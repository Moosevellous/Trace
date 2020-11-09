# Trace 3.0 Release Notes

[[Trace]] 3.0 has been a massive effort for the Working Group, with the back-end code rewritten and documented better than ever before. The last release was in January so the list of changes has certainly grown quite long.

## New features
The new release offers many new features, let's take a look at them now. 
### Front-end stuff
#### Layout and user interface
- Load module & ribbon group renamed 'New', streamlined interface
- 'Mechanical' button split out into group on ribbon 
- 'SWL Estimation' grouped into single button
- 'Curves' grouped into single button
- Added *split-button* controls, with a default option for common use and a drop down menu for rarer uses
- More custom forms preview the values before inserting into calc sheet
- Row numbers from blank calc sheets removed (this may be used later for other purposes)
- ***NEW*** Mechanical noise calculation sheets created, with calculations set up for regenerated noise
- ***NEW*** Conversion sheet (CVT) for multiband operations
- ***NEW*** Reviewer field in header block
- ***NEW*** KeyTips for all buttons (Hint: use ALT then X to get started)

#### Acoustics functions
- Regenerated noise predictions for:
	- Dampers
	- Silencers (Fantech and NAP)
- Mechanical noise layouts
- AS2670 evaluation
- Duct Directivity
- Added Fantech Airlay Silencers
- Frequency band cutoff frequencies as per ANSI S1.11
- Lnw and IIC rating functions
- Octave band mode for C<sub>tr</sub> correction
- Fixed ISO9613 A_div formula

#### Other features
- Move left & right buttons
- Regenerated noise style
- Toggle (switches a line into/out of text to evaluate effect)
- DevTools button to enable code export 
- Button to get to Trace Home Directory
### Back-end stuff
- Centralised ribbon event handler codes
- Mechanical module to align with ribbon group
- Header blocks for **every single function**	
- Comments labeling frequency bands within static arrays
- Renamed the functions for consistency eg:
	- GetASHRAEDuct => DuctAtten_ASHRAE
	- GetReynoldsDuctCircular => DuctAttenCircular_Reynolds
	- GetRoomLossRT => RoomLossTypicalRT
	- Syntax = *{Function}_{Method}* 
- *ReplaceLegacyFunctions* - to fix references to the old naming conventions - called when FixReferences is clicked.
- Centralised control variables 
	- T_Description
	- T_LossGainStart
	- T_LossGainEnd
	- T_RegenStart
	- T_RegenEnd
	- T_ParamStart
	- T_ParamEnd
	- T_Comment
	- T_FreqRow
	- T_ParamRng()
	- T_FreqStartRng 
	- T_BandType 
	- T_SheetType
- Functions to set control variables - STRUCTURE:
	- Array(Description, LossGainStart, LossGainEnd, RegenStart, RegenEnd, ParamStart, ParamEnd, Comment, FreqRow)
- Method for setting Data Validation on cells
- Method for setting Comments on cells
- Method for setting Units (as number formatting)

## Future developments
- Imperial units -  to be developed in consultation with the WSP USA.
- Centralised databases for:
	- SWL 
	- Transmission Loss
	- Alpha
	- Ceiling IL
- More Regenerated Noise functions including:
	- Elbow/Bend
	- Junction
