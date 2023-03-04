# New features / updates

## Data group
- Renamed to DataAndImport, because that's what it does
- Moved functions into that group

## Curves group
- Added Equal loudness curves as per ISO226
- Added G-weighting curve
- Option to rate Rw and Ctr on the same line (not the next line)

## Vibration
- VC curves now include VC-F, VC-G, and VC-H
- VC curves will now rate the spectrum above, when you click the button (aligning with others like Rw)

## Sheet tools
- Moved buttons from here to Format/Style group
- *New* Schedule Builder - generates a schedule to elements from other tabs

## Format/Style
- Clear Heat Map button
- Markers added in column A of calc sheets. Styles are defined for:
	- Sum
	- Average
	- Silencer
	- Louvre
	- Result
	- Schedule item
## Help group
- Moved things into this group from elsewhere. 
- *New* Feedback button to send an email - to be updated to a central email address (in progress)

# Bugfixes
- Descriptions inserted after the form is called (various)
- Fix Ref button - defaults to only the current sheet and no replacing legacy functions. This option is still in the dropdown but doesn't happen every time by default. 

# Changes in progress
- More engineering approximations
- RoughRT - a tool to insert an approximate RT based on typical alphas
- ASHRAE 2019 method for Duct Attenuation
- SWL approximation for Diesel Engines
- Custom form for Louvre directivity 
- Pressure drop across silencer

# Future developments
-   Centralised databases for:
    -   SWL
    -   Transmission Loss
    -   Alpha
    -   Ceiling IL
-   More Regenerated Noise functions including:
    -   Junction
    -   Grilles
- Vibration levels from Blasting
- indoor barriers
- Patron noise (AAAC guideline)
	-  Rindel Method
	-  Simulation method
	-  Hybrid method
- 'Extend' function to be made more flexible

# Other devs
-   Central email for feedback
	- TraceHelp@wsp.com?  
	- help@trace.com?




[[Trace]]