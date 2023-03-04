# Updates
## Load group
- Added dialog for adding new sheets, to improve user messaging
- Added 'Full Spectrum' blank sheet (20Hz-20kHz, one-third octave bands)
## Basics group
- Added CompositeTL button (two TLs only)
## Data group
- Combined Database and Import, reorganised buttons. Operation mode of functions now made clearer.
## Row Operations group
- 'Correction' renamed 'Value
## Noise group
- Added methods for Barrier Attenuation 
	- Kurze Anderson
	- Menounou
- New form for the barrier methods above AND ISO9613 (assumes Agr=0)
- Added Conformal Surface Area (under 10log button)
- Added form for CSA and Parallelepiped methods
## Curves group
- C correction as per ISO717-1
- C and Ctr corrections have their own buttons
- Added alpha_w function for weighted absorption as per ISO11654
- Added button for same
- Added buttons to insert A and C weighting curves. Option for static values or dynamically calculated.

## Sheet tools
- 'Fill Header Block (all)', which loops through all sheets

# Bugfixes
- Barrier sheet - fixed minor error for path 4
- BuildFormula will accept '=' character or insert if needed
- Changed default Plot layout to be more suitable for reports (legend position / size of chart)


# Changes in progress
- More engineering approximations
- RoughRT - a tool to insert an approximate RT based on typical alphas
- 'MarkRowAs' - function for marking column 1 with symbol
- ASHRAE 2019 method for Duct Attenuation


# Future developments
-   Centralised databases for:
    -   SWL
    -   Transmission Loss
    -   Alpha
    -   Ceiling IL
-   More Regenerated Noise functions including:
    -   Junction
    -   Grilles
- Equal loudness curves as per ISO226
- Vibration levels from Blasting
- indoor barriers
- Patron noise (AAAC guideline)
	-  Rindel Method
	-  Simulation method
	-  Hybrid method
- Extend function to be made more flexible
- Fix Ref button - changes to simplify updates

# Other devs
-   Feedback function?
	-   Auto generated email
	-   Form (makes it anonymous)
-   Central email
	- TraceHelp@wsp.com?  
	- help@trace.com?


[[Trace]]
