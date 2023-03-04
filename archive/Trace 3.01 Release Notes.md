# Trace 3.01
Minor update, with bugfixes and small UI improvements
## Updates
- AS2670 form now presents figures for Accel/Vel
- AS2670 now previews the values within the form
- Plot Tool maintains all marker types
- Conversion to octaves now works better, including on CVT sheet
- FixLegacyFunctions skips GetSheetControls, can be called any time
## Bugfixes
- Plot Tool edits an existing chart when selected (again)
- AS2670 curves now starts from 1Hz, was off by 3 bands
- FixLegacyFunctions now includes GetASHRAE - old old old function
- Catch error when inserting new sheet, if nothing is selected

## Changes in progress
- Developed Elbow Regen (some), getting there
- 'Rate curve' field added to AS2670 (not integrated yet)

## Future developments
- Imperial units -  to be developed in consultation with the WSP USA - this has now been kicked off with the USA
- Centralised databases for:
	- SWL 
	- Transmission Loss
	- Alpha
	- Ceiling IL
- More Regenerated Noise functions including:
	- Elbow/Bend
	- Junction

[[Trace]]